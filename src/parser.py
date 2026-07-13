"""
parser.py — Cloud Vision OCR + field extraction for Kaspi receipts (RU + KZ).

Uses Google Cloud Vision DOCUMENT_TEXT_DETECTION, tuned for dense document text,
which handles photos of screens (glare/angle) far better than a local engine.
One Vision call per document/page = one billable unit (first 1,000/month free,
then ~$1.50 per 1,000).

Images: bytes sent directly. PDFs: each page rasterized to PNG (pdf2image/poppler)
then sent, since synchronous Vision image requests take images, not multi-page PDFs.
"""

import os
import re
import io
import base64
from collections import Counter

import requests
from PIL import Image

AMOUNT_KEYS = ['Зачислено', 'Пополнено', 'Есептелді', 'Толықтырылды', 'Толықтырыл',
               'Перевод успешно совершен', 'Аударым']
RECEIPT_KEYS = ['Номер чека', 'номер чека', '№ квитанции', 'квитанции',
                'Чек нөмірі', 'ек нөмірі', 'Чек не', 'ек не', 'нөмірі']
DATE_KEYS = ['Дата и время', 'Дата', 'Күні', 'Күн', 'Уақыты']
SENDER_KEYS = ['Отправитель', 'Жіберуші']

NAME_RE = re.compile(r'^[A-ZА-ЯЁӘҒҚҢӨҰҮҺІ][a-zа-яёәғқңөұүһі]+\s+[A-ZА-ЯЁӘҒҚҢӨҰҮҺІ]\.$')
DATE_RE = re.compile(r'(\d{2})[.\/-](\d{2})[.\/-](\d{4})')
KASPI_RE = re.compile(r'kaspi|каспи', re.I)

VISION_URL = "https://vision.googleapis.com/v1/images:annotate"
MAX_DIM = 3000   # downscale very large phone photos to keep request size sane


def _maybe_downscale(image_bytes):
    """If the image is very large, downscale it so the API request stays small."""
    try:
        img = Image.open(io.BytesIO(image_bytes))
        w, h = img.size
        if max(w, h) > MAX_DIM:
            scale = MAX_DIM / max(w, h)
            img = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
            buf = io.BytesIO()
            fmt = "PNG" if (img.mode in ("RGBA", "P")) else "JPEG"
            img.convert("RGB").save(buf, format=fmt, quality=90)
            return buf.getvalue()
    except Exception:
        pass
    return image_bytes


def _vision_text(image_bytes):
    """
    One DOCUMENT_TEXT_DETECTION call via the REST endpoint, authenticated with a
    plain API key (works even when service-account keys are disabled by org policy).
    One call = one billable unit.
    """
    api_key = os.environ["GCP_API_KEY"]
    image_bytes = _maybe_downscale(image_bytes)
    payload = {
        "requests": [{
            "image": {"content": base64.b64encode(image_bytes).decode("ascii")},
            "features": [{"type": "DOCUMENT_TEXT_DETECTION"}],
            "imageContext": {"languageHints": ["ru", "kk"]},
        }]
    }
    r = requests.post(f"{VISION_URL}?key={api_key}", json=payload, timeout=60)
    try:
        r.raise_for_status()
    except requests.HTTPError:
        # Never surface the URL (it contains the API key). Report status only.
        raise RuntimeError(f"Vision API HTTP {r.status_code} "
                           f"({'check API enabled + billing + key restrictions' })")
    resp = r.json()["responses"][0]
    if "error" in resp and resp["error"].get("message"):
        raise RuntimeError(f"Vision API error: {resp['error']['message']}")
    return resp.get("fullTextAnnotation", {}).get("text", "")


def ocr_texts(path, is_pdf=False):
    """Return list of OCR text strings (one per image, or one per PDF page)."""
    if is_pdf:
        from pdf2image import convert_from_path
        texts = []
        for page in convert_from_path(path, dpi=300):
            buf = io.BytesIO()
            page.save(buf, format="PNG")
            texts.append(_vision_text(buf.getvalue()))
        return texts
    with open(path, "rb") as f:
        return [_vision_text(f.read())]


def _lines(text):
    return [l.strip() for l in text.splitlines() if l.strip()]


def _clean_amount(raw):
    groups = raw.split()
    if (len(groups) >= 2 and len(groups[-1]) == 1
            and all(len(g) == 3 for g in groups[1:-1]) and 1 <= len(groups[0]) <= 3):
        groups = groups[:-1]
    return re.sub(r'\D', '', ''.join(groups))


def find_amount(lines):
    def stitch_amount(idx):
        """
        Reconstruct the full amount from the keyword line and the lines below it,
        even when Vision split the number across lines (common on tall, glare-
        washed green boxes):
            Есептелді / 15 / 000 ₸   -> 15000
            Есептелді / 15 000 ₸     -> 15000
            Есептелді 15 000 ₸       -> 15000
        """
        groups = []
        # digits already on the keyword line (after the keyword word)
        for tok in re.findall(r'\d+', lines[idx]):
            groups.append(tok)
        j = idx + 1
        while j < len(lines) and j <= idx + 3:
            ln = lines[j]
            if re.search(r'Комисси|Чек|квитанц|Дата|Күн|ТМК|КНП|Kaspi|Устройств|Құрылғ', ln):
                break
            found = re.findall(r'\d+', ln)
            if found:
                groups.extend(found)
                if re.search(r'[₸ТTт]\s*$', ln) or '\uFFFD' in ln:
                    break            # currency glyph => number complete
            elif not re.search(r'[₸ТTт]', ln):
                break                # no digits, no currency => stop
            j += 1
        # drop a stray trailing single digit (misread ₸)
        if len(groups) >= 2 and len(groups[-1]) == 1 and all(len(g) == 3 for g in groups[1:-1]):
            groups = groups[:-1]
        digits = ''.join(groups)
        return digits if len(digits) >= 3 else ''

    for i, l in enumerate(lines):
        for k in AMOUNT_KEYS:
            if k in l:
                d = stitch_amount(i)
                if d:
                    return d
    # fallback: the large grouped number sitting just above the commission line
    ci = next((i for i, l in enumerate(lines) if re.search(r'Комисси', l)), len(lines))
    for l in lines[:ci]:
        m = re.search(r'(\d[\d\s\u00A0]{1,}\d(?:\s+\d)?)\s*[тТtT₸\uFFFD]?\s*$', l)
        if m:
            d = _clean_amount(m.group(1).replace('\u00A0', ' ').strip())
            if len(d) >= 3:
                return d
    return ''


def find_receipt(lines):
    for i, l in enumerate(lines):
        for k in RECEIPT_KEYS:
            if k in l:
                tail = l.split(k, 1)[-1]
                runs = re.findall(r'\d+', tail)
                if runs:
                    # value is on the same line as the label
                    best = max(runs, key=len)
                    if (len(best) >= 10 and i + 1 < len(lines)
                            and re.fullmatch(r'\d[\d\s]*', lines[i + 1])):
                        best += re.sub(r'\D', '', lines[i + 1])   # wrapped continuation
                    if len(best) >= 6:
                        return best
                else:
                    # label sits on its own line (Vision layout): value on next line(s).
                    # Stitch consecutive pure-digit lines together (handles wrapping).
                    digits = ''
                    for j in range(i + 1, min(i + 3, len(lines))):
                        if re.fullmatch(r'\d[\d\s]*', lines[j]):
                            digits += re.sub(r'\D', '', lines[j])
                        elif digits:
                            break
                        else:
                            break
                    if len(digits) >= 6:
                        return digits
    alln = re.findall(r'\d{8,}', '\n'.join(lines))
    return max(alln, key=len) if alln else ''


def find_date(lines):
    for i, l in enumerate(lines):
        if any(k in l for k in DATE_KEYS):
            for j in range(i, min(i + 2, len(lines))):
                m = DATE_RE.search(lines[j])
                if m:
                    return f"{m.group(1)}/{m.group(2)}/{m.group(3)}"
    m = DATE_RE.search('\n'.join(lines))
    return f"{m.group(1)}/{m.group(2)}/{m.group(3)}" if m else ''


def find_name(lines):
    sender = set()
    for i, l in enumerate(lines):
        if any(k in l for k in SENDER_KEYS):
            sender.add(i)
            sender.add(i + 1)
    for j, l in enumerate(lines):
        if j in sender:
            continue
        if NAME_RE.match(l):
            return l
    return ''


def _majority(values):
    vals = [v for v in values if v]
    if not vals:
        return '', 0.0
    counts = Counter(vals)
    top = counts.most_common()
    best_n = top[0][1]
    tied = [v for v, n in top if n == best_n]
    winner = sorted(tied, key=len)[0] if len(tied) > 1 else top[0][0]
    return winner, counts[winner] / len(vals)


def parse_document(path, is_pdf=False):
    texts = ocr_texts(path, is_pdf)
    buckets = {'amount': [], 'receiptNo': [], 'date': [], 'name': []}
    has_kaspi = False
    for t in texts:
        if KASPI_RE.search(t):
            has_kaspi = True
        L = _lines(t)
        buckets['amount'].append(find_amount(L))
        buckets['receiptNo'].append(find_receipt(L))
        buckets['date'].append(find_date(L))
        buckets['name'].append(find_name(L))
    fields, confidence = {}, {}
    for k, v in buckets.items():
        fields[k], confidence[k] = _majority(v)
    return fields, confidence, has_kaspi
