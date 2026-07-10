"""
main.py — Kaspi receipt processor (GitHub Actions edition).

Flow:
  1. Connect to Gmail over IMAP using an App Password (from env / GitHub Secrets).
  2. Find UNSEEN emails whose subject contains "Kaspi" and that have a .zip.
  3. Unzip (folders preserved). Skip Excel files. OCR PDFs/images.
  4. Kaspi documents  -> "Kaspi" sheet  (Folder, Document, Date, Name, Amount, Receipt No)
     Non-Kaspi docs   -> "No Kaspi" sheet (Folder, Document)
  5. Validation sheet: documents processed vs rows written.
  6. Reply to the ORIGINAL SENDER with the .xlsx attached, then mark the email Seen.

No per-day OCR cap (Tesseract runs locally on the runner), no 6-minute limit,
no credit card. Runs free on a public GitHub repo via a scheduled workflow.
"""

import os
import re
import sys
import ssl
import zipfile
import tempfile
import imaplib
import smtplib
import email
from email.message import EmailMessage
from email.header import decode_header, make_header
from email.utils import parseaddr
from datetime import datetime

import openpyxl
from openpyxl.styles import Font

from parser import parse_document

# ---- Config (via environment / GitHub Secrets) -----------------------------
GMAIL_USER = os.environ["GMAIL_USER"]           # e.g. automatixmodern@gmail.com
GMAIL_APP_PASSWORD = os.environ["GMAIL_APP_PASSWORD"]   # 16-char app password
SUBJECT_KEYWORD = os.environ.get("SUBJECT_KEYWORD", "Kaspi")
CONF_THRESHOLD = float(os.environ.get("CONF_THRESHOLD", "0.5"))  # flag fields below this

EXCLUDE_EXT = {"xls", "xlsx", "xlsm", "xltx", "csv"}
IMAGE_EXT = {"jpg", "jpeg", "jfif", "png", "gif", "bmp", "webp", "tif", "tiff"}
PDF_EXT = {"pdf"}

IMAP_HOST, IMAP_PORT = "imap.gmail.com", 993
SMTP_HOST, SMTP_PORT = "smtp.gmail.com", 465


def _decode(s):
    try:
        return str(make_header(decode_header(s)))
    except Exception:
        return s or ""


def is_junk(name):
    return ("__MACOSX" in name) or re.search(r'(^|/)\.', name) or name.endswith("/")


# ---- Excel builder ---------------------------------------------------------
def build_excel(kaspi_rows, nokaspi_rows, total_docs, excluded, out_path):
    wb = openpyxl.Workbook()
    bold = Font(bold=True)

    s1 = wb.active
    s1.title = "Kaspi"
    s1.append(["Folder", "Document", "Date", "Name", "Amount", "Receipt Number"])
    for c in s1[1]:
        c.font = bold
    for row in kaspi_rows:
        s1.append(row)
    # keep long receipt numbers + dates as text (no scientific notation / auto-dates)
    for r in range(2, s1.max_row + 1):
        s1.cell(r, 3).number_format = "@"
        s1.cell(r, 6).number_format = "@"

    s2 = wb.create_sheet("No Kaspi")
    s2.append(["Folder", "Document"])
    for c in s2[1]:
        c.font = bold
    for row in nokaspi_rows:
        s2.append(row)

    rows_written = len(kaspi_rows) + len(nokaspi_rows)
    s3 = wb.create_sheet("Validation")
    for row in [
        ["Documents in folders (Excel excluded)", total_docs],
        ["Excel files excluded", excluded],
        ['Rows in "Kaspi" tab', len(kaspi_rows)],
        ['Rows in "No Kaspi" tab', len(nokaspi_rows)],
        ["Validation result", "PASSED" if total_docs == rows_written else "MISMATCH — CHECK"],
    ]:
        s3.append(row)
    for c in s3["A"]:
        c.font = bold

    wb.save(out_path)
    return rows_written


# ---- Process one email -----------------------------------------------------
def process_message(raw_bytes):
    """Returns (result_dict) or None if the message has no zip to act on."""
    msg = email.message_from_bytes(raw_bytes)
    sender_name, sender_addr = parseaddr(msg.get("From", ""))
    subject = _decode(msg.get("Subject", ""))

    # collect zip attachments
    zips = []
    for part in msg.walk():
        fn = part.get_filename()
        if fn and _decode(fn).lower().endswith(".zip"):
            payload = part.get_payload(decode=True)
            if payload:
                zips.append((_decode(fn), payload))
    if not zips:
        return None

    kaspi_rows, nokaspi_rows, flagged = [], [], []
    total_docs, excluded = 0, 0

    with tempfile.TemporaryDirectory() as tmp:
        for zname, zbytes in zips:
            zpath = os.path.join(tmp, "in.zip")
            with open(zpath, "wb") as f:
                f.write(zbytes)
            try:
                zf = zipfile.ZipFile(zpath)
            except zipfile.BadZipFile:
                continue

            for info in zf.infolist():
                name = info.filename
                if info.is_dir() or is_junk(name):
                    continue
                parts = name.split("/")
                doc = parts[-1]
                folder = "/".join(parts[:-1]) if len(parts) > 1 else "(root)"
                ext = doc.rsplit(".", 1)[-1].lower() if "." in doc else ""

                if ext in EXCLUDE_EXT:
                    excluded += 1
                    continue
                if ext not in IMAGE_EXT and ext not in PDF_EXT:
                    # unknown format — still count it, send to No-Kaspi for visibility
                    total_docs += 1
                    nokaspi_rows.append([folder, doc])
                    continue

                total_docs += 1
                # extract this one file to disk for OCR
                fpath = os.path.join(tmp, "doc." + ext)
                with open(fpath, "wb") as out:
                    out.write(zf.read(info))

                try:
                    fields, conf, has_kaspi = parse_document(fpath, is_pdf=(ext in PDF_EXT))
                except Exception as e:
                    kaspi_rows.append([folder, doc, "", f"ERROR: {e}", "", ""])
                    flagged.append(f"{folder}/{doc}  (could not read: {e})")
                    continue

                if has_kaspi:
                    kaspi_rows.append([folder, doc, fields["date"], fields["name"],
                                       fields["amount"], fields["receiptNo"]])
                    missing = [k for k in ("date", "name", "amount", "receiptNo")
                               if not fields[k] or conf[k] < CONF_THRESHOLD]
                    if missing:
                        flagged.append(f"{folder}/{doc}  (empty/uncertain: {', '.join(missing)})")
                else:
                    nokaspi_rows.append([folder, doc])

    out_xlsx = os.path.join(tempfile.gettempdir(),
                            f"Kaspi_Result_{datetime.now():%Y%m%d_%H%M}.xlsx")
    rows_written = build_excel(kaspi_rows, nokaspi_rows, total_docs, excluded, out_xlsx)

    return {
        "to_addr": sender_addr,
        "to_name": sender_name,
        "subject": subject,
        "msg_id": msg.get("Message-ID", ""),
        "xlsx": out_xlsx,
        "total_docs": total_docs,
        "excluded": excluded,
        "kaspi_n": len(kaspi_rows),
        "nokaspi_n": len(nokaspi_rows),
        "rows_written": rows_written,
        "flagged": flagged,
    }


# ---- Send reply ------------------------------------------------------------
def send_reply(res):
    ok = (res["total_docs"] == res["rows_written"])
    body = (
        "Hello,\n\n"
        "Your Kaspi documents have been processed automatically.\n\n"
        f"Documents found (Excel files excluded): {res['total_docs']}\n"
        f"Excel files excluded: {res['excluded']}\n"
        f'Rows written ("Kaspi" tab): {res["kaspi_n"]}\n'
        f'Rows written ("No Kaspi" tab): {res["nokaspi_n"]}\n'
        f"Validation (documents = rows): {'PASSED' if ok else 'MISMATCH — please check'}\n"
    )
    if res["flagged"]:
        body += ("\nDocuments needing manual review (photo OCR is imperfect on\n"
                 "screen photos — please verify these against the originals):\n - "
                 + "\n - ".join(res["flagged"]) + "\n")
    body += "\nThe result file is attached.\n"

    em = EmailMessage()
    em["From"] = GMAIL_USER
    em["To"] = res["to_addr"]
    em["Subject"] = "Re: " + res["subject"]
    if res["msg_id"]:
        em["In-Reply-To"] = res["msg_id"]
        em["References"] = res["msg_id"]
    em.set_content(body)

    with open(res["xlsx"], "rb") as f:
        em.add_attachment(
            f.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="Kaspi_Result.xlsx",
        )

    ctx = ssl.create_default_context()
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=ctx) as smtp:
        smtp.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        smtp.send_message(em)


# ---- Main ------------------------------------------------------------------
def main():
    ctx = ssl.create_default_context()
    imap = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, ssl_context=ctx)
    imap.login(GMAIL_USER, GMAIL_APP_PASSWORD)
    imap.select("INBOX")

    # unseen emails whose subject contains the keyword
    typ, data = imap.search(None, 'UNSEEN', 'SUBJECT', f'"{SUBJECT_KEYWORD}"')
    ids = data[0].split() if data and data[0] else []
    print(f"Found {len(ids)} candidate email(s).")

    processed = 0
    for num in ids:
        typ, msg_data = imap.fetch(num, "(RFC822)")
        raw = msg_data[0][1]
        try:
            res = process_message(raw)
        except Exception as e:
            print(f"  [error] message {num!r}: {e}", file=sys.stderr)
            continue

        if res is None:
            print(f"  message {num!r}: subject matched but no .zip attachment — leaving unread.")
            continue

        send_reply(res)
        imap.store(num, "+FLAGS", "\\Seen")   # mark done only after a successful reply
        processed += 1
        print(f"  Replied to {res['to_addr']}: {res['total_docs']} docs, "
              f"{res['rows_written']} rows, {len(res['flagged'])} flagged.")

    imap.logout()
    print(f"Done. Processed {processed} email(s).")


if __name__ == "__main__":
    main()
