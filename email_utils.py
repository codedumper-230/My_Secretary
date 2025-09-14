import imaplib
import email
from email.header import decode_header

def connect_to_gmail(username, app_password):
    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    imap.login(username, app_password)
    return imap

def fetch_recent_emails(imap, n=5):
    imap.select("inbox")
    status, messages = imap.search(None, "ALL")
    email_ids = messages[0].split()[-n:]

    emails = []
    for eid in email_ids:
        _, msg_data = imap.fetch(eid, "(RFC822)")
        msg = email.message_from_bytes(msg_data[0][1])

        subject, _ = decode_header(msg["Subject"])[0]
        if isinstance(subject, bytes):
            subject = subject.decode()

        from_ = msg.get("From")
        date_ = msg.get("Date")

        # Extract body (plain text)
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    body = part.get_payload(decode=True).decode(errors="ignore")
                    break
        else:
            body = msg.get_payload(decode=True).decode(errors="ignore")

        emails.append({
            "subject": subject,
            "from": from_,
            "date": date_,
            "body": body[:1000]  # limit to avoid huge payloads
        })

    return emails
