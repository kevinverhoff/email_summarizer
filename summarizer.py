import os
import csv
import win32com.client
import pythoncom
from transformers import T5ForConditionalGeneration, T5Tokenizer, pipeline


VIP_FILE = "vip_list.csv"
SUMMARY_OUTPUT = os.path.expanduser("~/Desktop/email_summary.txt")

### ------------------------------------------------------------
### 1. Load VIP List
### ------------------------------------------------------------
def load_vip_list():
    vip_emails = set()
    with open(VIP_FILE, "r", encoding="utf-8") as f:
        reader = csv.reader(f)
        for row in reader:
            vip_emails.add(row[0].strip().lower())
    return vip_emails


### ------------------------------------------------------------
### 2. Connect to Outlook + Pull Inbox Items
### ------------------------------------------------------------
def connect_classic_outlook():
    pythoncom.CoInitialize()

    # CLSID of Classic Outlook.Application
    CLASSIC_OUTLOOK_CLSID = "{0006F03A-0000-0000-C000-000000000046}"
    return win32com.client.Dispatch(CLASSIC_OUTLOOK_CLSID)

def get_outlook_emails():
    outlook_app = connect_classic_outlook()
    outlook = outlook_app.GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)
    print(f"Found {messages.Count} emails in Inbox")
    return messages


### ------------------------------------------------------------
### 3. Load Local T5 Summarizer
### ------------------------------------------------------------
def load_summarizer():
    tokenizer = T5Tokenizer.from_pretrained("wordcab/t5-small-email-summarizer")
    print("Loaded t5 small email tokenizer")
    model = T5ForConditionalGeneration.from_pretrained("wordcab/t5-small-email-summarizer")
    print("Loaded t5 small email summarizer model")
    return pipeline("summarization", model=model, tokenizer=tokenizer)


### ------------------------------------------------------------
### 4. Summarize a single email
### ------------------------------------------------------------
def summarize_email(summarizer, body):
    if not body:
        return "(No content)"

    note = ""
    truncated_body = body

    # If email body is too long
    if len(body) > 2500:
        note = "Email too long. Showing summary of first 2500 characters. "
        truncated_body = body[:2500]

    summary = summarizer(
        truncated_body,
        min_length=15,
        do_sample=False
    )[0]["summary_text"]

    return note + summary


### ------------------------------------------------------------
### 5. Process + Tag Emails
### ------------------------------------------------------------
def archive_email(msg):
    try:
        outlook_app = connect_classic_outlook()
        outlook = outlook_app.GetNamespace("MAPI")

        # Default Archive folder (Outlook auto-archive folder)
        archive_folder = outlook.GetDefaultFolder(32)  # 32 = Archive

        # Append text to subject before archiving
        try:
            msg.Subject = f"{msg.Subject} ~THIS EMAIL WAS AUTO-ARCHIVED~"
        except:
            pass  # If Outlook blocks subject edits on some meeting items

        # Move the email
        msg.Move(archive_folder)

    except Exception as e:
        print("Error archiving email:", e)

def safe_get(field):
    try:
        return field if field else "missing"
    except:
        return ""

def is_meeting_invite(msg):
    try:
        # 26 is Outlook's MeetingItem enumeration
        if msg.Class == 26:
            return True
        # Many meeting-related items use MessageClass starting with IPM.Schedule
        mc = str(msg.MessageClass)
        if mc.startswith("IPM.Schedule."):
            return True
        return False
    except:
        return False

def process_emails():
    vip_list = load_vip_list()
    messages = get_outlook_emails()
    summarizer = load_summarizer()

    vip_emails = []
    zendesk_emails = []
    meeting_emails = []
    other_emails = []

    for msg in messages:
        try:
            # Meeting invites first — classify and summarize separately
            if is_meeting_invite(msg):
                sender = safe_get(msg.SenderName)
                subject = safe_get(msg.Subject)
                body = safe_get(msg.Body)
                to = safe_get(msg.ReplyRecipients)

                try:
                    if subject.strip().startswith("Accepted"):
                        archive_email(msg)
                except:
                    pass

                summary = summarize_email(summarizer, body)

                meeting_emails.append({
                    "category": "Meeting",
                    "sender": sender,
                    "subject": subject,
                    "to": to,
                    "summary": summary
                })
                continue  # skip normal classification

            # Normal emails
            sender = safe_get(msg.SenderEmailAddress)
            subject = safe_get(msg.Subject)
            body = safe_get(msg.Body)
            to = safe_get(msg.To)

            # Tag logic
            if sender in vip_list:
                category = "VIP"
            elif "@teamschools.zendesk.com" in sender or 'data@kippnj.org' in sender:
                category = "Zendesk"
            else:
                category = "General"

            summary = summarize_email(summarizer, body)

            email_data = {
                "category": category,
                "sender": sender,
                "subject": subject,
                "to": to,
                "summary": summary
            }

            if category == "VIP":
                vip_emails.append(email_data)
            elif category == "Zendesk":
                zendesk_emails.append(email_data)
            else:
                other_emails.append(email_data)

        except Exception as e:
            print("Error reading message:", e)
            continue

    return vip_emails, zendesk_emails, meeting_emails, other_emails


### ------------------------------------------------------------
### 6. Write Summary to Local TXT File
### ------------------------------------------------------------
def write_summary(vip, zendesk, meetings, other):
    with open(SUMMARY_OUTPUT, "w", encoding="utf-8") as f:
        f.write("EMAIL SUMMARY\n===========================\n\n")

        def write_section(title, data):
            f.write(f"\n### {title} ###\n\n")
            for email in data:
                f.write(f"SUBJECT: {email['subject']}\n")
                f.write(f"FROM: {email['sender']}\n")
                f.write(f"TO: {email['to']}\n")
                f.write(f"SUMMARY: {email['summary']}\n")
                f.write("-" * 40 + "\n")

        write_section("VIP Emails", vip)
        write_section("Zendesk Tickets", zendesk)
        write_section("Meeting Invites", meetings)
        write_section("Other Emails", other)

    print(f"Summary written to: {SUMMARY_OUTPUT}")


### ------------------------------------------------------------
### MAIN
### ------------------------------------------------------------
if __name__ == "__main__":
    print("Processing emails...")
    vip, zendesk, meetings, other = process_emails()
    print("Writing summary...")
    write_summary(vip, zendesk, meetings, other)