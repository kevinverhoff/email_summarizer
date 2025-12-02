import os
import csv
import win32com.client
from plyer import notification
from transformers import AutoTokenizer, AutoModelForSeq2SeqLM, pipeline

VIP_FILE = "vip_list.csv"               # stored in repo
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
def get_outlook_emails():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)
    print(f"Found {messages.Count} emails in Inbox")
    return messages


### ------------------------------------------------------------
### 3. Load Local T5 Summarizer
### ------------------------------------------------------------
def load_summarizer():
    tokenizer = AutoTokenizer.from_pretrained("t5-small-email-summarizer")
    print("Loaded t5 small email tokenizer")
    model = AutoModelForSeq2SeqLM.from_pretrained("t5-small-email-summarizer")
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
    if len(body) > 2000:
        note = "Email too long. Showing summary of first 2000 characters. "
        truncated_body = body[:2000]

    summary = summarizer(
        truncated_body,
        max_length=60,
        min_length=15,
        do_sample=False
    )[0]["summary_text"]

    return note + summary


### ------------------------------------------------------------
### 5. Process + Tag Emails
### ------------------------------------------------------------
def process_emails():
    vip_list = load_vip_list()
    messages = get_outlook_emails()
    summarizer = load_summarizer()

    vip_emails = []
    zendesk_emails = []
    other_emails = []
    for msg in list(messages)[:5]:
#    for msg in messages:
        try:
            sender = msg.SenderEmailAddress.lower()
            subject = msg.Subject or ""
            body = msg.Body or ""
            to = msg.To or ""

            # Tag logic
            if sender in vip_list:
                category = "VIP"
            elif "@teamschools.zendesk.com" in sender:
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

            # Sort into buckets
            if category == "VIP":
                vip_emails.append(email_data)
            elif category == "Zendesk":
                zendesk_emails.append(email_data)
            else:
                other_emails.append(email_data)
            print('Processed email from:', sender)

        except Exception as e:
            print("Error reading message:", e)
            continue

    return vip_emails, zendesk_emails, other_emails


### ------------------------------------------------------------
### 6. Write Summary to Local TXT File
### ------------------------------------------------------------
def write_summary(vip, zendesk, other):
    with open(SUMMARY_OUTPUT, "w", encoding="utf-8") as f:
        f.write("EMAIL SUMMARY\n===========================\n\n")

        def write_section(title, data):
            f.write(f"\n### {title} ###\n\n")
            for email in data:
                f.write(f"FROM: {email['sender']}\n")
                f.write(f"TO: {email['to']}\n")
                f.write(f"SUBJECT: {email['subject']}\n")
                f.write(f"SUMMARY: {email['summary']}\n")
                f.write("-" * 40 + "\n")

        write_section("VIP Emails", vip)
        write_section("Zendesk Tickets", zendesk)
        write_section("Other Emails", other)

    print(f"Summary written to: {SUMMARY_OUTPUT}")


### ------------------------------------------------------------
### 7. Pop-up Notification
### ------------------------------------------------------------
def notify():
    notification.notify(
        title="Your Email Inbox Summary Is Ready",
        message="Your email summary has been generated.",
        timeout=8
    )


### ------------------------------------------------------------
### MAIN
### ------------------------------------------------------------
if __name__ == "__main__":
    print("Processing emails...")
    vip, zendesk, other = process_emails()
    print("Writing summary...")
    write_summary(vip, zendesk, other)
    notify()