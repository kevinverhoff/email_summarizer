### Email Summarizer

This repository contains a Python script that summarizes your Outlook emails and highlights messages from VIP contacts.

### Repository

https://github.com/kevinverhoff/email_summarizer/

### Features

- Scans Outlook Classic emails for new messages.
- Identifies emails from VIP contacts.
- Generates a summary of your inbox and saves it to a text file.

### Requirements

- Outlook Classic must be open while running the script.
- Python 3.x
- Packages: Install required packages using `pip install -r requirements.txt` (ensure you have packages like pywin32, transformers, torch, tkinter).

### Setup

1. Clone the repository:

```
git clone https://github.com/kevinverhoff/email_summarizer/
cd email_summarizer
```


2. Create a `vip_list.csv` file in the repo folder. Each line should contain one VIP email address:

```
vip1@example.com
vip2@example.com
...
```

3. Make sure Outlook Classic is running.

### Usage

Run the script:

```
python email_summarizer.py
```

The script will:

- Read your emails from Outlook Classic.
- Identify emails from VIP contacts listed in vip_list.csv.
- Generate a summary and save it to your Desktop as email_summary.txt.

### Notes

Ensure Outlook Classic is open before running the script.

The `vip_list.csv` file must be in the same directory as the script.
