# 📧 Email Automation

A production-ready cold email automation tool using Gmail API. Built for sending personalized outreach emails with safe sending limits, duplicate prevention, and resume capability.

## Features

- ✅ **Gmail API Integration** - Safe from account bans (no SMTP)
- ✅ **Personalized Emails** - Dynamic `{first_name}` and `{company_name}` placeholders
- ✅ **Duplicate Prevention** - Tracks sent status in Excel
- ✅ **Resume Capability** - Pick up where you left off
- ✅ **Safe Batch Sending** - 22 emails/day with random delays
- ✅ **Progress Tracking** - Real-time console output

## Setup

### 1. Install Dependencies

```bash
python3 -m venv venv
source venv/bin/activate
pip install openpyxl python-dotenv google-auth-oauthlib google-api-python-client
```

### 2. Google Cloud Setup

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project
3. Enable **Gmail API**
4. Create OAuth 2.0 credentials (Desktop app)
5. Download `credentials.json` to this directory
6. Add your email as a **test user** in OAuth consent screen

### 3. Prepare Excel File

Your Excel file needs these columns:
- `personal_email` - Recipient email
- `first_name` - For personalization
- `company_name` - For personalization
- `company_linkedin_url` - Fallback for company name
- `position` - Job title

The script will add an `email_status` column to track sent emails.

## Configuration

Edit the script to configure:

```python
# File paths
EXCEL_FILENAME = "your_contacts.xlsx"
INPUT_SHEET = "found_blitz"

# Batch settings
BATCH_CONFIG = {
    'emails_per_day': 22,      # Safe limit
    'delay_min': 20,           # Seconds between emails
    'delay_max': 35,
    'batch_size': 10,          # Emails before long break
    'batch_break': 300,        # 5 minutes
}
```

## Usage

```bash
source venv/bin/activate
python send_emails_blitz_amf.py
```

First run will open a browser for Gmail OAuth. After authorization, type `yes` to start sending.

## Email Template

The script sends personalized emails asking about backend opportunities. Edit `EMAIL_SUBJECT` and `EMAIL_BODY_TEMPLATE` in the script to customize.

## Safety Notes

- ⚠️ Keep `credentials.json` and `token.pickle` private
- ⚠️ Don't exceed 50 emails/day on personal Gmail
- ⚠️ Use random delays to appear human

## License

MIT
