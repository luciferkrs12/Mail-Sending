# DaKshaa T26 Email Invitation System

Automated email invitation system for sending HTML-formatted event invitations to students. Reads recipient details from an Excel file and sends personalized emails via SMTP.

## 📋 Features

- ✅ Read recipient data from Excel (Name, Email)
- ✅ Beautiful HTML-formatted email template
- ✅ Personalized greetings for each recipient
- ✅ Progress tracking during sending
- ✅ Detailed logging of successful and failed sends
- ✅ Secure credential handling
- ✅ Support for Gmail, Outlook, and other SMTP servers

## 🚀 Quick Start

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Create Sample Excel File (Optional)

Run this to create a sample Excel file with test data:

```bash
python create_sample_excel.py
```

This creates `recipients.xlsx` with sample data. **Replace with your actual recipient list before sending!**

### 3. Prepare Your Excel File

Your Excel file should have exactly 2 columns:

| Name | Email |
|------|-------|
| Rajesh Kumar | rajesh.kumar@example.com |
| Priya Sharma | priya.sharma@example.com |

- **Column A**: Recipient's full name
- **Column B**: Recipient's email address
- First row should be headers (will be skipped)

### 4. Configure SMTP Settings

#### Option A: Using Environment Variables (Recommended)

1. Copy `.env.example` to `.env`:
   ```bash
   copy .env.example .env
   ```

2. Edit `.env` and add your credentials:
   ```
   SMTP_SERVER=smtp.gmail.com
   SMTP_PORT=587
   SMTP_EMAIL=your-email@gmail.com
   SMTP_PASSWORD=your-app-password
   ```

#### Option B: Enter Manually

The script will prompt you for credentials when you run it.

### 5. Run the Script

```bash
python send_invitations.py
```

The script will:
1. Ask for the Excel file path (default: `recipients.xlsx`)
2. Show a preview of recipients
3. Ask for confirmation
4. Request SMTP credentials (if not in `.env`)
5. Send emails with progress updates
6. Generate a log file with results

## 🔐 SMTP Configuration

### Gmail Setup

1. Enable 2-Step Verification on your Google account
2. Generate an App Password:
   - Go to https://myaccount.google.com/apppasswords
   - Select "Mail" and your device
   - Copy the 16-character password
3. Use these settings:
   ```
   SMTP_SERVER=smtp.gmail.com
   SMTP_PORT=587
   SMTP_EMAIL=your-email@gmail.com
   SMTP_PASSWORD=your-16-char-app-password
   ```

**Note**: Gmail has a limit of ~500 emails per day.

### Outlook/Hotmail Setup

```
SMTP_SERVER=smtp-mail.outlook.com
SMTP_PORT=587
SMTP_EMAIL=your-email@outlook.com
SMTP_PASSWORD=your-password
```

**Note**: Outlook has a limit of ~300 emails per day.

### Other SMTP Servers

Contact your email provider for SMTP settings.

## 📊 Output

After sending, you'll get:

1. **Console output** with real-time progress
2. **Log file** (`email_log_YYYYMMDD_HHMMSS.txt`) containing:
   - Timestamp
   - List of successful sends
   - List of failed sends with error messages

## ⚠️ Important Notes

- **Test first**: Send to yourself or a test email before bulk sending
- **Sending limits**: Most providers limit daily emails (Gmail: 500, Outlook: 300)
- **App passwords**: Gmail requires App Passwords, not your regular password
- **Spam filters**: Some recipients may receive emails in spam folder
- **Batch sending**: For large lists, consider splitting into multiple batches

## 🐛 Troubleshooting

### "Authentication failed"
- For Gmail: Make sure you're using an App Password, not your regular password
- Check that 2-Step Verification is enabled
- Verify email and password are correct

### "Connection refused"
- Check SMTP server and port are correct
- Ensure your firewall isn't blocking the connection
- Try port 465 (SSL) instead of 587 (TLS)

### "Recipients not found"
- Verify Excel file path is correct
- Check that columns A and B contain Name and Email
- Ensure there's data starting from row 2 (row 1 is headers)

### Emails going to spam
- Ask recipients to check spam folder
- Add your domain to their safe senders list
- Consider using a dedicated email service for bulk sending

## 📁 Project Structure

```
invitation/
├── send_invitations.py      # Main script
├── create_sample_excel.py   # Helper to create sample Excel
├── requirements.txt         # Python dependencies
├── .env.example            # Environment variables template
├── .env                    # Your actual credentials (create this)
├── recipients.xlsx         # Your recipient list
└── README.md              # This file
```

## 📧 Email Template

The email includes:
- Professional HTML design with gradient header
- Personalized greeting with recipient's name
- Complete event details (dates, venue, activities)
- Clickable registration link
- Mobile-responsive design
- KSRCT branding

## 🔒 Security

- Never commit `.env` file to version control
- Use App Passwords instead of regular passwords
- Keep your credentials secure
- Review recipient list before sending

## 📝 License

Created for K. S. Rangasamy College of Technology - DaKshaa T26 Event

---

**Need help?** Contact Team DaKshaa T26
