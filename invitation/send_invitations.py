#!/usr/bin/env python3
"""
DaKshaa T26 Event Invitation Email Sender
Reads recipient details from Excel and sends HTML-formatted invitations via SMTP
"""

import smtplib
import os
import sys
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook
from dotenv import load_dotenv
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
import time

# Load environment variables from .env file if it exists
load_dotenv()

SCRIPT_DIR = Path(__file__).resolve().parent


def load_logos_for_email():
    """
    Load logo files and return dict of {cid: (bytes, mime_subtype)}.
    Logos are attached as inline MIME parts and referenced via cid: in HTML.
    """
    logos = {}
    configs = [
        ('dakshaa_logo', 'eventlogio_nobg.png', 'png'),
        ('ksrct_logo', 'ksrct_logo_nobg.png', 'png'),
    ]
    for cid, filename, subtype in configs:
        path = SCRIPT_DIR / filename
        try:
            with open(path, 'rb') as f:
                logos[cid] = (f.read(), subtype)
        except FileNotFoundError:
            pass
    return logos


def get_html_template():
    """Generate HTML email template."""

    return """
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>DaKshaa T26 | National-Level Techno-Cultural Fest</title>
</head>

<body style="margin:0; padding:0; background-color:#edf2f7; font-family: Arial, Helvetica, sans-serif;">

<table width="100%" cellpadding="0" cellspacing="0" style="padding:25px 15px;">
  <tr>
    <td align="center">

      <table width="650" cellpadding="0" cellspacing="0"
             style="background-color:#ffffff; border-radius:10px; overflow:hidden;">

        <!-- HERO HEADER -->
        <tr>
          <td style="background-color:#0b3c5d; padding:30px 25px; text-align:center;">
            <img src="cid:dakshaa_logo"
                 alt="DaKshaa T26 Logo"
                 style="max-width:220px; height:auto; margin-bottom:18px;">
            <h1 style="margin:0; font-size:28px; color:#ffffff;">
              DaKshaa T26
            </h1>
            <p style="margin:8px 0 0; font-size:15px; color:#dbe9f5;">
              National-Level Techno-Cultural Fest
            </p>
            <p style="margin:6px 0 0; font-size:14px; color:#bcd6ec;">
              12<sup>th</sup> – 14<sup>th</sup> February 2026 | KSRCT
            </p>
          </td>
        </tr>

        <!-- CONTENT -->
        <tr>
          <td style="padding:35px; color:#333333; font-size:15px; line-height:1.75;">

            <p>Hey There 👋</p>

            <p>
              We’re excited to invite you to <strong>DaKshaa T26</strong>, a
              <strong>National-Level Techno-Cultural Fest</strong> organized by
              <strong>K. S. Rangasamy College of Technology (Autonomous), Tiruchengode</strong>.
            </p>

            <!-- OVERVIEW -->
            <h3 style="color:#0b3c5d; margin-top:25px;">Event Overview</h3>

            <p>
              DaKshaa T26 brings together students from across the country for
              three high-energy days of innovation, creativity, competition, and celebration.
              The fest blends cutting-edge technology with cultural expression,
              creating a platform where ideas turn into real impact.
            </p>

            <p>
              The fest features <strong>20+ technical events</strong>,
              <strong>15+ hands-on workshops</strong>, <strong>8+ hackathons</strong>,
              national conferences, tech talks by industry leaders,
              startup pitching sessions, cultural shows, and sports events.
            </p>

            <p>
              With flagship events like <strong>Neura Hack 2.0 (36-hour hackathon)</strong>
              and <strong>VibeCode'26</strong>, and a
              <strong>prize pool worth ₹10 Lakhs</strong>,
              DaKshaa T26 is set to be one of the biggest student festivals of 2026.
            </p>

            <!-- EVENT INFO -->
            <table width="100%" cellpadding="0" cellspacing="0"
                   style="background-color:#f5f9fd; border-left:5px solid #0b3c5d; margin:30px 0;">
              <tr>
                <td style="padding:18px;">
                  <strong>📅 Dates:</strong> 12<sup>th</sup> – 14<sup>th</sup> February 2026<br>
                  <strong>📍 Venue:</strong> K. S. Rangasamy College of Technology, Tamil Nadu
                </td>
              </tr>
            </table>

            <!-- CTA BUTTONS -->
            <table width="100%" cellpadding="0" cellspacing="0" style="margin:35px 0;">
              <tr>
                <td align="center">

                  <!-- Primary CTA -->
                  <a href="https://dakshaa.ksrct.ac.in"
                     style="background-color:#0b3c5d; color:#ffffff;
                            padding:15px 34px; text-decoration:none;
                            font-size:15px; border-radius:30px;
                            display:inline-block; margin:6px;">
                    Explore Events &amp; Register
                  </a>

                  <!-- Secondary CTA -->
                  <a href="https://dakshaa.ksrct.ac.in/assets/Brochure-BKLLtMz-.pdf"
                     style="background-color:#ffffff; color:#0b3c5d;
                            padding:13px 32px; text-decoration:none;
                            font-size:14px; border-radius:30px;
                            border:2px solid #0b3c5d;
                            display:inline-block; margin:6px;">
                    Download Brochure (PDF)
                  </a>

                </td>
              </tr>
            </table>

            <p>
              Get ready to learn, compete, and celebrate.
              We can’t wait to see you at <strong>DaKshaa T26</strong>!
            </p>

            <p>
              Cheers,<br>
              <strong>Team DaKshaa T26</strong>
            </p>

          </td>
        </tr>

        <!-- COLLEGE LOGO FOOTER -->
        <tr>
          <td style="background-color:#f1f1f1; text-align:center; padding:18px;">
            <img src="cid:ksrct_logo"
                 alt="KSRCT Logo"
                 style="max-width:160px; height:auto;">
            <p style="margin:8px 0 0; font-size:12px; color:#666666;">
              K. S. Rangasamy College of Technology
            </p>
          </td>
        </tr>

      </table>

    </td>
  </tr>
</table>

</body>
</html>
"""


def read_recipients_from_excel(file_path):
    """Read recipient emails from Excel file (only emails)"""
    recipients = []
    
    try:
        workbook = load_workbook(filename=file_path, read_only=True)
        sheet = workbook.active
        
        # Skip header row and read data
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row[0]:  # Check if email exists
                email = str(row[0]).strip()
                if email:
                    recipients.append(email)
        
        workbook.close()
        return recipients
    
    except FileNotFoundError:
        print(f"❌ Error: Excel file not found at {file_path}")
        return None
    except Exception as e:
        print(f"❌ Error reading Excel file: {str(e)}")
        return None


def get_smtp_config():
    """Get SMTP configuration from environment variables or user input"""
    config = {}
    
    # Try to get from environment variables first
    config['server'] = "smtp.gmail.com"
    config['port'] = "587"
    config['email'] = "dakshaa@ksrct.ac.in"  # Change to your Gmail address
    config['password'] = "afqx fcmf qsvo iqyv"  # Change to your Gmail app password
    
    # If not in environment, prompt user
    if not all([config['server'], config['port'], config['email'], config['password']]):
        print("\n📧 SMTP Configuration")
        print("=" * 50)
        
        if not config['server']:
            print("\nCommon SMTP servers:")
            print("  Gmail: smtp.gmail.com")
            print("  Outlook: smtp-mail.outlook.com")
            config['server'] = input("Enter SMTP server: ").strip()
        
        if not config['port']:
            config['port'] = input("Enter SMTP port (default 587): ").strip() or "587"
        
        if not config['email']:
            config['email'] = input("Enter your email address: ").strip()
        
        if not config['password']:
            import getpass
            config['password'] = getpass.getpass("Enter your email password/app password: ")
    
    config['port'] = int(config['port'])
    return config


def send_single_email(email, smtp_config, logos, idx, total):
    """Send a single email to one recipient (thread-safe)."""
    try:
        # Create new connection for this thread
        server = smtplib.SMTP(smtp_config['server'], smtp_config['port'], timeout=30)
        server.starttls()
        server.login(smtp_config['email'], smtp_config['password'])
        
        # Create message
        msg = MIMEMultipart('related')
        msg['Subject'] = "You're Invited to DaKshaa T26 – National-Level Techno-Cultural Fest | KSRCT"
        msg['From'] = f"Team DaKshaa T26 <{smtp_config['email']}>"
        msg['To'] = email
        
        html_content = get_html_template()
        html_part = MIMEText(html_content, 'html', 'utf-8')
        msg.attach(html_part)
        
        # Attach inline logos
        for cid, (payload, subtype) in logos.items():
            img = MIMEImage(payload, _subtype=subtype)
            img.add_header('Content-Disposition', 'inline', filename=cid)
            img.add_header('Content-ID', f'<{cid}>')
            msg.attach(img)
        
        # Send email
        server.send_message(msg)
        server.quit()
        
        print(f"✅ [{idx}/{total}] Sent to {email}")
        return {'status': 'success', 'email': email}
        
    except Exception as e:
        print(f"❌ [{idx}/{total}] Failed to send to {email}: {str(e)}")
        return {'status': 'failed', 'email': email, 'error': str(e)}


def send_invitation_emails(recipients, smtp_config, excel_file):
    """Send invitation emails to all recipients with inline logo images (CID) using parallel processing."""
    
    print(f"\n📨 Preparing to send {len(recipients)} invitation emails...")
    print("=" * 50)
    
    # Test connection first
    print(f"\n🔌 Testing connection to {smtp_config['server']}:{smtp_config['port']}...")
    try:
        server = smtplib.SMTP(smtp_config['server'], smtp_config['port'], timeout=30)
        server.starttls()
        server.login(smtp_config['email'], smtp_config['password'])
        server.quit()
        print("✅ Connection successful!\n")
    except smtplib.SMTPAuthenticationError:
        print("\n❌ Authentication failed! Please check your email and password.")
        print("   For Gmail, make sure you're using an App Password, not your regular password.")
        print("   Generate one at: https://myaccount.google.com/apppasswords")
        return None, None
    except Exception as e:
        print(f"\n❌ Error connecting to SMTP server: {str(e)}")
        return None, None
    
    logos = load_logos_for_email()
    if logos:
        print(f"🖼️  Logos loaded: {', '.join(logos.keys())}")
    else:
        print("⚠️  No logo files found; images will not display inline.")
    
    successful = []
    failed = []
    
    # Send emails sequentially with delay to avoid spam filters
    print("📬 Sending emails sequentially (slow mode to avoid spam)...\n")
    for idx, email in enumerate(recipients, 1):
        result = send_single_email(email, smtp_config, logos, idx, len(recipients))
        if result['status'] == 'success':
            successful.append(result['email'])
        else:
            failed.append({'email': result['email'], 'error': result.get('error', 'Unknown error')})
        time.sleep(2)  # 2-second delay between emails
    
    print("\n" + "=" * 50)
    
    # Print summary
    print(f"\n📊 Summary:")
    print(f"   ✅ Successfully sent: {len(successful)}")
    print(f"   ❌ Failed: {len(failed)}")
    
    return successful, failed


def main():
    """Main function"""
    print("\n" + "=" * 60)
    print("  DaKshaa T26 - Email Invitation Sender")
    print("  K. S. Rangasamy College of Technology")
    print("=" * 60)
    
    # Use fixed Excel file path
    excel_file = "recipients.xlsx"
    
    # Check if file exists
    if not Path(excel_file).exists():
        print(f"\n❌ Error: File '{excel_file}' not found!")
        print("\nPlease make sure the Excel file exists with columns:")
        print("  - Column A: Email")
        sys.exit(1)
    
    # Read recipients
    print(f"\n📖 Reading recipients from {excel_file}...")
    recipients = read_recipients_from_excel(excel_file)
    
    if not recipients:
        print("\n❌ No valid recipients found in the Excel file!")
        sys.exit(1)
    
    print(f"✅ Found {len(recipients)} recipients")
    
    # Show preview
    print("\n📋 Preview of recipients:")
    for i, r in enumerate(recipients[:5], 1):
        print(f"   {i}. {r}")
    if len(recipients) > 5:
        print(f"   ... and {len(recipients) - 5} more")
    
    # Confirm
    confirm = input("\n⚠️  Proceed with sending emails? (yes/no): ").strip().lower()
    if confirm not in ['yes', 'y']:
        print("\n❌ Cancelled by user")
        sys.exit(0)
    
    # Get SMTP config
    smtp_config = get_smtp_config()
    
    # Send emails
    successful, failed = send_invitation_emails(recipients, smtp_config, excel_file)
    
    if successful is not None:
        print("\n✅ Email sending process completed!")
    else:
        print("\n❌ Email sending process failed!")
        sys.exit(1)


if __name__ == "__main__":
    main()
