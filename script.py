import smtplib
import pandas as pd
from email.message import EmailMessage

# Gmail credentials
EMAIL_ADDRESS = 'areeba099876@gmail.com'
EMAIL_PASSWORD = ''  # 16-character App Password

# Load contacts from Excel
contacts = pd.read_excel('candidates.xlsx')  

# Loop through each contact
for index, row in contacts.iterrows():
    name = row['Name']
    email = row['Email']

    # Create the email content
    msg = EmailMessage()
    msg['Subject'] = 'Interview Invitation'
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = email

    body = f"""
    Dear {name},

We are pleased to inform you that you have been shortlisted for an interview for the IT department at Bank AL Habib.

Please bring a copy of your updated CV and appear for the interview as per the following details:

Date: 12th May
Time: 9:00 AM
Venue: Bank AL Habib, Main Branch

We look forward to meeting you.


    Best regards,  
   Areeba Khan
    """
    msg.set_content(body)

    # Send the email
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            smtp.send_message(msg)
            print(f"[✓] Email sent to {name} at {email}")
    except Exception as e:
        print(f"[✗] Failed to send email to {name} at {email}. Error: {e}")
