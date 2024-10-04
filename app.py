import os
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Load the email data from either CSV or Excel file
def load_email_data(file_path):
    if file_path.endswith('.xlsx'):
        df = pd.read_excel(file_path)
    elif file_path.endswith('.csv'):
        df = pd.read_csv(file_path)
    else:
        raise ValueError("Unsupported file format. Use either .xlsx or .csv")
    return df

# Log emails that failed to send
def log_failed_email(name, email):
    with open("failed_emails.txt", "a") as log_file:
        log_file.write(f"{name}, {email}\n")

# Send email using Outlook's SMTP server
def send_email(to_email, subject, message_body_html, logo_path, recipient_name):
    # Email server setup for Outlook/Office365
    smtp_server = "smtp.office365.com"
    smtp_port = 587
    sender_email = os.getenv("SENDER_EMAIL")  # Get from environment variables
    sender_password = os.getenv("SENDER_PASSWORD")  # Get from environment variables

    try:
        # Create a secure connection to the SMTP server
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()  # Secure the connection
        server.login(sender_email, sender_password)

        # Create the email content
        msg = MIMEMultipart('related')
        msg['From'] = sender_email
        msg['To'] = to_email
        msg['Subject'] = subject

        # Attach the HTML message body
        html_content = MIMEText(message_body_html, 'html')
        msg.attach(html_content)

        # Attach the logo image with Content-ID (CID)
        with open(logo_path, 'rb') as img_file:
            logo = MIMEImage(img_file.read())
            logo.add_header('Content-ID', '<company_logo>')
            logo.add_header('Content-Disposition', 'inline', filename='company_logo.jpg')
            msg.attach(logo)

        # Send the email
        server.sendmail(sender_email, to_email, msg.as_string())
        print(f"Email sent to {to_email}")

    except Exception as e:
        print(f"Failed to send email to {to_email}. Error: {str(e)}")
        log_failed_email(recipient_name, to_email)  # Log the failed email

    finally:
        server.quit()  # Close the connection to the server

# Main function to automate the email sending
def automate_emails(file_path, subject, message_template_html, logo_path):
    df = load_email_data(file_path)

    for index, row in df.iterrows():
        recipient_email = row['Email']
        recipient_name = row['Name']

        # Customize the HTML message body with the participant's name
        message_body_html = message_template_html.format(
            participant_name=recipient_name,
            event_date="May 15, 2023",
            event_time="2:00 PM",
            event_location="Virtual Meeting"
        )

        # Send the email
        send_email(recipient_email, subject, message_body_html, logo_path, recipient_name)

if __name__ == "__main__":
    # Path to your CSV or Excel file
    email_data_file_path = "participants.xlsx"  # Change to .csv if you're using CSV

    # Email subject
    email_subject = "Invitation: Upcoming Virtual Event"

    # Path to your logo image (make sure the file is in the same directory or provide a full path)
    logo_image_path = "company_logo.jpg"

    # Simplified HTML message template
    email_body_template_html = """
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Event Invitation</title>
    </head>
    <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
        <table width="100%" cellpadding="0" cellspacing="0" style="max-width: 600px; margin: 0 auto;">
            <tr>
                <td style="text-align: center; padding: 20px 0;">
                    <img src="cid:company_logo" alt="Company Logo" style="max-width: 200px;">
                </td>
            </tr>
            <tr>
                <td style="padding: 20px;">
                    <h1 style="color: #4A90E2;">Hello {participant_name},</h1>
                    <p>You are cordially invited to our upcoming virtual event. Here are the details:</p>
                    <ul>
                        <li><strong>Date:</strong> {event_date}</li>
                        <li><strong>Time:</strong> {event_time}</li>
                        <li><strong>Location:</strong> {event_location}</li>
                    </ul>
                    <p>We look forward to seeing you there!</p>
                    <p>Best regards,<br>Your Event Team</p>
                </td>
            </tr>
            <tr>
                <td style="text-align: center; padding: 20px; background-color: #f0f0f0;">
                    <p style="margin: 0;">© 2023 Your Company. All rights reserved.</p>
                </td>
            </tr>
        </table>
    </body>
    </html>
    """

    # Start the email automation process
    automate_emails(email_data_file_path, email_subject, email_body_template_html, logo_image_path)
