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
            logo.add_header('Content-ID', '<gdg_logo>')
            logo.add_header('Content-Disposition', 'inline', filename='gdg_klef.jpg')
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
        message_body_html = message_template_html.replace("{participant_name}", recipient_name)

        # Send the email
        send_email(recipient_email, subject, message_body_html, logo_path, recipient_name)


if __name__ == "__main__":
    # Path to your CSV or Excel file
    email_data_file_path = "GDG_Test.xlsx"  # Change to .xlsx if you're using Excel

    # Email subject
    email_subject = "Important Instructions for GenAI Study Jams"

    # Path to your logo image (make sure the file is in the same directory or provide a full path)
    logo_image_path = "gdg_klef.jpg"

    # HTML message template (remains unchanged)
    email_body_template_html = """
    <!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Important Instructions for Google Cloud Intro Lab and GenAI Course</title>
</head>
<body style="font-family: Arial, sans-serif; color: #333; background-color: #f4f4f4; margin: 0; padding: 0;">
    <table cellpadding="0" cellspacing="0" border="0" width="100%" style="max-width: 600px; margin: 20px auto; background-color: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 10px rgba(0, 0, 0, 0.1);">
        <tr>
            <td style="background-color: #4285F4; padding: 30px 20px; text-align: center;">
                <img src="https://i.imgur.com/gEaH6w2.png" alt="GDG Logo" style="width: 80px; height: auto;">
                <h1 style="color: #ffffff; font-size: 24px; margin: 20px 0 10px;">Google Cloud &amp; GenAI Course</h1>
                <p style="color: #ffffff; font-size: 16px; margin: 0;">Important Instructions</p>
            </td>
        </tr>
        <tr>
            <td style="padding: 30px 20px;">
                <p style="font-size: 16px; line-height: 1.6; margin-bottom: 20px;">Dear {participant_name},</p>
                <p style="font-size: 16px; line-height: 1.6; margin-bottom: 20px;">We hope this email finds you well. We are reaching out to provide further instructions regarding your GenAI course enrollment. Please follow the steps below carefully to ensure a smooth experience throughout the course:</p>

                <ol style="padding-left: 20px; margin-bottom: 30px;">
                    <li style="margin-bottom: 20px;">
                        <h2 style="color: #4285F4; font-size: 18px; margin-bottom: 10px;">Enrollment Confirmation</h2>
                        <p style="font-size: 16px; line-height: 1.6; margin: 0;">You must have received an enrollment email from Google. Please complete the sign-up process for the Google Cloud Intro Lab to gain access to free credits, which will be essential for completing all tasks and labs.</p>
                    </li>
                    <li style="margin-bottom: 20px;">
                        <h2 style="color: #0F9D58; font-size: 18px; margin-bottom: 10px;">Google Cloud Lab Tutorial</h2>
                        <p style="font-size: 16px; line-height: 1.6; margin: 0;">Watch <a href="https://www.youtube.com/watch?v=WVdUW1wJwyI" style="color: #4285F4; text-decoration: none;">this video tutorial</a> thoroughly. It contains step-by-step guidance to help you get started. Please ensure you follow every instruction as mentioned in the video.</p>
                    </li>
                    <li style="margin-bottom: 20px;">
                        <h2 style="color: #F4B400; font-size: 18px; margin-bottom: 10px;">Syllabus and Task Links</h2>
                        <p style="font-size: 16px; line-height: 1.6; margin: 0;">You can access the complete course syllabus and tasks <a href="https://docs.google.com/spreadsheets/d/1XoxkJ4OvDwN44iDJt6EUesfuJIYiBj8Y5vl_sVL_dKk/pubhtml?gid=0&single=true" style="color: #4285F4; text-decoration: none;">here</a>. This sheet contains links to each task that you need to complete.</p>
                    </li>
                    <li style="margin-bottom: 20px;">
                        <h2 style="color: #DB4437; font-size: 18px; margin-bottom: 10px;">Join the GenAI Discussion Group</h2>
                        <p style="font-size: 16px; line-height: 1.6; margin: 0;">Join the GDG KLEF GenAI Topic on <a href="https://t.me/gdgklef" style="color: #4285F4; text-decoration: none;">Telegram</a>. We will be posting video tutorials and guidance to help you complete each task.</p>
                    </li>
                    <li>
                        <h2 style="color: #4285F4; font-size: 18px; margin-bottom: 10px;">Course Completion Requirement</h2>
                        <p style="font-size: 16px; line-height: 1.6; margin: 0;">Since you have received the enrollment email from Google, it is mandatory to complete all tasks and follow each step as outlined. Completing these tasks will enhance your Google Skill Boost account, which can be used for certifications and will open up numerous opportunities for you.</p>
                    </li>
                </ol>

                <table cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color: #FFF3E0; border-left: 4px solid #F4B400; margin-bottom: 30px;">
                    <tr>
                        <td style="padding: 15px;">
                            <p style="font-size: 16px; line-height: 1.6; margin: 0;"><strong>Important:</strong> Incomplete tasks may leave a negative impression on your Skill Boost profile, so it is important to stay on track.</p>
                        </td>
                    </tr>
                </table>

                <p style="font-size: 16px; line-height: 1.6; margin-bottom: 20px;">We will continue to share step-by-step instructions and updates. Rest assured, you are not alone in this process—we are here to guide you at every step.</p>

                <p style="font-size: 16px; line-height: 1.6; margin-bottom: 20px;">Feel free to reach out if you have any questions.</p>

                <p style="font-size: 16px; line-height: 1.6; margin-bottom: 20px;">Best regards,<br><strong>Team GDG</strong></p>
            </td>
        </tr>
        <tr>
            <td style="background-color: #f4f4f4; padding: 20px; text-align: center;">
                <p style="font-size: 14px; color: #666; margin-bottom: 10px;">Connect with us:</p>
                <a href="https://www.linkedin.com/company/google-developer-groups-on-campus-klu/" target="_blank" rel="noopener noreferrer" style="display: inline-block; margin: 0 10px;">
                    <img src="https://cdn-icons-png.flaticon.com/128/3536/3536505.png" alt="LinkedIn" style="width: 24px; height: 24px;">
                </a>
                <a href="https://www.instagram.com/gdg_klef" target="_blank" rel="noopener noreferrer" style="display: inline-block; margin: 0 10px;">
                    <img src="https://cdn-icons-png.flaticon.com/128/2111/2111463.png" alt="Instagram" style="width: 24px; height: 24px;">
                </a>
            </td>
        </tr>
    </table>
</body>
</html>
    """
  
    # Start the email automation process
    automate_emails(email_data_file_path, email_subject, email_body_template_html, logo_image_path)
