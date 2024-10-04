# CSV-Mail Automation Project

## Overview
This project automates the process of sending personalized emails to recipients listed in a CSV or Excel file. It is designed to streamline the email sending process, making it easier to manage and customize email campaigns.

## Why This Project?
The CSV-mail automation project is useful for anyone who needs to send bulk emails with personalized content. It saves time and effort by automating the process, ensuring that each recipient receives a tailored email.

## Getting Started

### Prerequisites
- A CSV or Excel file containing the recipient data (Name and Email columns).
- A logo image file.
- Environment variables for the sender's email and password.
- Python installed on your system.
- Required libraries: `pandas`, `smtplib`, `email`, and `dotenv`.

### Installation
1. **Clone the Repository:**
   ```bash
   git clone https://github.com/your-repo-url/csv-mail-automation.git
   ```

2. **Install Dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Set Up Environment Variables:**
   Create a `.env` file in the root directory with the following content:
   ```plaintext
   SENDER_EMAIL=your-email@example.com
   SENDER_PASSWORD=your-email-password
   ```

### Usage

#### Prepare Your Data File
Ensure your CSV or Excel file has columns named `Name` and `Email`.

#### Customize the Email Template
Edit the `email_body_template_html` variable in the script to match your needs. You can add placeholders for dynamic content.

#### Run the Script
```bash
python main.py
```

## Project Structure

- **main.py:** The main script that automates the email sending process.
- **.env:** File containing environment variables for the sender's email and password.
- **requirements.txt:** List of dependencies required to run the project.
- **GDG_LOGO.png:** The logo image to be included in the emails.

## How It Works

1. **Load Email Data:**
   The script loads the email data from a CSV or Excel file.
2. **Prepare Email Content:**
   It customizes the HTML email body with the recipient's name and other dynamic content.
3. **Send Emails:**
   The script uses SMTP to send emails to each recipient.
4. **Log Failed Emails:**
   Emails that fail to send are logged in a `failed_emails.txt` file.

## Features

- **Personalized Emails:** Each email is customized with the recipient's name.
- **Automated Process:** The script automates the entire process, saving time and effort.
- **Error Logging:** Failed emails are logged for easy tracking.
- **Customizable Template:** The email template can be easily modified to fit different needs.

## Contributing
Contributions are welcome Here are some ways you can contribute:

### Reporting Issues
If you encounter any issues, please open an issue on the GitHub repository.

### Pull Requests
Feel free to submit pull requests with improvements or new features.

## License
This project is licensed under the [MIT License](https://opensource.org/licenses/MIT).

## Authors
- [Ala Gowtham Siva Kumar](https://github.com/gowtham-2oo5)
- [A Satya Barghav](https://github.com/satyabarghav)

## Acknowledgments
Special thanks to all contributors and the open-source community for their support and resources.

## Troubleshooting Tips
- Ensure the CSV or Excel file is in the correct format.
- Check the environment variables for correctness.
- Verify the SMTP settings and credentials.
