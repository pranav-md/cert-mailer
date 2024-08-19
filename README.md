# Automated Certificate Generator and Sender

This Python script automates the process of generating and sending certificates based on user data stored in an Excel worksheet. It streamlines the workflow by handling the following tasks:

1. **Fetch Excel Worksheet**: The script reads data from an Excel file, retrieving user information row by row.
2. **Generate Certificates**: For each user, it populates a certificate template with the relevant data.
3. **Email Certificates**: Once generated, the certificate is automatically sent to the user's email address.

## Features

- **Seamless Data Handling**: Effortlessly manages user data stored in Excel files.
- **Customizable Templates**: Supports the use of customizable certificate templates.
- **Automated Email Sending**: Integrates with email services to automate certificate distribution.
- **Error Handling**: Includes basic error handling for smooth operation.

## Usage

1. Place the Excel worksheet containing user data in the designated directory.
2. Customize the certificate template as needed.
3. Run the script to generate and send certificates to each user listed in the Excel file.

## Dependencies

- `PIL` (Pillow): For handling image manipulation of the certificate template.
- `xlrd`: For reading Excel files.
- `smtplib`: For sending emails via SMTP.
- `email.mime`: For creating and handling email messages and attachments.


## How to Run

1. Clone this repository.
2. Install the necessary dependencies:
   ```bash
   pip install Pillow xlrd
3. Configure the email settings in the script (e.g., SMTP server, port, login credentials).
4. Place the Excel worksheet containing user data in the designated directory.
5. Customize the certificate template as needed.
6. Run the script:
   ```bash
   python3 certificate_generator.py


