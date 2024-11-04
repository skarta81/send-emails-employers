import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import time
from pathlib import Path

# Constants
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
EMAIL_ADDRESS = '****'
PASSWORD = '****'
CSV_FILE_PATH = r'C:\Users\shake\Downloads\Employers_Cert\Employers.csv'
CERTIFICATES_FOLDER = r'C:\Users\shake\Downloads\Employers_Cert\Certificates'

def load_data(csv_path):
    """Load CSV data into a DataFrame."""
    return pd.read_csv(csv_path)

def create_email_body(employee_name):
    """Generate the email body text in Hebrew, formatted for RTL."""
    body = f"""
    <div dir="rtl">
        שלום רב,<br><br>
        מצורפת תעודת הוקרה מגדוד מגן 9260, על תמיכתכם בשרות המילואים של {employee_name}.<br><br>
        בהוקרה והערכה רבה,<br><br>
        סא״ל יאיר שימל,<br>
        מפקד גדוד מגן 9260
    </div>
    """
    return body

def get_certificate_path(employee_name):
    """Retrieve the path of the certificate for the given employee."""
    certificate_path = Path(CERTIFICATES_FOLDER) / f"{employee_name}.jpg"
    return certificate_path if certificate_path.exists() else None

def send_email(to_addresses, subject, body, certificate_path, employee_name, current_iteration, total):
    """Send an email with the specified subject, body, and certificate attachment."""
    msg = MIMEMultipart()
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = ', '.join(to_addresses)
    msg['Subject'] = subject

    # Attach body text
    msg.attach(MIMEText(body, 'html', 'utf-8'))

    # Attach the certificate image as a file attachment only
    with open(certificate_path, 'rb') as cert_file:
        mime_image_attachment = MIMEImage(cert_file.read())
        mime_image_attachment.add_header('Content-Disposition', 'attachment', filename=certificate_path.name)
        msg.attach(mime_image_attachment)

    # Send the email
    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_ADDRESS, PASSWORD)
            server.sendmail(EMAIL_ADDRESS, to_addresses, msg.as_string())
        print(f"[{current_iteration}/{total}] Email sent successfully to {', '.join(to_addresses)} for employee {employee_name}")
    except Exception as e:
        print(f"[{current_iteration}/{total}] Failed to send email to {', '.join(to_addresses)} for employee {employee_name}: {e}")

def process_emails(data):
    """Process each row in the data to send emails to employers with a delay."""
    subject = "תעודת הוקרה מגדוד מגן 9260"
    total_rows = len(data)
    for index, row in data.iterrows():
        employee_name = row['שם מלא של החייל'].strip()
        to_addresses = [row[column] for column in ['כתובת דואר אלקטרוני של המעסיק (אליו תישלח התעודה)', 'כתובת דואר אלקטרוני 2', 'כתובת דואר אלקטרוני 3', 'כתובת דואר אלקטרוני 4'] if pd.notna(row[column])]

        if to_addresses:
            certificate_path = get_certificate_path(employee_name)

            if certificate_path:
                body = create_email_body(employee_name)
                send_email(to_addresses, subject, body, certificate_path, employee_name, index + 1, total_rows)
                time.sleep(2)  # Add a delay of 2 seconds
            else:
                print(f"[{index + 1}/{total_rows}] No certificate found for {employee_name}, skipping email.")
        else:
            print(f"[{index + 1}/{total_rows}] No email addresses found for {employee_name}, skipping email.")

def main():
    # Load CSV data
    data = load_data(CSV_FILE_PATH)

    # Process and send emails
    process_emails(data)

if __name__ == "__main__":
    main()
