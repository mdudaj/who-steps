import requests
import pandas as pd
import smtplib
import zipfile
import os
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv
from datetime import datetime

# Load environment variables from .env file
load_dotenv()

# Define the API endpoint
url = "https://api.ona.io/api/v1/data/"

# API authentication credentials
api_user = os.getenv('API_USER')
api_password = os.getenv('API_PASSWORD')
auth = (api_user, api_password)

# The form dictionary
form_dict = {
    '739344': 'Tanzania WHO STEPS 1-2 v3.2', 
    '740784': 'Tanzania WHO STEPS 3 (v. 3.2)',
    '739347': 'Tanzania WHO STEPS Lab (v. 3.2)',
}

# The IDs of the forms you want to download data from
form_ids = ['739344', '740784', '739347']

# Create a zip file
timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
zip_file_name = f'Tanzania-WHO-STEPS-Survey-Data-{timestamp}.zip'
with zipfile.ZipFile(zip_file_name, 'w') as zipf:
    for form_id, form_name in form_dict.items():
        # Send a GET request to the API
        response = requests.get(url + form_id, auth=auth)

        # Convert the JSON response to a pandas DataFrame
        data = pd.DataFrame(response.json())

        # Check if the DataFrame is empty
        if not data.empty:
            # Save the DataFrame to an Excel file
            excel_file = f'{form_name.replace(" ", "_")}_{timestamp}.xlsx'
            data.to_excel(excel_file, engine='openpyxl', index=False)

            # Add the Excel file to the zip file
            zipf.write(excel_file)

# Set up the email
msg = MIMEMultipart()
msg['From'] = 'nimrhqs.noreply@gmail.com'
# Add recipients
recipients = ['mchainajr@gmail.com', 'john.a.mduda@gmail.com']
msg['To'] = ", ".join(recipients)

msg['Subject'] = 'WHO STEPS Survey Data'

# Attach the zip file
part = MIMEBase('application', "octet-stream")
part.set_payload(open(zip_file_name, "rb").read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', 'attachment', filename=zip_file_name)  
msg.attach(part)

# Send the email
try:
    email_host = os.getenv('EMAIL_HOST')
    email_port = os.getenv('EMAIL_PORT')
    email_user = os.getenv('EMAIL_USER')
    email_password = os.getenv('EMAIL_PASSWORD')

    smtp = smtplib.SMTP(email_host, email_port)
    smtp.starttls()
    smtp.login(email_user, email_password)
    smtp.sendmail(msg['From'], recipients, msg.as_string())
    smtp.quit()
except smtplib.SMTPAuthenticationError as e:
    print("Failed to connect to the server. Wrong user/password?")
except Exception as e:
    print("SMTP error: ", e)
else:
    print("Successfully sent message to recipient(s).")
