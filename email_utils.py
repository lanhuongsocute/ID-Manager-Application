import smtplib
from email.message import EmailMessage
import os

def send_email(gmail, file_path, id_text, smtp_config):
    msg = EmailMessage()
    msg['Subject'] = f'Tài liệu đính kèm cho ID: {id_text}'
    msg['From'] = smtp_config['from']
    msg['To'] = gmail
    msg.set_content(f'Kính gửi, tài liệu dành cho ID: {id_text}')

    with open(file_path, 'rb') as f:
        msg.add_attachment(f.read(), maintype='application', subtype='octet-stream', filename=os.path.basename(file_path))

    with smtplib.SMTP_SSL(smtp_config['server'], smtp_config['port']) as smtp:
        smtp.login(smtp_config['from'], smtp_config['password'])
        smtp.send_message(msg)
