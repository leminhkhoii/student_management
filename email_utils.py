from email.mime.base import MIMEBase
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

#entd uvtq hzpl rnae

def send_email(recipients, student_name, subject, body):
    sender_email = "quoctrung87377@gmail.com"
    sender_password = "entd uvtq hzpl rnae"  # Replace with your actual app-specific password

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = ", ".join(recipients)  # Join recipients list into a comma-separated string
    msg['Subject'] = subject
    personalized_body = f"Xin chào sinh viên {student_name},\n\n{body}"
    msg.attach(MIMEText(personalized_body, 'plain'))

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        text = msg.as_string()
        server.sendmail(sender_email, recipients, text)  # Send to recipients as a list
        server.quit()
        print("Email sent successfully")
    except Exception as e:
        print(f"Failed to send email: {e}")




