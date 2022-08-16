import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def send_email(usr: str, pss: str, to_who: str, subject: str, body: str):
    
    mimemsg = MIMEMultipart()
    mimemsg['From'] = 'from'
    mimemsg['To'] = to_who
    mimemsg['Subject'] = subject
    mimemsg.attach(MIMEText(body, 'html'))

    connection = smtplib.SMTP(host='smtp.office365.com', port=587)
    connection.starttls()
    connection.login(usr,pss)
    connection.send_message(mimemsg)
    connection.quit()


