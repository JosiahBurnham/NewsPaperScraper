from os import stat
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import pathlib

class EmailHelper:
    @staticmethod
    def send_email(email_recipient,
                email_subject,
                email_message,
                attachment_location = ''):

        """ sends an email

        Sends and email using SMTP with a google account.

        :param email_recipient:
         the recipient of the email
        :param email_subject:
         the subject of the email
        :param email_message:
         the message in the email
        :param attachment_location:
         the PATH to the attachment being sent

        :return: True
        """
        email_sender = 'YOUR EMAIL HERE'

        msg = MIMEMultipart()
        msg['From'] = email_sender
        msg['To'] = email_recipient
        msg['Subject'] = email_subject

        msg.attach(MIMEText(email_message, 'plain'))

        if attachment_location != '':
            filename = str(pathlib.Path(__file__).parent.absolute())  +"\\"+ attachment_location
            attachment = open(filename, "rb")
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition',
                            "attachment; filename= %s" % attachment_location)
            msg.attach(part)

        try:
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.ehlo()
            server.starttls()
            server.login('YOUR EMAIL', 'YOUR PASSWORD TO YOUR EMAIL')
            text = msg.as_string()
            server.sendmail(email_sender, email_recipient, text)
            print('EMAIL SENT')
            server.quit()
        except AttributeError:
            print("SMPT server connection error")
        return True
