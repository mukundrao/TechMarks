import smtplib
import os
import imghdr
from email.message import EmailMessage
import sys


def send_mail(fr, to, course, html, image):                              #this function sends marks to students through mail along with their PAC as an attachment
    message = EmailMessage()
    message['From'] = fr
    message['To'] = to
    message['Subject'] = f"{course} Internals Marks and Student Performance Chart"

    message.add_alternative(html, subtype="html")                        #command to add html content to mail

    with open(image, "rb") as imeg:                                      #procedure to add the PAC image to mail as an attachment
        file_data = imeg.read()                                          #reading the image
        file_type = imghdr.what(imeg.name)                               #deriving the image file type
        file_name = imeg.name                                            #deriving the image name

    message.add_attachment(file_data, maintype = "image", subtype=file_type, filename=file_name)                       #attaching the PAC image to mail

    with smtplib.SMTP(host = 'smtp.gmail.com', port = 587) as smtp:                     # establishing communication with Gmail server using SMTP at port no 587
        smtp.ehlo()                                                                     #identifying the sender's domain name to SMTP
        smtp.starttls()                                                                 #enabling encryption for communication with Gmail server

        if os.environ.get('mailid') == None or os.environ.get('sagpwd') == None:        #condition to check for presence of required data in system environment
            sys.exit("Incorrect username and password!Failed to login to Google Account")

        try:
            smtp.login(os.environ.get('mailid'), os.environ.get('sagpwd'))              #logging into the sender's Google Account to send mails
        except smtplib.SMTPAuthenticationError:                           #exception to handle incorrect username and password credentials given by the sender for logging in
            sys.exit("Incorrect username and password!Failed to login to Google Account")

        try:
            smtp.send_message(message)                                     #sending the mail to a student
        except smtplib.SMTPRecipientsRefused:
            smtp.quit()
            return

        smtp.quit()                                                    #ending the session/communication with Gmail server

