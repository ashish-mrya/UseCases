import smtplib

gmail_user = 'ashish.mrya@gmail.com'
gmail_password = '...'


sent_from = "ashish.mrya@gmail.com"
to = ['ashish.mrya@gmail.com']
subject = 'OMG Super Important Message'
body = 'Hey, what\'s up?\n\n - You'
email_text = "This is a dummy email for testing"

try:
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.ehlo()
    server.login(gmail_user, gmail_password)
    print("Connection successful")
    server.sendmail(sent_from, to, email_text)
    server.close()
    print("email succesfully sent")
except:
    print("Unable to login to gmail server")

