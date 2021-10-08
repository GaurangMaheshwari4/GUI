import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
sender = "gaurangmaheshwari602@gmail.com"
password = input("Enter password: ")
receiver = ["gaurang.ibd@gmail.com","gaurangmaheshwari.ibd@gmail.com"]
message = '''Hi there'''
msg = MIMEMultipart("alternative")
msg["Subject"] = "SMTP email"
msg["From"] = "gaurangmaheshwari602@gmail.com"
msg["To"] = ",".join(receiver)
part1 = MIMEText(message,'plain')
msg.attach(part1)

server = smtplib.SMTP("smtp.gmail.com",587)
server.starttls()
server.login(sender,password)
print("Login Success")
server.sendmail(sender,receiver,msg.as_string())
print("Email sent to",receiver)
server.quit()