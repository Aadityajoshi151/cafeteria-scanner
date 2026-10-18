import smtplib, getpass
from email.message import EmailMessage

sender = input("Gmail address: ").strip()
password = getpass.getpass("App password: ").replace(" ", "")
to = input("Send test to: ").strip()

msg = EmailMessage()
msg["Subject"] = "Cafeteria scanner test"
msg["From"] = sender
msg["To"] = to
msg.set_content("If you can read this, email from the Pi works.")

with smtplib.SMTP("smtp.gmail.com", 587, timeout=30) as s:
    s.starttls()
    s.login(sender, password)
    s.send_message(msg)
print("Sent!")
