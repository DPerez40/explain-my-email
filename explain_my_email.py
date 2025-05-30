import email
from email import policy
from email.parser import BytesParser

# Load the .eml file
with open("sample.eml", "rb") as f:
    msg = BytesParser(policy=policy.default).parse(f)

# Extract and print key fields
print("From:", msg["from"])
print("To:", msg["to"])
print("Subject:", msg["subject"])
print("Date:", msg["date"])

# Extract the message body
if msg.is_multipart():
    for part in msg.walk():
        if part.get_content_type() == "text/plain":
            print("\nBody:\n", part.get_payload(decode=True).decode())
            break
else:
    print("\nBody:\n", msg.get_payload(decode=True).decode())
