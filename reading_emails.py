import imaplib
import email
from email.header import decode_header
import webbrowser
import os

# account credentials
username = "austinschradertest@outlook.com"
password = "royalb123"

# number of top emails to fetch
N = 1

# create an IMAP4 class with SSL, use your email provider's IMAP server
imap = imaplib.IMAP4_SSL("imap.outlook.com")
# authenticate
imap.login(username, password)

# select a mailbox (in this case, the inbox mailbox)
# use imap.list() to get the list of mailboxes
status, messages = imap.select("Inbox")

# total number of emails
messages = int(messages[0])

for i in range(messages, messages-N, -1):
    # fetch the email message by ID
    res, msg = imap.fetch(str(i), "(RFC822)")
    
    for response in msg:
        if isinstance(response, tuple):
            # parse a bytes email into a message object
            msg = email.message_from_bytes(response[1])
            # decode the email subject
            subject = decode_header(msg["Subject"])[0][0]
            if isinstance(subject, bytes):
                # if it's a bytes, decode to str
                subject = subject.decode()
            # email sender
            from_ = msg.get("From")

            # if the email message is multipart
            if msg.is_multipart():
                # iterate over email parts
                for part in msg.walk():
                    # extract content type of email
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))
                    try:
                        # get the email body
                        body = part.get_payload(decode=True).decode()
                    except:
                        pass

                    if "attachment" in content_disposition:
                        # download attachment
                        filename = part.get_filename()
                        if filename:
                            if not os.path.isdir(subject):
                                # make a folder for this email (named after the subject)
                                pass
                            # name the attachment pdf as subject + its current filename
                            filepath = os.path.join(subject + " " + filename)
                            # download attachment and save it
                            open(filepath, "wb").write(part.get_payload(decode=True))

# close the connection and logout
imap.close()
imap.logout()