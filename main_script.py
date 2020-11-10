# Import all dependencies
import imaplib
import email
from tkinter.filedialog import askdirectory

# Email credentials
user_mail = input("What is your email address? :")
mdp_mail = input("What is your email password? :")

# IMAP server used
server = 'imap-mail.outlook.com'
# Path to the download folder
path = askdirectory(title='Select Folder')  # shows dialog box and return the path


# Download attachments from an IMAP web server
def connect(server, user_mail, mdp_mail):
    m = imaplib.IMAP4_SSL(server)
    m.login(user_mail, mdp_mail)
    m.select()
    return m


def download(m, emailid, path):
    resp, data = m.fetch(emailid, "(BODY.PEEK[])")
    email_body = data[0][1]
    mail = email.message_from_bytes(email_body)
    if mail.get_content_maintype() != 'multipart':
        return
    for part in mail.walk():
        if part.get_content_maintype() != 'multipart' and part.get('Content-Disposition') is not None:
            open(path + '/' + part.get_filename(), 'wb').write(part.get_payload(decode=True))


def query():
    m = connect(server, user_mail, mdp_mail)
    m.select("Inbox")
    typ, msgs = m.search(None, 'UNSEEN')
    msgs = msgs[0].split()
    for emailid in msgs:
        download(m, emailid, path)
        m.store(emailid, '+FLAGS', '\\Seen')


query()
