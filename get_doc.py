import imaplib
import email

user_mail = 'pythontest0757@outlook.com'
mdp_mail = 'vapcaF-2jukku-wyqtog'
server = 'imap-mail.outlook.com'
outputdir = '/Users/maximethomas/Desktop/Test'


def connect(server, user_mail, mdp_mail):
    m = imaplib.IMAP4_SSL(server)
    m.login(user_mail, mdp_mail)
    m.select()
    return m


def download(m, emailid, outputdir):
    resp, data = m.fetch(emailid, "(BODY.PEEK[])")
    email_body = data[0][1]
    mail = email.message_from_bytes(email_body)
    if mail.get_content_maintype() != 'multipart':
        return
    for part in mail.walk():
        if part.get_content_maintype() != 'multipart' and part.get('Content-Disposition') is not None:
            open(outputdir + '/' + part.get_filename(), 'wb').write(part.get_payload(decode=True))


def query():
    m = connect(server, user_mail, mdp_mail)
    m.select("Inbox")
    typ, msgs = m.search(None, 'UNSEEN')
    msgs = msgs[0].split()
    for emailid in msgs:
        download(m, emailid, outputdir)
        m.store(emailid, '+FLAGS', '\\Seen')


query()
