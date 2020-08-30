import imaplib
import email

user_mail = 'your email'
mdp_mail = 'your email password'
server = 'imap-mail.outlook.com' #This is an example. See https://support.microsoft.com/fr-fr/office/paramètres-de-messagerie-pop-et-imap-pour-outlook-8361e398-8af4-4e97-b147-6c6c4ac95353 for more information
outputdir = 'your output directory'


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
