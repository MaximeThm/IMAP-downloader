# Import all dependencies
import imaplib
import email
import os
import glob
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
import time

# Email credentials
user_mail = 'pythontest0757@outlook.com'
mdp_mail = 'vapcaF-2jukku-wyqtog'

# Portal Credentials
user_portal = "U501171"
mdp_portal = "Factures23!"

# Path to the Chrome WebDriver
driver_file = r'/Users/maximethomas/Documents/chromedriver'
# IMAP server used
server = 'imap-mail.outlook.com'
# Path to the temp folder used
path = '/Users/maximethomas/Desktop/Test/'
#Delay for webpage to upload
delay = 1

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

# Connect to the portal
driver = webdriver.Chrome(driver_file)
def connect_to_portal(user_portal, mdp_portal):
    driver.get("https://portaildepose-ca-lorraine.edokial.com/")
    actions = ActionChains(driver)

    button = driver.find_element_by_id("btnOkCookie")
    actions.move_to_element(button)
    actions.click(button)
    actions.perform()

    username = driver.find_element_by_id("MainContent_LoginUser_UserName")
    username.clear()
    username.send_keys(user_portal)

    password = driver.find_element_by_id("MainContent_LoginUser_Password")
    password.clear()
    password.send_keys(mdp_portal)

    driver.find_element_by_xpath('//*[@id="MainContent_LoginUser_LoginButton"]').click()
    time.sleep(1)

    driver.get("https://portaildepose-ca-lorraine.edokial.com/Page_EchangesDepot.aspx")

#Upload file
def upload_file(filename, path):
    num = int(''.join(list(filter(str.isdigit, filename))))
    dst = path + str(num) + ".pdf"
    os.rename(filename, dst)

    try:
        upload = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_AjaxFileUpload11_Html5InputFile"]')
        upload.send_keys(dst)
    except Exception:
        pass

    driver.find_element_by_xpath('//*[@id="ctl00_MainContent_AjaxFileUpload11_UploadOrCancelButton"]').click()
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(delay)
    button = driver.find_element_by_class_name('btn-colonne-gauche-valider')
    button.click()

    time.sleep(delay)
    driver.find_element_by_xpath(
        '//*[@id="ctl00_MainContent_Aliases1_Repeater1_ctl06_cboConstraints"]/option[2]').click()
    time.sleep(delay)
    driver.find_element_by_xpath(
        '//*[@id="ctl00_MainContent_Aliases1_Repeater1_ctl08_cboConstraints"]/option[3]').click()
    time.sleep(delay)
    driver.find_element_by_xpath('//*[@id="ctl00_MainContent_Aliases1_Repeater1_ctl09_txtAliasQuery"]').send_keys(num)

# Launch all functions
query()
connect_to_portal(user_portal, mdp_portal)
for filename in glob.glob(os.path.join(path, '*.pdf')):
    upload_file(filename, path)


    # X-path pour valider :
    # //*[@id="ctl00_MainContent_btnAliasesValid"]
