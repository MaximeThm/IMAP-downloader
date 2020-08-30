from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
import time

user_portal = "user name"
mdp_portal = "password!"
driver_file = "where the Chrome driver is located'
outputdir = "path to source folder"

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


connect_to_portal(user_portal, mdp_portal)
try:
    upload = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_AjaxFileUpload11_Html5InputFile"]')
    upload.send_keys(outputdir)
except Exception:
    pass

driver.find_element_by_xpath('//*[@id="ctl00_MainContent_AjaxFileUpload11_UploadOrCancelButton"]').click()
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
time.sleep(1)
button = driver.find_element_by_class_name('btn-colonne-gauche-valider')
button.click()

time.sleep(1)
driver.find_element_by_xpath('//*[@id="ctl00_MainContent_Aliases1_Repeater1_ctl06_cboConstraints"]/option[2]').click()
time.sleep(1)
driver.find_element_by_xpath('//*[@id="ctl00_MainContent_Aliases1_Repeater1_ctl08_cboConstraints"]/option[3]').click()
time.sleep(1)
driver.find_element_by_xpath('//*[@id="ctl00_MainContent_Aliases1_Repeater1_ctl09_txtAliasQuery"]').send_keys('123456789')

# X-path pour valider :
# //*[@id="ctl00_MainContent_btnAliasesValid"]
