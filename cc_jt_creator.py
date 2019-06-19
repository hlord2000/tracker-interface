import PySimpleGUI as sg
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
import os
import calendar

# DEF FUNCTIONS #


def login(user, pwd):
    # Login.
    try_click_on_load("//*[@id=\"memo\"]/p[3]/a")
    try_text_on_load("//*[@id=\"user_login\"]", user)
    try_text_on_load("//*[@id=\"user_password\"]", pwd)
    driver.execute_script("window.alert = function() {};")


def try_link_follow(linkFollow, waitElement):
    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, waitElement))
        )
    except:
        print("FAILED TO FOLLOW LINK")
    finally:
        elem = driver.find_element_by_xpath(waitElement)
        driver.get(linkFollow)


def try_text_on_load(targetElement, textToBePassed):
    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, targetElement))
        )
    except:
        print("FAILED TO ENTER TEXT")
    finally:
        elem = driver.find_element_by_xpath(targetElement)
        elem.clear()
        elem.send_keys(textToBePassed)
        time.sleep(.01)
        elem.send_keys(Keys.RETURN)


def try_click_on_load(waitElement):
    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, waitElement))
        )
    except:
        print("FAILED TO CLICK")
    finally:
        elem = driver.find_element_by_xpath(waitElement)
        elem.click()

def search_case_from_case_num(caseNum):
    driver.get("https://tracker4.pacga.org/search?utf8=%E2%9C%93&search%5Blast_name%5D=&search%5Bfi"
               "rst_name%5D=&search%5Bmiddle_name%5D=&search%5Bcase_role_kwid%5D=&search%5Bda_num%5D=&s"
               "earch%5Bwarrant_num%5D=&search%5Bdocket_num%5D=" + caseNum + "&commit=Search&search%5Baka%5D=&searc"
               "h%5Bssn%5D=&search%5Bdob%5D=&search%5Byob%5D=&search%5Bdl_num%5D=&search%5Bsid_num%5D=&search%"
               "5Bfbi_num%5D=&search%5Bgdc_num%5D=&search%5Bdoc_ef_num%5D=&search%5Bjail_num%5D=&search%5Bstatus_note%"
               "5D=&search%5Boca_num%5D=&search%5Bcrimelab_num%5D=&search%5Bfile_num%5D=&search%5Bcounty_id")

# DEF FUNCTIONS #

# Get information from user.

# Defines the style for submit and cancel buttons in this program.
submitButton = [
    [sg.Button(button_text="Submit", auto_size_button=True, button_color=('#FFFFFF', '#483D8B'), focus=True)]]
cancelButton = [
    [sg.Button(button_text="Cancel", auto_size_button=True, button_color=('#FFFFFF', '#483D8B'))]]

# Establishes layout for Login form.
loginLayout = [[sg.Text('Tracker Interface - Login')],
               [sg.Text('Username')],
               [sg.InputText(size=(25, 2))],
               [sg.Text('Password')],
               [sg.InputText(password_char="*", size=(25, 2))],  # password_char hides text with *
               [sg.Column(submitButton), sg.Column(cancelButton)]
               ]

# Names window as 'Tracker Interface - Login' and assigns its layout to loginLayout.
window = sg.Window('Tracker Interface - Login').Layout(loginLayout)
# Reads button events and values entered, stores them in variables.
event, values = window.Read()
# If cancel button is pressed, the program is closed.
if event == "Cancel":
    print("THIS SHOULD CLOSE")
    window.Close()
elif event == "Submit":
    window.Close()
# Store username and password variables for later use in login.

username = values[0]
password = values[1]

# Initialize and identify webdriver.  C:\geckodriver.exe
driver = webdriver.Firefox(executable_path='geckodriver.exe')
driver.set_window_position(-2000, 0)
# Install requisite Tracker plugin
ext_dir = os.path.abspath("trackerhelper5@pacga.org.xpi")
driver.install_addon(ext_dir, temporary=True)
# Get tracker homepage
driver.get("http://tracker4.pacga.org")
# Call Login Function
login(username, password)

columnTop = [[sg.Text('Tracker Interface - Case Search')],
             [sg.Text('Year:              '), sg.Spin([i for i in range(2000, 2100)], initial_value=2019)],
             [sg.Text('Case Number: '), sg.InputText(size=(4, 2))]]

jobMakerLayout = [
    [sg.Column(columnTop)],
    [sg.Text('_' * 75, size=(65, 1))],
    [sg.Text('Case Options:')],
    [sg.Radio('CC-JT Letter Creator', "RADIO1", default=True), sg.Radio('CC Letter Creator', "RADIO1"),
     sg.Radio('Evidence Request Creator (WIP)', "RADIO1")],
    [sg.Column(submitButton), sg.Column(cancelButton)]
]

window = sg.Window('Tracker Interface - Case Search').Layout(jobMakerLayout)
while True:
    event, values = window.Read()
    # CREATE CALLS FOR VARIOUS FUNCTIONS HERE
    driver.set_window_position(0, 0)
    if values[2]:
        ccJTLoop = True
        while ccJTLoop:
            # Initialize variables

            baseYear = values[0]
            year = baseYear[-2]+baseYear[-1]
            caseNum = values[1]
            fullCaseNum = ""
            numJTDates = ""
            jtDateSecond = ""
            # Construct case number
            fullCaseNum = "ST-" + str(year) + "-CR-" + str(caseNum)

            # Try to enter case number

            search_case_from_case_num(fullCaseNum)

            # Click on case.

            try_click_on_load("/html/body/div[1]/div[2]/div[1]/table/tbody/tr/th/div[2]/table/tbody/tr[3]/td[1]/span/a")

            # Now, get the documents page for this case.
            casePage = driver.current_url
            documentsPage = casePage + "/documents/new?document_template_id=20426"

            # Open new tab

            driver.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't')
            # Set new tab as focused tab
            driver.switch_to.window(driver.window_handles[-1])

            # Open documents page
            driver.execute_script("window.open(\'" + documentsPage + "\')")

            if int(numJTDates) == 2:
                date = calendar.month_name[int(jtDateSecond[0] + jtDateSecond[1])] + " " + jtDateSecond[3] + \
                       jtDateSecond[
                           4] + ", " + jtDateSecond[6:10]
                print(date)
                driver.switch_to.window(driver.window_handles[1])
                try_text_on_load("//*[@id=\"parameter_32313\"]", date)
            # Switch back to first tab after

            while True:
                if input("Set victim services?").upper() == "Y":
                    driver.switch_to.window(driver.window_handles[0])
                    driver.get(casePage + "/victims")
                    try_click_on_load(
                        "/html/body/div[1]/div[2]/div[1]/table/tbody/tr/th/table[2]/tbody/tr[1]/td/table/tbody/tr/td[2]/a[1]")
                    try_click_on_load("//*[@id=\"vssr_service_kwid_1212\"]")
                    try_click_on_load("//*[@id=\"vssr_service_kwid_1211\"]")
                    try_text_on_load("//*[@id=\"victim_service_service_note\"]", "CC-JT letter sent.")
                    break
                else:
                    break

            driver.switch_to.window(driver.window_handles[0])

            while True:
                if input("More letters? (Y/N)").upper() == "Y":
                    break
                elif input("More letters? (Y/N)").upper() == "N":
                    continueLoop = False
                    break
                else:
                    print("Please enter a valid response.")
    elif values[3]:
        # CALL CC LOOP
        print("b")
    elif values[4]:
        print("a")
        #CALL EVIDENCE REQ LOOP
    if event == "Cancel":
        exit()

# Loop for asking for info

