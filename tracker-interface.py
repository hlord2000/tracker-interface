import PySimpleGUIQt as sg
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import time
import os
import datetime
import fnmatch
import psutil
from selenium import *
# DEF FUNCTIONS #


def login(user, pwd):
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


def try_find_element(waitElement):
    if len(driver.find_elements_by_xpath(waitElement)) == 0:
        return False
    else:
        return True


def search_case_from_case_num(caseNum):
    driver.get("https://tracker4.pacga.org/search?utf8=%E2%9C%93&search%5Blast_name%5D=&search%5Bfi"
               "rst_name%5D=&search%5Bmiddle_name%5D=&search%5Bcase_role_kwid%5D=&search%5Bda_num%5D=&s"
               "earch%5Bwarrant_num%5D=&search%5Bdocket_num%5D=" + caseNum + "&commit=Search&search%5Baka%5D=&searc"
               "h%5Bssn%5D=&search%5Bdob%5D=&search%5Byob%5D=&search%5Bdl_num%5D=&search%5Bsid_num%5D=&search%"
               "5Bfbi_num%5D=&search%5Bgdc_num%5D=&search%5Bdoc_ef_num%5D=&search%5Bjail_num%5D=&search%5Bstatus_note%"
               "5D=&search%5Boca_num%5D=&search%5Bcrimelab_num%5D=&search%5Bfile_num%5D=&search%5Bcounty_id")


def driver_setup():
    driver.set_window_position(-2000, 0)
    # Install requisite Tracker plugin
    ext_dir = os.path.abspath("trackerhelper5@pacga.org.xpi")
    driver.install_addon(ext_dir, temporary=True)
    # Get tracker homepage
    driver.get("http://tracker4.pacga.org")


def go_to_case(neededValues):
    baseYear = neededValues[0]
    year = baseYear[-2] + baseYear[-1]
    caseNum = neededValues[1]
    fullCaseNum = "ST-" + str(year) + "-CR-" + str(caseNum)
    search_case_from_case_num(fullCaseNum)
    try_click_on_load("/html/body/div[1]/div[2]/div[1]/table/tbody/tr/th/div[2]/table/tbody/tr[3]/td[1]/span/a")


def set_victim_services(textToEnter):
    casePage = driver.current_url
    while True:
        events = window.Read()
        if events[0] == "VSERV":
            driver.switch_to.window(driver.window_handles[0])
            driver.get(casePage + "/victims")
            if driver.find_element_by_xpath("//*[contains(text(), 'No victims.')]"):
                break
            x = 0
            while True:
                try_click_on_load("/html/body/div[1]/div[2]/div[1]/table/tbody/tr/th/table[2]"
                                  "/tbody/tr[1]/td/table/tbody/tr/td[2]/a[1]")
                elem = driver.find_elements_by_xpath("//*[contains(text(), 'Victim:')]")
                print([elements.text for elements in elem])
                try_click_on_load("/html/body/div[1]/div[2]/div[1]/table/tbody/tr/th/table[2]"
                                  "/tbody/tr[1]/td/table/tbody/tr/td[2]/a[1]")
                elem = driver.find_elements_by_xpath("//*[contains(text(), 'Victim:')]")
                elem[x].click()
                try_click_on_load("//*[@id=\"vssr_service_kwid_1212\"]")
                try_click_on_load("//*[@id=\"vssr_service_kwid_1211\"]")
                try_text_on_load("//*[@id=\"victim_service_service_note\"]", textToEnter)
                if x == len(elem) - 1:
                    break
                x += 1
        elif events[0] == "Submit":
            break
        elif events[0] == "Cancel":
            break
        else:
            continue


def letter_setup(textToEnter, pageURL):
    window.Element('VSERV').Update(visible=True)

    # Initialize variables
    go_to_case(values)

    # Now, get the documents page for this case.

    casePage = driver.current_url
    documentsPage = casePage + pageURL

    # Open new tab
    driver.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't')
    # Set new tab as focused tab
    driver.switch_to.window(driver.window_handles[-1])

    # Open documents page
    driver.execute_script("window.open(\'" + documentsPage + "\')")
    set_victim_services(textToEnter)
    driver.switch_to.window(driver.window_handles[0])


def erq_setup(pageURL):
    # Now, get the documents page for this case.

    casePage = driver.current_url
    documentsPage = casePage + pageURL

    # Open new tab
    driver.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 't')
    # Set new tab as focused tab
    driver.switch_to.window(driver.window_handles[-1])
    # Open documents page
    driver.execute_script("window.open(\'" + documentsPage + "\')")
    dates = find_later_date(14)
    print(dates)
    driver.switch_to.window(driver.window_handles[-1])
    time.sleep(1)
    elem = driver.find_element_by_xpath("/html/body/div[1]/div[2]/div[1]/table/tbody/tr/th/form/table/tbody/tr[4]/td[2]/input")
    elem.send_keys(dates)
    time.sleep(.01)
    elem.send_keys(Keys.RETURN)


def find_later_date(distanceInTime):
    currentDate = datetime.datetime.now()
    requestByDate = str(currentDate + datetime.timedelta(days=distanceInTime))
    formattedRequestByDate = requestByDate[5:7]+"/"+requestByDate[8:10]+"/"+requestByDate[0:4]
    return formattedRequestByDate

# END DEF FUNCTIONS #


# Initialize the browser
driver = webdriver.Firefox(executable_path='geckodriver.exe')
driver_setup()

# Defines the style for submit and cancel buttons in this program.
submitButton = [
    [sg.Button(button_text="Submit",
               auto_size_button=True,
               button_color=('#FFFFFF', '#483D8B'),
               focus=True)
     ]
]
cancelButton = [
    [sg.Button(button_text="Cancel",
               auto_size_button=True,
               button_color=('#FFFFFF',
                             '#483D8B')
               )
     ]
]

# Establishes layout for Login form.
loginLayout = [[sg.Text('Tracker Interface - Login')],
               [sg.Text('Username')],
               [sg.InputText()],
               [sg.Text('Password')],
               [sg.InputText(password_char="*")],  # password_char hides text with *
               [sg.Column(submitButton),
                sg.Column(cancelButton)]
               ]

# Names window as 'Tracker Interface - Login' and assigns its layout to loginLayout.
window = sg.Window('Tracker Interface - Login', icon="SEAL.jpg").Layout(loginLayout)

# Reads button events and values entered, stores them in variables.
event, values = window.Read()

# If cancel button is pressed, the program is closed.
if event == "Cancel":
    print("THIS SHOULD CLOSE")
    window.Close()
    driver.close()
    exit()
elif event == "Submit":
    window.Close()
# Store username and password variables for later use in login.

username = values[0]
password = values[1]

# Login.
login(username, password)

columnTop = [[sg.Text('Tracker Interface - Case Search')],
             [sg.Text('Year: '), sg.Spin([i for i in range(2000, 2100)], initial_value=2019)],
             [sg.Text('Case Number: '), sg.InputText()]]

jobMakerLayout = [
    [sg.Column(columnTop)],
    [sg.Text('_' * 100)],
    [sg.Text('Case Options:')],
    [sg.Radio('Display Case', "RADIO1", default=True),
     sg.Radio('CC-JT Letter Creator', "RADIO1"),
     sg.Radio('CC Letter Creator', "RADIO1"),
     sg.Radio('Evidence Request Creator (WIP)', "RADIO1")],
    [sg.Radio('Case Status/Change of Plea Updater', "RADIO1")],
    [sg.Column(submitButton),
     sg.Button(button_text="Add victim services?",
               button_color=('#FFFFFF', '#7766E8'),
               visible=False, key='VSERV'),
     sg.Column(cancelButton)]
]

popupOnError = [
    []
]

window = sg.Window('Tracker Interface - Case Search', icon="SEAL.jpg").Layout(jobMakerLayout)
while True:
    event, values = window.Read()
    driver.switch_to.window(driver.window_handles[0])
    driver.get("http://tracker4.pacga.org")
    if event[0] == "C":
        window.Close()
        driver.quit()
        exit()
    # Makes button invisible again.
    print(values)
    # Brings web browser back into view.
    driver.set_window_position(0, 0)
    if values[2]:
        go_to_case(values)
    elif values[3]:
        letter_setup("CC-JT letter sent.", "/documents/new?document_template_id=20426")
    elif values[4]:
        letter_setup("CC letter sent.", "/documents/new?document_template_id=20787")
    elif values[5]:
        go_to_case(values)
        if try_find_element("//*[contains(text(), 'Evidence Request Form 9-5-17')]"):
            sg.Popup("Evidence request already created.",
                     title='Tracker Interface - Case Search',
                     icon="SEAL.jpg",
                     keep_on_top=True,
                     line_width=40,

                     )
            continue
        else:
            erq_setup("/documents/new?document_template_id=17572")

            # Click on word icon.
            try_click_on_load("/html/body/div[1]/div[2]/div[1]/table/tbody/tr/th/table/tbody/tr[2]/td[2]/span/img")

            # Wait for launch of office program.
            time.sleep(1)
            # Kill program.
            try:
                os.system("TASKKILL /F /IM soffice.bin")
            except:
                pass
            try:
                os.system("TASKKILL /F /IM winword.exe")
            except:
                pass

            for file in os.listdir(os.path.expandvars(r'%LOCALAPPDATA%\Temp\trackerhelper')):

                # Record which file is being operated on.
                print(file)

                # Gives absolute path for file.
                fileDir = os.path.expandvars(r'%LOCALAPPDATA%\Temp\trackerhelper')+ "\\" + file

                # Checks if file selected is rtf.
                if fnmatch.fnmatch(file, '*.rtf'):
                    print("Done once.")
                    # Selects rtf file.
                    fileToOpen = open(fileDir)

                    # Reads rtf file.
                    fileData = fileToOpen.read()
                    # Closes IO stream
                    fileToOpen.close()
                    time.sleep(1)
                    # Deletes file.
                    os.remove(fileDir)

                    # Opens criteria for search.
                    criteria = open("searchCriteriaEvidenceRequest.txt", 'r')
                    criteriaData = criteria.read()
                    criteria.close()

                    # Opens the search criterion's replacement
                    replacement = open('replaceEvidenceRequest.txt', 'r')
                    replacementData = replacement.read()
                    replacement.close()

                    # Replaces the data.
                    newData = fileData.replace(criteriaData, replacementData)
                    print(newData)

                    # Write data to file with same name.
                    f = open(fileDir, 'x')
                    f.write(newData)
                    f.close()
                    while not try_find_element("//*[@id=\"save_button\"]"):
                        elem = driver.find_element_by_xpath("//*[@id=\"save_button\"]")
                        elem.click()

                else:
                    print('File not .rtf.')
                    pass

            driver.switch_to.window(driver.window_handles[0])
            try_click_on_load("/html/body/div[1]/div[2]/div[1]/table/tbody/tr/th/div[8]/"
                              "table/tbody/tr/td/table/tbody/tr/td[2]/a[3]")
            try_text_on_load("//*[@id=\"action_event_event_date\"]", find_later_date(14))

            # Select dropdown options.
            while not try_find_element('//*[@id=\"action_event_event_type_kwid\"]'):
                elem = Select(driver.find_element_by_xpath('//*[@id=\"action_event_event_type_kwid\"]'))
                elem.select_by_value('2722')
            while not try_find_element('//*[@id="action_event_staff_id"]'):
                elem = Select(driver.find_element_by_xpath('//*[@id=\"action_event_staff_id\"]'))
                elem.select_by_value('5801262')

            # Save action item.
            try_click_on_load('/html/body/div[1]/div[2]/div[1]/table/tbody/tr/th/div[5]/form/table/tbody/tr/td/input')

    elif values[6]:
        letter_setup("Case Status letter sent.", "/documents/new?document_template_id=19308")
    else:
        print("ERROR NO RADIO BUTTON SELECTED")
        exit()
