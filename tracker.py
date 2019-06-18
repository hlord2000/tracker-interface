from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import os
import calendar

def init_browser():
    global driver
    # Initialize and identify webdriver.  C:\geckodriver.exe
    driver = webdriver.Firefox(executable_path='geckodriver.exe')
    # Install requisite Tracker plugin
    driver.install_addon(os.path.abspath("trackerhelper5@pacga.org.xpi"), temporary=True)
def login(user,pwd):
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

def cc_jt_loop():
    # Loop for asking for info
    continueLoop = True
    while continueLoop:
        # Initialize variables

        year = ""
        caseNum = ""
        fullCaseNum = ""
        numJTDates = ""
        jtDateSecond = ""

        year = input('Please enter year of case (e.g. 18, 19, etc.)...')
        caseNum = input('Please enter case number (e.g. 1234)...')

        while True:
            numJTDates = input('How many Jury Trial dates are included?')
            if int(numJTDates) == 2:
                jtDateSecond = input('Please enter the date of second Jury Trial (e.g. mm/dd/yyyy)...')
                break
            elif int(numJTDates) == 1:
                break
            else:
                print("Please enter a valid number.")
                continue

        # Construct case number
        fullCaseNum = "ST-" + year + "-CR-" + caseNum

        # Try to enter case number

        driver.get("https://tracker4.pacga.org/search?utf8=%E2%9C%93&search%5Blast_name%5D=&search%5Bfi"
                   "rst_name%5D=&search%5Bmiddle_name%5D=&search%5Bcase_role_kwid%5D=&search%5Bda_num%5D=&s"
                   "earch%5Bwarrant_num%5D=&search%5Bdocket_num%5D=" + fullCaseNum + "&commit=Search&search%5Baka%5D=&searc"
                                                                                     "h%5Bssn%5D=&search%5Bdob%5D=&search%5Byob%5D=&search%5Bdl_num%5D=&search%5Bsid_num%5D=&search%"
                                                                                     "5Bfbi_num%5D=&search%5Bgdc_num%5D=&search%5Bdoc_ef_num%5D=&search%5Bjail_num%5D=&search%5Bstatus_note%"
                                                                                     "5D=&search%5Boca_num%5D=&search%5Bcrimelab_num%5D=&search%5Bfile_num%5D=&search%5Bcounty_id%5D=&log_"
                                                                                     "current_user=hlord")
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
            date = calendar.month_name[int(jtDateSecond[0] + jtDateSecond[1])] + " " + jtDateSecond[3] + jtDateSecond[
                4] + ", " + jtDateSecond[6:10]
            print(date)
            driver.switch_to.window(driver.window_handles[1])
            try_text_on_load("//*[@id=\"parameter_32313\"]", date)

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

        while True:
            if input("More letters? (Y/N)").upper() == "Y":
                driver.switch_to.window(driver.window_handles[0])
                break
            elif input("More letters? (Y/N)").upper() == "N":
                continueLoop = False
                break
            else:
                print("Please enter a valid response.")