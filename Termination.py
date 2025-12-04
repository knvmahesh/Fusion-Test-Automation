from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
import time
import traceback
import XLutils


# Helper wrappers to provide clearer errors per action
def safe_click(driver, by, locator, timeout=15, desc=None):
    desc = desc or f"click {locator}"
    try:
        elem = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, locator)))
        elem.click()
    except Exception as e:
        raise Exception(f"Failed to {desc}: {e}")


def safe_send_keys(driver, by, locator, text, timeout=15, desc=None):
    desc = desc or f"send keys to {locator}"
    try:
        elem = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, locator)))
        elem.clear()
        elem.send_keys(text)
    except Exception as e:
        raise Exception(f"Failed to {desc}: {e}")


def main():
    # ------------------------------------------------------------
    #  LAUNCH BROWSER
    # ------------------------------------------------------------
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
    driver.maximize_window()

    # ------------------------------------------------------------
    #  EXCEL PATH
    # ------------------------------------------------------------
    path = r"C:\Users\velur\Desktop\Selenium-Python\Termination.xlsx"
    rows_Termination = XLutils.getRowCount(path, 'Termination')

    # ------------------------------------------------------------
    #  LOGIN
    # ------------------------------------------------------------
    try:
        Base_url = XLutils.readData(path, "URL", 2, 1)
        driver.get(Base_url)

        UN = XLutils.readData(path, "LD", 2, 1)
        PW = XLutils.readData(path, "LD", 2, 2)

        driver.find_element(By.ID, "userid").send_keys(UN)
        driver.find_element(By.ID, "password").send_keys(PW)
        driver.find_element(By.ID, "btnActive").click()
        
        # - clicking on home button
        xpath_home= '//a[@id="pt1:_UIShome"]/*[name()="svg"]/*[name()="g"][4]/*[name()="path"]'
        safe_click(driver, By.XPATH, xpath_home, 15, desc="click Home")
        # - clicking on my client groups
        xpath_MCG = '//a[@id="groupNode_workforce_management"]'
        safe_click(driver, By.XPATH, xpath_MCG, 15, desc="click My Client Groups")
        # - clicking on person Managment button
        xpath_MCG = '//a[@id="itemNode_workforce_management_person_management_0"]'
        safe_click(driver, By.XPATH, xpath_MCG, 15, desc="click Person Management")

    except Exception as e:
        # For login failure, write to row 2 (login row) so it's visible
        XLutils.writeData(path, "Termination", 2, 6, f"LOGIN FAILED: {str(e)}")
        driver.quit()
        return

    # ------------------------------------------------------------
    #  LOOP FOR EACH PERSON
    # ------------------------------------------------------------
    for r in range(2, rows_Termination + 1):

        person = XLutils.readData(path, 'Termination', r, 1)

        print("\n-------------------------------")
        print(f"Processing Person: {person}")
        print("-------------------------------")

        try:
            # SEARCH PERSON
            xpath_PersonNumber = "//*[@id='_FOpt1:_FOr1:0:_FONSr2:0:MAt1:0:pt1:Perso1:0:SP3:q1:value10::content']"
            input_box = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, xpath_PersonNumber))
            )
            input_box.clear()
            input_box.send_keys(person + Keys.ENTER)

            time.sleep(3)

            safe_click(driver, By.XPATH, "//button[text()='Search']", 10, desc="click Search")

            # ACTIONS → PERSON & EMPLOYMENT
            safe_click(driver, By.XPATH, "//button[@title='Actions']", 10, desc="click Actions")

            xpath_person = '//*[@id="_FOpt1:_FOr1:0:_FONSr2:0:MAt1:0:pt1:Perso1:0:SP3:table1:am2:dc_i1:3:dcm1"]/td[2]'
            safe_click(driver, By.XPATH, xpath_person, 10, desc="select Person & Employment")

            # WORK RELATIONSHIP
            xpath_work = '//*[@id="_FOpt1:_FOr1:0:_FONSr2:0:MAt1:0:pt1:Perso1:0:SP3:table1:am2:dc_i1:3:dci1:12:dccmi1"]/td[2]'
            safe_click(driver, By.XPATH, xpath_work, 10, desc="select Work Relationship")

            # ACTION → TERMINATE
            safe_click(driver, By.XPATH, "//a[@title='Actions']", 10, desc="click Actions (work relationship)")

            safe_click(driver, By.XPATH, "//td[text()='Terminate']", 10, desc="click Terminate")

            # ACTION DETAILS
            excel_action = XLutils.readData(path, 'Termination', r, 2)
            actionDD_xpath = "//*[contains(@id, 'Action::content')]"
            safe_send_keys(driver, By.XPATH, actionDD_xpath, excel_action + Keys.ENTER, 10, desc="set Action Details")

            # NOTIFICATION DATE
            excel_notif = XLutils.readData(path, 'Termination', r, 3)
            notif_date = excel_notif.strftime('%d-%b-%Y')

            notifField_xpath = "//input[@aria-label='Notification Date']"
            safe_send_keys(driver, By.XPATH, notifField_xpath, notif_date + Keys.ENTER, 10, desc="set Notification Date")

            # TERMINATION DATE
            excel_termi = XLutils.readData(path, 'Termination', r, 4)
            termi_date = excel_termi.strftime('%d-%b-%Y')

            termField_xpath = "//input[@aria-label='Termination Date']"
            safe_send_keys(driver, By.XPATH, termField_xpath, termi_date + Keys.ENTER, 10, desc="set Termination Date")

            # REHIRE
            excel_rehire = XLutils.readData(path, 'Termination', r, 5)
            rehireField_xpath = "//*[contains(@id, 'RehireRecom::content')]"
            safe_send_keys(driver, By.XPATH, rehireField_xpath, excel_rehire + Keys.ENTER, 10, desc="set Rehire")

            # REVIEW
            safe_click(driver, By.XPATH, "//button[text()='Review']", 10, desc="click Review")

            time.sleep(5)

            # SUBMIT
            safe_click(driver, By.XPATH, "//span[normalize-space(.)='Submit']", 10, desc="click Submit")

            time.sleep(3)

            # YES
            safe_click(driver, By.XPATH, "//button[@accesskey='Y']", 10, desc="click Yes")

            # OK
            safe_click(driver, By.XPATH, "//button[@accesskey='K']", 10, desc="click OK")

            # CLOSE
            safe_click(driver, By.XPATH, "//button[contains(., 'Close')]", 10, desc="click Close")

            # ---------------------------------------------------
            # WRITE SUCCESS MESSAGE TO THE ROW FOR THIS PERSON
            # ---------------------------------------------------
            XLutils.writeData(path, "Termination", r, 6, f"SUCCESS – Person {person} Terminated")

            print(f"✅ SUCCESS – Person {person}")

        except Exception as e:
            # write error to the same row so failures are visible per-person
            tb = traceback.format_exc()
            XLutils.writeData(path, "Termination", r, 6, f"FAILED – Person {person} – {str(e)}")
            print(f"❌ FAILED – Person {person} – {str(e)}")
            print(tb)
            # continue to next row
            continue

    driver.quit()
    print("\n✔ COMPLETED ALL ROWS")


if __name__ == '__main__':
    main()
