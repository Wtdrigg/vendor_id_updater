from time import sleep
import clipboard
from xpath_map import Map
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service as ChromeService
from subprocess import CREATE_NO_WINDOW


class Updater:
    
    # The Constructor creates all class attributes, some are set as placeholder values and are updated later.
    # of note is the self.map attribute which is an object created from the Map class, this object contains
    # all the Xpath information used to indentify what boxes to interact with in the web browser.
    def __init__(self):
        self.driver = None
        self.actions = None
        self.map = Map()
        self.supplier_id = ''
        self.vc_id = ''
        self.counter = 1

    # The Chromedriver_setup() method creates the chromedriver object and specifies the exe path for chromedriver.exe.
    # It also specifies the path for your browser User_data, a copy of which is needed to run the chrome with your
    # normal profile. It also creates a new object using the ChromeService class, which is used to have chromedriver run
    # without a CMD window being opened.
    @staticmethod
    def chromedriver_setup():
        print('Starting Webdriver')
        my_options = webdriver.ChromeOptions()
        my_options.add_argument('user-data-dir=C:/Users/Tyler/programingprojects/vendor_id_updater/User Data')
        chrome_service = ChromeService('C:/Users/Tyler/programingprojects/vendor_id_updater/chromedriver.exe')
        chrome_service.creationflags = CREATE_NO_WINDOW
        driver_Setup = webdriver.Chrome(options=my_options, 
                                        executable_path='C:/Users/Tyler/programingprojects/vendor_id_updater'
                                                        '/chromedriver.exe',
                                        service=chrome_service)
        return driver_Setup

    # The actionchains_setup() method creates an ActionChains object. This is used later to input keystroke combinations
    # into the webdriver.
    def actionchains_setup(self):
        actions = ActionChains(self.driver)
        return actions

    # The clipboard_copy() method copies any arguments provided into the clipboard. This is useful for quickly getting
    # the entire text from a website.
    @staticmethod
    def clipboard_copy(text):
        results = clipboard.copy(text)
        return results

    # The clipboard_paste() method takes anything copied into your clipboard and returns it as a string.
    @staticmethod
    def clipboard_paste():
        results = clipboard.paste()
        return results

    # The open_vcommerce() method has the webdriver object open the Vendor ID update list in Vcommerce and then wait
    # for the list to load.
    def open_vcommerce(self):
        print('Opening Vcommerce')
        self.driver.get('https://solutions.sciquest.com/apps/Router/ApprNotifications?SelectedTab'
                        '=ApprNotificationSupplierRegistration&AuthUser=6637615&OrgName=Vulcan&tmstmp=1661890913947')
        sleep(3)

        # Sometimes Vcommerce will go directly to the page without going to the login screen first. The following
        # logic checks to see if the login page has loaded, and will log in if so. If the login page has not loaded
        # it will check to see if the Vcommerce search results have loaded and continue on if so (and raise a timeout
        # exception if not).
        login_verify = self.driver.find_elements(By.XPATH, self.map.vc_xpath['login_button'])
        if login_verify:
            login_button_element = self.driver.find_element(By.XPATH, self.map.vc_xpath['login_button'])
            login_button_element.click()
            WebDriverWait(self.driver, 20).until(ec.presence_of_element_located((By.XPATH, self.map.vc_xpath[
                'search_results'])))
        else:
            WebDriverWait(self.driver, 20).until(ec.presence_of_element_located((By.XPATH, self.map.vc_xpath[
                'search_results'])))

    # The open_riskonnect() method has the driver object open riskonnect in a second tab, then switch to that tab
    # and log in. Like the previous method, this contains logic to log in to Riskonnect if prompted, or continue on
    # if not prompted.
    def open_riskonnect(self):
        print('Opening Riskonnect\n')
        self.driver.execute_script("window.open('https://riskonnectvmc.lightning.force.com/lightning/page/home');")
        self.driver.switch_to.window(self.driver.window_handles[1])
        sleep(3)
        login_verify = self.driver.find_elements(By.XPATH, self.map.rk_xpath['login_button'])
        if login_verify:
            login_button_element = self.driver.find_element(By.XPATH, self.map.rk_xpath['login_button'])
            login_button_element.click()
        else:
            pass

    # The switch_tab_to_1() method has the driver object switch to the first browser tab, which has Vcommerce loaded.
    def switch_tab_to_1(self):
        self.driver.switch_to.window(self.driver.window_handles[0])

    # The switch_tab_to_2() method has the driver object switch to the second browser tab, which has Riskonnect loaded.
    def switch_tab_to_2(self):
        self.driver.switch_to.window(self.driver.window_handles[1])

    # The switch_tab_to_3() method has the driver object switch to the third browser tab, which has the Vcommerce
    # vendor summary page loaded.
    def switch_tab_to_3(self):
        self.driver.switch_to.window(self.driver.window_handles[2])

    # The close_current_tab method closes the currently open tab.
    def close_current_tab(self):
        self.driver.close()
        
    # The vc_open_summary() method has the webdriver select and open the first vendor link on the Vcommerce
    # vendor list in a new tab.
    def vc_open_summary(self):
        print(f'Accessing vendor {self.counter}')
        item_element = self.driver.find_element(By.XPATH, self.map.vc_xpath['first_item'])
        self.actions.key_down(Keys.CONTROL).perform()
        self.actions.click(item_element).perform()
        self.actions.key_up(Keys.CONTROL).perform()

    # The vc_access_summary() method has the driver switch to the Vcommerce summary page and wait for it to load.
    # It will then copy the entire page to the clipboard using Control+A to select all and Control+C to copy.
    def vc_access_summary(self):
        print('Accessing summary')
        vc_summary_element = self.driver.find_element(By.XPATH, self.map.vc_xpath['summary_button'])
        vc_summary_element.click()
        WebDriverWait(self.driver, 20).until(ec.presence_of_element_located((By.XPATH, self.map.vc_xpath['summary_page'
                                                                                                         '_body'])))
        vc_summary_body_element = self.driver.find_element(By.XPATH, self.map.vc_xpath['summary_page_body'])
        vc_summary_body_element.send_keys(Keys.CONTROL + 'a')
        vc_summary_body_element.send_keys(Keys.CONTROL + 'c')

    # The parse_summary() method takes the copied data from the vc_summary and locates the supplier ID number and
    # Vcommerce ID numbers, which are then saved as class attributes.
    def parse_summary(self):
        print('Reading vendor summary')
        # Takes the clipboard data as a giant string
        summary_text = self.clipboard_paste()
        # creates two lists from the data by splitting on whitespace
        summary_list = summary_text.split()
        summary_list_2 = summary_text.split()
        # looks through the first list for the string 'Number'. When found it takes the iteration count provided by
        # enumerate() and checks if the list item at the prior iteration (count - 1) is the string 'Supplier'. If so,
        # that means that the list item at the next iteration (count + 1) is the supplier number and is saved as
        # the supplier_id class attribute
        for count, item in enumerate(summary_list):
            if item == 'Number':
                if summary_list[count - 1] == 'Supplier':
                    self.supplier_id = summary_list[count + 1]
                    break
                else:
                    summary_list.remove(item)
                    continue
        # looks through the second list for the string 'ID'. When found, this means the next iteration in the list
        # is the Vcommerce ID number, so we take the iteration count provided by enumerate() and then add 1 to it to
        # get the list index position of the next iteration, which contains the number we are looking for. We then
        # save the list item at that index position as the vc_id class attribute
        for count, item in enumerate(summary_list_2):
            if item == 'ID':
                self.vc_id = summary_list_2[count + 1]
                break
    
    # The rk_search_for_vc_id() method searches Riskonnect for the Vcommerce number and opens the matching
    # item in a new tab. Action chains are used to open that item in a new tab by simulating a Control+click.
    def rk_search_for_vc_id(self):
        print('Searching for vendor')
        sleep(0.5)
        search_button_element = self.driver.find_element(By.XPATH, self.map.rk_xpath['search_button'])
        search_button_element.click()
        sleep(0.5)
        search_bar_element = self.driver.find_element(By.XPATH, self.map.rk_xpath['search_bar'])
        search_bar_element.click()
        search_bar_element.clear()
        search_bar_element.send_keys(self.vc_id)
        sleep(0.5)
        search_bar_element.send_keys(Keys.ENTER)
        WebDriverWait(self.driver, 20).until(ec.presence_of_element_located((By.XPATH, self.map.rk_xpath['search_result'
                                                                                                         '_first'])))
        search_result_element = self.driver.find_element(By.XPATH, self.map.rk_xpath['search_result_first'])
        self.actions.key_down(Keys.CONTROL).perform()
        self.actions.click(search_result_element).perform()
        self.actions.key_up(Keys.CONTROL).perform()
    
    # The rk_update_supplier_id() method opens the supplier ID box in Riskonnect, clears the old value, and replaces it
    # with the value saved as the supplier ID attribute. If riskonnect indicates that it cannot save due a duplicate
    # already existing with that number, then the Vcommerce and supplier ID numbers are both saved to a text file
    # for manual review
    def rk_update_supplier_id(self):
        print('Updating supplier ID')
        # locates the edit button on the page then clicks and waits half a second.
        WebDriverWait(self.driver, 20).until(ec.presence_of_element_located((By.XPATH, self.map.rk_xpath['edit_'
                                                                                                         'button'])))
        edit_button_element = self.driver.find_element(By. XPATH, self.map.rk_xpath['edit_button'])
        edit_button_element.click()
        sleep(0.5)
        # locates the supplier id box on the page
        supplier_id_box = self.driver.find_element(By.XPATH, self.map.rk_xpath['supplier_id_box'])
        # clears the supplier id box and waits half a second
        supplier_id_box.clear()
        sleep(0.5)
        # enters the saved supplier id into the box and waits half a second
        supplier_id_box.send_keys(self.supplier_id)
        sleep(0.5)
        # locates and clicks the save button, then waits a full second.
        save_button_element = self.driver.find_element(By.XPATH, self.map.rk_xpath['save_button'])
        save_button_element.click()
        sleep(1)
        # checks to see if the duplicate warning has appeared.
        duplicate_check = self.driver.find_elements(By.XPATH, self.map.rk_xpath['duplicate_alert'])
        # if the warning is found, this saves the numbers in a text file.
        if duplicate_check:
            f = open('C:/Users/driggerst/Python/vendor_id_updater/duplicates_list.txt', 'a')
            f.write(f'vcommerce ID:{self.vc_id} did not accept supplier ID:{self.supplier_id} due to a duplicate\n')
            f.close()
        # if the warning is not found, no further action is taken.
        else:
            pass

    # The rk_return_home() method has the driver return Riskonnect to its home screen.
    def rk_return_home(self):
        print('Resetting Riskonnect')
        sleep(0.5)
        WebDriverWait(self.driver, 20).until(ec.presence_of_element_located((By.XPATH, self.map.rk_xpath['home_'
                                                                                                         'button'])))
        home_button_element = self.driver.find_element(By.XPATH, self.map.rk_xpath['home_button'])
        home_button_element.click()

    # The vc_remove_item() method has the driver click the 'remove item' button on the vendor list once the update
    # is completed.
    def vc_remove_item(self):
        print('Removing vendor')
        remove_button_element = self.driver.find_element(By.XPATH, self.map.vc_xpath['remove_button'])
        remove_button_element.click()

    # The wait_for_remove() method has the driver wait until the item has been removed from the Vcommerce list
    # before moving forward.
    def wait_for_remove(self):
        sleep(2)
        WebDriverWait(self.driver, 20).until(ec.presence_of_element_located((By.XPATH, self.map.vc_xpath['remove_'
                                                                                                         'button'])))
        print(f'Vendor {self.counter} updated\n')
        self.counter += 1

    # The prep_update() method runs both the open_vcommerce() method and the open_riskonnect() methods.
    # This is for use in the __main__ script and GUI.
    def prep_update(self):
        self.open_vcommerce()
        self.open_riskonnect()

    # The process_update() method runs all methods in the order required to process the vendor update. This is for
    # use in the __main__ script and GUI.
    def process_update(self):
        self.switch_tab_to_1()
        self.vc_open_summary()
        self.switch_tab_to_3()
        self.vc_access_summary()
        self.close_current_tab()
        self.parse_summary()
        self.switch_tab_to_2()
        self.rk_search_for_vc_id()
        self.switch_tab_to_3()
        self.rk_update_supplier_id()
        self.close_current_tab()
        self.switch_tab_to_2()
        self.rk_return_home()
        self.switch_tab_to_1()
        self.vc_remove_item()
        self.wait_for_remove()
