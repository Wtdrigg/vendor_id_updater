from time import sleep
import clipboard
from xpath_map import Map
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.action_chains import ActionChains
from tkinter import *
from tkinter.ttk import *
from selenium.webdriver.chrome.service import Service as ChromeService
from subprocess import CREATE_NO_WINDOW


class Updater:
    
    # The init method here instantiates 3 objects upon creation of an Updater object. First it creates the
    # chromedriver object from selenium to control the browser, then it creates a Map object from xpath_map.py which
    # stores all the browser xpaths, finally it creates an ActionChains object which is used for more precise
    # control in some methods. In addition to this it also creates a counter attribute which is an integer of 1. This
    # is used for tracking the number of vendors processed for user feedback.
    def __init__(self):
        self.driver = None
        self.gui_elements = self.setup_gui()
        self.supplier_id = ''
        self.vc_id = ''
        self.map = Map()
        self.actions = None
        self.counter = 1

    # This method creates the chromedriver object and specifies the exe path for chromedriver.exe. It also specifies
    # the path for your browser User_data, a copy of which is needed to run the chrome with your normal profile.

    @staticmethod
    def chromedriver_setup():
        print('Starting Webdriver')
        my_options = webdriver.ChromeOptions()
        my_options.add_argument('user-data-dir=C:/Users/Tyler/ProgramingProjects/vendor_id_updater/User Data')
        chrome_service = ChromeService('C:/Users/Tyler/ProgramingProjects/vendor_id_updater/chromedriver.exe')
        chrome_service.creationflags = CREATE_NO_WINDOW
        driver_Setup = webdriver.Chrome(options=my_options, executable_path='C:/Users/Tyler/ProgramingProjects/'
                                                                            'vendor_id_updater/chromedriver.exe',
                                        service=chrome_service)
        return driver_Setup

    # This method has the driver object open the Vendor ID update list in Vcommerce then wait for the list to load.
    # If 20 seconds pass without loading it will create a timeout exception.
    def open_vcommerce(self):
        print('Opening Vcommerce')
        self.driver.get('https://solutions.sciquest.com/apps/Router/ApprNotifications?SelectedTab'
                        '=ApprNotificationSupplierRegistration&AuthUser=6637615&OrgName=Vulcan&tmstmp=1661890913947')
        sleep(3)
        login_button_element = self.driver.find_element(By.XPATH, self.map.vc_xpath['login_button'])
        login_button_element.click()
        WebDriverWait(self.driver, 20).until(ec.presence_of_element_located((By.XPATH, self.map.vc_xpath['search_'
                                                                                                         'results'])))

    # This method has the driver object open riskonnect in another tab, then switch to that tab and log in.
    def open_riskonnect(self):
        print('Opening Riskonnect\n')
        self.driver.execute_script("window.open('https://riskonnectvmc.lightning.force.com/lightning/page/home');")
        self.driver.switch_to.window(self.driver.window_handles[1])
        WebDriverWait(self.driver, 20).until(ec.presence_of_element_located((By.XPATH, self.map.rk_xpath['login_'
                                                                                                         'button'])))
        login_button_element = self.driver.find_element(By.XPATH, self.map.rk_xpath['login_button'])
        login_button_element.click()

    # This method has the driver object switch to the first browser tab, which has Vcommerce loaded.
    def switch_tab_to_1(self):
        self.driver.switch_to.window(self.driver.window_handles[0])

    # This method has the driver object switch to the second browser tab, which has Riskonnect loaded.
    def switch_tab_to_2(self):
        self.driver.switch_to.window(self.driver.window_handles[1])

    # This method has the driver object switch to the third browser tab, which has the vendor summary loaded.
    def switch_tab_to_3(self):
        self.driver.switch_to.window(self.driver.window_handles[2])

    # Closes the currently open tab.
    def close_current_tab(self):
        self.driver.close()
        
    # This method has the driver select and open the first vendor link on the Vcommerce vendor list in a new tab.
    def vc_open_summary(self):
        print(f'Accessing vendor {self.counter}')
        item_element = self.driver.find_element(By.XPATH, self.map.vc_xpath['first_item'])
        self.actions.key_down(Keys.CONTROL).perform()
        self.actions.click(item_element).perform()
        self.actions.key_up(Keys.CONTROL).perform()

    # This method has the driver switch to the Vcommerce summary page and wait for it to load.
    def vc_access_summary(self):
        print('Accessing summary')
        vc_summary_element = self.driver.find_element(By.XPATH, self.map.vc_xpath['summary_button'])
        vc_summary_element.click()
        WebDriverWait(self.driver, 20).until(ec.presence_of_element_located((By.XPATH, self.map.vc_xpath['summary_page'
                                                                                                         '_body'])))
        vc_summary_body_element = self.driver.find_element(By.XPATH, self.map.vc_xpath['summary_page_body'])
        vc_summary_body_element.send_keys(Keys.CONTROL + 'a')
        vc_summary_body_element.send_keys(Keys.CONTROL + 'c')

    # This method takes the copied data from the vc_summary and locates the supplier ID number and Vcommerce ID
    # number, which are saved as class attributes.
    def parse_summary(self):
        print('Reading vendor summary')
        summary_text = clipboard.paste()
        summary_list = summary_text.split()
        summary_list_2 = summary_text.split()
        for item in summary_list:
            if item == 'Number':
                temp_index = summary_list.index(item)
                if summary_list[temp_index - 1] == 'Supplier':
                    self.supplier_id = summary_list[temp_index + 1]
                    break
                else:
                    summary_list.remove(item)
                    continue
        for item in summary_list_2:
            if item == 'ID':
                temp_index = summary_list_2.index(item)
                self.vc_id = summary_list_2[temp_index + 1]
                break
    
    # This method searches Riskonnect for the Vcommerce number and selects opens the matching item in a new tab.
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
    
    # This method opens the supplier ID box in Riskonnect, clears the old value, and replaces it with the value saved
    # as the supplier ID attribute.
    def rk_update_supplier_id(self):
        print('Updating supplier ID')
        WebDriverWait(self.driver, 20).until(ec.presence_of_element_located((By.XPATH, self.map.rk_xpath['edit_'
                                                                                                         'button'])))
        edit_button_element = self.driver.find_element(By. XPATH, self.map.rk_xpath['edit_button'])
        edit_button_element.click()
        sleep(0.5)
        supplier_id_box = self.driver.find_element(By.XPATH, self.map.rk_xpath['supplier_id_box'])
        supplier_id_box.clear()
        sleep(0.5)
        supplier_id_box.send_keys(self.supplier_id)
        sleep(0.5)
        save_button_element = self.driver.find_element(By.XPATH, self.map.rk_xpath['save_button'])
        save_button_element.click()
        sleep(1)
        duplicate_check = self.driver.find_elements(By.XPATH, self.map.rk_xpath['duplicate_alert'])
        if duplicate_check:
            f = open('C:/Users/Tyler/ProgramingProjects/vendor_id_updater/duplicates_list.txt', 'a')
            f.write(f'vcommerce ID:{self.vc_id} did not accept supplier ID:{self.supplier_id} due to a duplicate\n')
            f.close()
        else:
            pass

    # This method has the driver return Riskonnect to its home screen.
    def rk_return_home(self):
        print('Resetting Riskonnect')
        sleep(0.5)
        WebDriverWait(self.driver, 20).until(ec.presence_of_element_located((By.XPATH, self.map.rk_xpath['home_'
                                                                                                         'button'])))
        home_button_element = self.driver.find_element(By.XPATH, self.map.rk_xpath['home_button'])
        home_button_element.click()

    # This method has the driver click the 'remove item' button on the vendor list once the update is completed.
    def vc_remove_item(self):
        print('Removing vendor')
        remove_button_element = self.driver.find_element(By.XPATH, self.map.vc_xpath['remove_button'])
        remove_button_element.click()

    # This method has the driver wait until the item has been removed from the Vcommerce list before moving forward.
    def wait_for_remove(self):
        sleep(2)
        WebDriverWait(self.driver, 20).until(ec.presence_of_element_located((By.XPATH, self.map.vc_xpath['remove_'
                                                                                                         'button'])))
        print(f'Vendor {self.counter} updated\n')
        self.counter += 1

    # This method runs both the open_vcommerce() method and the open_riskonnect() method. This is for use in the
    # __main__ script.
    def prep_update(self):
        self.open_vcommerce()
        self.open_riskonnect()

    # This method runs all methods in order to process the vendor update. This is for use in the __main__ script.
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

    def setup_gui(self):
        root = Tk()
        root.geometry('500x150')
        root.title('Vendor ID Updater')
        label_1 = Label(root, text='How many vendors do you want to process?')
        label_1.place(x=10, y=10)
        entry_box = Entry(root)
        entry_box.place(x=300, y=10)
        submit_button = Button(root, text='SUBMIT', command=self.submit_button_click)
        submit_button.place(x=150, y=50)
        exit_button = Button(root, text='EXIT', command=self.exit_button_click)
        exit_button.place(x=300, y=50)
        return root, label_1, entry_box, submit_button, exit_button

    def use_gui(self):
        root = self.gui_elements[1]
        root.mainloop()

    def submit_button_click(self):
        root, label_1, entry_box, submit_button, exit_button = self.gui_elements[0:]
        user_input = entry_box.get()
        try:
            user_input = int(user_input)
        except ValueError:
            label_1.configure(text='ERROR: Invalid Input, please input a number')
            entry_box.delete(0, 'end')
        else:
            self.driver = self.chromedriver_setup()
            self.actions = ActionChains(self.driver)
            label_1.configure(text='Preparing vendor update...')
            clipboard.copy(user_input)
            entry_box.delete(0, 'end')
            root.update()
            self.prep_update()
            label_1.configure(text=f'Updating {user_input} vendor(s)...')
            root.update()
            for i in range(user_input):
                self.process_update()
            label_1.configure(text=f'{user_input} vendor(s) have been updated successfully!')
            root.update()

    def exit_button_click(self):
        root = self.gui_elements[0]
        try:
            self.driver.quit()
        except AttributeError:
            root.destroy()
        else:
            root.destroy()
