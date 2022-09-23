from tkinter import *
from tkinter.ttk import *


# The UpdaterGUI class contains all code for a simple GUI to control the updater object and provide user feedback
# as to its progress.
class UpdaterGUI:

    # The constructor creates all GUI objects and uses .place() to organize them into the GUI.
    def __init__(self, updater_obj):
        self.updater_obj = updater_obj
        self.root = Tk()
        self.root.geometry('500x150')
        self.root.title('Vendor ID Updater')
        self.label_1 = Label(self.root, text='How many vendors do you want to process?')
        self.label_1.place(x=10, y=10)
        self.entry_box = Entry(self.root)
        self.entry_box.place(x=300, y=10)
        self.submit_button = Button(self.root, text='SUBMIT', command=self.submit_button_click)
        self.submit_button.place(x=150, y=50)
        self.exit_button = Button(self.root, text='EXIT', command=self.exit_button_click)
        self.exit_button.place(x=300, y=50)
        self.progress_bar = None
        self.root.mainloop()

    # This method is run when the GUI submit button is clicked. It takes the user input, checks to see if it is an
    # integer, then process a number of vendors that matches the user input.
    def submit_button_click(self):
        user_input = self.entry_box.get()
        # check user input to verify it is an integer
        try:
            user_input = int(user_input)
        # updates the GUI label with an error if user input is not a number.
        except ValueError:
            self.label_1.configure(text='ERROR: Invalid Input, please input a number')
            self.entry_box.delete(0, 'end')
        # if user input is an integer, then the chromedriver_setup() and actionchains_setup() methods will be called.
        else:
            self.updater_obj.driver = self.updater_obj.chromedriver_setup()
            self.updater_obj.actions = self.updater_obj.actionchains_setup()
            # then the user input is copied to the clipboard and the GUI is updated with feedback
            self.label_1.configure(text='Preparing vendor update...')
            self.updater_obj.clipboard_copy(user_input)
            self.entry_box.delete(0, 'end')
            self.root.update()
            # the prep_update method is called to get the webdriver where it needs to be and a progress bar is created
            self.updater_obj.prep_update()
            self.progress_bar = Progressbar(self.root, orient=HORIZONTAL, length=250, mode='determinate')
            self.progress_bar.place(x=125, y=100)
            self.root.update()
            progress_percent = 100 / user_input
            progress_amount = 0
            # the process_update() method is called and runs for the number of times specified in the users input.
            # the GUI progress bar is updated along the way.
            for i in range(user_input):
                self.label_1.configure(text=f'Updating vendor {i + 1} of {user_input}')
                self.root.update()
                self.updater_obj.process_update()
                progress_amount += progress_percent
                self.progress_bar['value'] = progress_amount
                self.root.update()
            # when complete, the GUI is updated and the submit button command changes to instead run the
            # re_submit_button_click() method the next time it is clicked.
            self.label_1.configure(text=f'{user_input} vendor(s) have been updated successfully!')
            self.submit_button.configure(command=self.re_submit_button_click)
            self.root.update()

    # The re_submit_button_click() method is the same as the submit_button_click() method, only without recreating the
    # webdriver and actionchains objects (as they will already exist by this point), and without running the
    # Prep_update() method.
    def re_submit_button_click(self):
        self.progress_bar.destroy()
        user_input = self.entry_box.get()
        try:
            user_input = int(user_input)
        except ValueError:
            self.label_1.configure(text='ERROR: Invalid Input, please input a number')
            self.entry_box.delete(0, 'end')
        else:
            self.label_1.configure(text='Preparing vendor update...')
            self.updater_obj.clipboard_copy(user_input)
            self.entry_box.delete(0, 'end')
            self.root.update()
            self.progress_bar = Progressbar(self.root, orient=HORIZONTAL, length=250, mode='determinate')
            self.progress_bar.place(x=125, y=100)
            self.root.update()
            progress_percent = 100 / user_input
            progress_amount = 0
            for i in range(user_input):
                self.label_1.configure(text=f'Updating vendor {i + 1} of {user_input}')
                self.root.update()
                self.updater_obj.process_update()
                progress_amount += progress_percent
                self.progress_bar['value'] = progress_amount
                self.root.update()
            self.label_1.configure(text=f'{user_input} vendor(s) have been updated successfully!')
            self.root.update()

    # the exit_button_click() method closes the webdriver and GUI.
    def exit_button_click(self):
        self.updater_obj.driver.close()
        self.updater_obj.switch_tab_to_1()
        self.updater_obj.driver.close()
        self.root.destroy()
