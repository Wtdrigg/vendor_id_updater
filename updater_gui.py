from tkinter import *
from tkinter.ttk import *


class UpdaterGUI:

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

    def submit_button_click(self):
        user_input = self.entry_box.get()
        try:
            user_input = int(user_input)
        except ValueError:
            self.label_1.configure(text='ERROR: Invalid Input, please input a number')
            self.entry_box.delete(0, 'end')
        else:
            self.updater_obj.driver = self.updater_obj.chromedriver_setup()
            self.updater_obj.actions = self.updater_obj.actionchains_setup()
            self.label_1.configure(text='Preparing vendor update...')
            self.updater_obj.clipboard_copy(user_input)
            self.entry_box.delete(0, 'end')
            self.root.update()
            self.updater_obj.prep_update()
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
            self.submit_button.configure(command=self.re_submit_button_click)
            self.root.update()

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

    def exit_button_click(self):
        try:
            self.updater_obj.driver.quit()
        except AttributeError:
            self.root.destroy()
        else:
            self.root.destroy()
