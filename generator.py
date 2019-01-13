from tkinter import *
from tkinter import ttk
from tkinter import messagebox



def check_valid_folder_name(name):
    for i in name:
        if ord(i) > 64 and ord(i) < 91:
            continue
        if ord(i) > 96 and ord(i) < 123:
            continue
        if ord(i) > 47 and ord(i) < 58:
            continue
        if ord(i) == 95 or ord(i) == 32:
            continue
        else:
            return False
    return True

class Window_Generator:
    def __init__(self, master, title, entry1, entry2, button, button_fun):
        self.button_fun = button_fun
        self.master = master
        self.master.attributes("-topmost", True)
        self.entry1_var = StringVar()
        self.entry2_var = StringVar()
        self.master.rowconfigure(1, pad = 20, weight = 1)
        self.master.rowconfigure(4, pad = 20, weight = 1)
        self.main_label = ttk.Label(self.master, text = title)
        self.main_label.grid(column = 1, row =1, columnspan = 2)
        #labels
        self.entry1_label = ttk.Label(self.master, text = entry1, justify = 'center').grid(column = 1, row = 2, sticky = 'e')
        self.entry2_label = ttk.Label(self.master, text = entry2, justify = 'center').grid(column = 1, row = 3)
        #text fields
        self.entry1_field = ttk.Entry(self.master, textvariable = self.entry1_var).grid(column = 2, row = 2)
        self.entry2_field = ttk.Entry(self.master, textvariable = self.entry2_var).grid(column = 2, row =3)
        #button
        self.button = ttk.Button(self.master, text = button, command = self.button_press)\
                      .grid(row = 4, column=2, columnspan = 2)
        for child in self.master.winfo_children():
            child.grid_configure(padx = 10, pady = 5)
        master.grab_set()
        master.wait_window()

    def button_press(self):
        if not (self.entry1_var.get() and self.entry2_var.get()):
            return
        if not (check_valid_folder_name(self.entry1_var.get()) and check_valid_folder_name(self.entry1_var.get())):
            messagebox.showinfo(message = "Names can only contain alphanumeric characters.")
            return
        else:
            self.button_fun(self.entry1_var.get(), self.entry2_var.get())
            self.master.destroy()
            return


