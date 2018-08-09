##Test Script

import subprocess       #To open/close external programs
import time             #To use sleep function, program waits
import shutil           #To rename files/folders
import os               #To search for file names
import psutil           #To search for inventor pid to kill
import win32gui
import win32con
from tkinter import *
from tkinter import ttk
from tkinter import messagebox


##GUI


def change(folder):
    #check for folder_name.txt in source folder
    source_path = 'C:\\InventorWork'
    destination_path = 'C:\\' + folder
    if os.path.isfile(source_path + '\\' + 'folder_name.txt'):
         pass
    else:
        print('no txt file')
        raise LookupError('Cannot find folder_name.txt')
    #check for folder_name in destination folder
    if os.path.isfile(destination_path + '\\' + 'folder_name.txt'):
        pass
    else:
        no_proj(None, None, folder)     #create folder_name.txt with folder as name
    iw_change = 'C:\\' + folder
    if os.path.exists(iw_change):
        if kill():
            print('closed Inventor')
            rename()
            print('ran rename()')
            time.sleep(1)
            print(iw_change)
            time.sleep(1)
            os.rename(iw_change, 'C:\\InventorWork')
            print('renamed from' + iw_change)
            open_inv()
            print('opened inventor')
            print('done')
            return True
        else:
            print('inventor not closed')
            raise OSError('Inventor not closed')

def make(project,machine):
    if kill():
        if rename():
            newIW(project, machine)
            return True
        else:
            return False
    else:
        return False
    

##Close current inventor session

def kill():
    handle = WindowEnumerate()
    print(handle)
    if not handle:
        print('inventor already closed')
        return True
    win32gui.SetForegroundWindow(handle)
    win32gui.SendMessage(handle,win32con.WM_CLOSE,0,0)  #waits for message to be processed
    if not WindowEnumerate():
        print('kill returning True')
        return True
    else:
        print('kill returning False')
        return False

#windows enumeration handler functions, enumerates windows to get handlers
    #then gets window names
    #adapted from: https://www.blog.pythonlibrary.org/2014/10/20/pywin32-how-to-bring-a-window-to-front/

def windowEnumerationHandler(hwnd, top_windows):
    top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))

def WindowEnumerate():
    top_windows = []
    win32gui.EnumWindows(windowEnumerationHandler, top_windows)
    for i in top_windows:
        if 'Autodesk Inventor Professional' in i[1]:
            print(i)
            return i[0]
    return False
    

def open_inv():
    p = subprocess.Popen('C:\\Program Files\\Autodesk\\Inventor 2016\\Bin\\Inventor.exe');
    print('opened inventor')
    print(p.poll() == None);
    time.sleep(5);
    print(p.poll() == None);



#Rename current inventorwork folder to match folder_name.txt
def rename():
    txtpath = 'C:\\InventorWork\\folder_name.txt'
    folder_txt = open(txtpath)
    source_name = folder_txt.read()
    folder_txt.close()
    time.sleep(1)
    result = False
    #loading window for rename, if folder is locked
    while not result:
        try:
            os.rename('C:\\InventorWork', 'C:\\' + source_name)
            print('Renamed InventorWork to ' + source_name)
            result = True
        except PermissionError:
            continue
    return True


def newIW(project, machine):
    p_name = 'InventorWork_' + project.replace(' ', '') + '_' + machine.replace(' ', '') # strips input spaces formates to _project_machine
    batf = 'M:\\Project Engineering\\ENGDATA\\Reference Data\\Dev and Training Resc\\cad\\30 Inventor\\New InventorWork.bat'
    os.startfile(batf)         #inventor must be opened to generate project file in IW
    time.sleep(2)       #sleep buffer to make sure project file is generated, not sure if required?
    no_proj(project,machine, None)
    open_inv()
    return 'done newIW'

def no_proj(project,machine, folder):
    #if folder name input, generate folder_name.txt file in folder input
    if folder != None:
        destination_path = 'C:\\' + folder
        open(destination_path + '\\folder_name.txt', 'a+')
        folder_txt = open(destination_path + '\\folder_name.txt', 'w')
        folder_txt.write(folder)
        folder_txt.close
    #else generate txt from project and machine input
    else:
        open('C:\\InventorWork\\folder_name.txt', 'a+')
        folder_txt = open('C:\\InventorWork\\folder_name.txt', 'w')
        folder_txt.write('InventorWork_' + project + '_' + machine)
        folder_txt.close
    
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

### GUI ################

class igui:

    def __init__(self, master):
        self.master = master            #master = root = tk()
        self.master.title("InventorWork Manager")   #master.title = root.title = tk().title

        #initializing tkk frame to hold widgets, sets size to expand and padding
        self.mainframe = ttk.Frame(master, padding = '3 3 12 12')
        self.mainframe.grid(column = 0, row = 0, sticky = (N,S,E,W))
        self.mainframe.columnconfigure(0, weight = 1)
        self.mainframe.rowconfigure(0, weight = 1)

        #labels and buttons
        self.newlabel = ttk.Label(self.mainframe, text = 'New InventorWork',
                                  justify = 'left')
        self.newlabel.grid(column = 1, row = 1, sticky = W)
        self.changelabel = ttk.Label(self.mainframe, text = 'Change InventorWork',
                                     justify = 'left')
        self.changelabel.grid(column = 1, row = 4, sticky = W)
        self.NewButton = ttk.Button(self.mainframe, text = 'New InventorWork', command = self.callwindow)
        self.NewButton.grid(column = 1, row = 2, sticky = W)
        self.ChangeButton = ttk.Button(self.mainframe, text = 'Change InventorWork', command = self.yousure)
        self.ChangeButton.grid(column = 2, row = 5, sticky = (E,W))
        self.HelpButton = ttk.Button(self.mainframe, text='?', command = self.help)
        self.HelpButton.grid(column = 3, row = 1, sticky = (W))
        #scrollbar for folder selection and seperator
        self.folder_var = StringVar()
        self.separator = ttk.Separator(self.mainframe, orient = HORIZONTAL).grid(row = 3, columnspan = 3, sticky = (E, W))
        self.iscroll = ttk.Combobox(self.mainframe, textvariable = self.folder_var,
                                    width = 40)

        self.update_list()
        self.iscroll.bind('<<ComboboxSelected>>', self.comboselect)
        self.iscroll.grid(column = 1, row = 5)
        for child in self.mainframe.winfo_children():
            child.grid_configure(padx=5, pady=5)

    #function called everyime scroll window selection changed, sets variable
            #equal to selection
    def comboselect(self,eventObject):
        self.change_selection = self.folder_var.get()
        print(self.change_selection)
        return self.change_selection

    #function called when change button is clicked

    def changeclick(self):
        try:
            change(self.change_selection)
            self.update_list()
            print('hwat')
        except LookupError as e:
            print(repr(e))
            self.rename_load.destroy()
            self.rename_load.update()
            self.noproject = messagebox.showinfo(message = "No folder_name.txt file in InventorWork folder. Please name the"
                                                             " current InventorWork folder.")
            self.cwindow = Toplevel(self.master)
            self.app = window_change(self.cwindow)
            self.cwindow.grab_set()
            self.master.wait_window(self.cwindow)
            print('load window SHOULD BE HERE')
            self.rename_load = Toplevel(self.master)
            self.rename_load.transient()
            self.app2 = loadwindow(self.rename_load)
            self.rename_load.update()
            if self.app.changepressed:
                change(self.change_selection)
            print('from changeclick func')
            self.rename_load.destroy()
            self.rename_load.update()
            self.update_list()
        except OSError as e:
            print(repr(e))
            self.update_list()



        

    def yousure(self):
        self.yousure = messagebox.askyesno('Restart', 'Are you sure you want to restart Inventor and change InventorWork folder?', parent = self.master)
        if self.yousure:
            self.rename_load = Toplevel(self.master)
            self.rename_load.transient()
            self.app = loadwindow(self.rename_load)
            self.rename_load.update()
            print ('supposed to have ran loading window')
            self.changeclick()
            self.rename_load.destroy()
        else:
            pass


    def help(self):
        self.helpmessage = messagebox.showinfo(message = "This app manages InventorWork folders. InventorWork names are kept"
                                                           " track of by generating a folder_name.txt file in the InventorWork"
                                                           " folder. The user is prompted if a folder_name.txt file"
                                                            " doesn't exist. Inventor is restarted when changing folders if Inventor is open.\n\n Written by Adrian Tai")
        
            

    #method creates new window and passes new window as master into class that
            #contains the labels and buttons for change IW window
    def callwindow(self):
        if os.path.isfile('C:\\InventorWork\\folder_name.txt'):
            self.mwindow = Toplevel(self.master)
            results = window_make_txt_exists(self.mwindow)
            #self.mwindow.grab_set()
            #self.master.wait_window(self.mwindow)
            print('updating list')
            self.update_list()
        else:
            self.mwindow = Toplevel(self.master)
            self.app = window_make_notxt(self.mwindow)
            self.mwindow.grab_set()
            self.master.wait_window(self.mwindow)
            self.update_list()
            



    def update_list(self):
        self.folder_list =[]
        for i in os.listdir('C:\\'):
            if 'InventorWork' in i:
                self.folder_list.append(i)
        self.iscroll['values'] = self.folder_list

class window_make_txt_exists:
    def __init__(self, master):
        self.master = master     #sets self.master= master = initializing variable input
                                # intializing varialbe = self.mwindow =
                                #Toplevel(self.mainframe) = new window with mainframe
                                #as parent
        
        self.project = StringVar()
        self.machine = StringVar()
        self.master.rowconfigure(1, pad = 20, weight = 1)
        self.master.rowconfigure(4, pad = 20, weight = 1)
        self.main_label = ttk.Label(self.master, text = 'New InventorWork')
        self.main_label.grid(column = 1, row =1, columnspan = 2)
        self.project_label = ttk.Label(self.master, text = 'Project', justify = 'center').grid(column = 1, row = 2, sticky = 'e')
        self.machine_label = ttk.Label(self.master, text = 'Machine', justify = 'center').grid(column = 1, row = 3)
        self.project_entry = ttk.Entry(self.master, textvariable = self.project).grid(column = 2, row = 2)
        self.machine_entry = ttk.Entry(self.master, textvariable = self.machine).grid(column = 2, row =3)
        self.makebutton = ttk.Button(self.master, text = 'Make', command = self.newclick).grid(row = 4, column=2, columnspan = 2)
        for child in self.master.winfo_children():
            child.grid_configure(padx = 10, pady = 5)
        master.grab_set()
        master.wait_window()


    def newclick(self):
        #if not (self.project.get() and self.machine.get()):
            #return False
        if not (check_valid_folder_name(self.project.get()) and check_valid_folder_name(self.machine.get())):
                messagebox.showinfo(message = "Names can only contain alphanumeric characters.")
        else:
            self.yesno = messagebox.askyesno('Restart', "Are you sure you want to restart Inventor and create new InventorWork?", parent = self.master)
            if self.yesno:
                self.rename_load = Toplevel(self.master)
                self.rename_load.transient()
                self.app = loadwindow(self.rename_load)
                self.rename_load.update()
                print ('supposed to have ran loading window')
                print('destroyed toplevel')
                check = make(self.project.get(), self.machine.get())
                self.master.destroy()
                if check:
                    return True
                else:
                    return False

class window_change:
    def __init__(self, master):
        self.master = master     #sets self.master= master = initializing variable input
                                # intializing varialbe = self.mwindow =
                                #Toplevel(self.mainframe) = new window with mainframe
                                #as parent
        #variable detetcing ok button click
        self.changepressed = False
        self.project = StringVar()
        self.machine = StringVar()
        self.master.rowconfigure(4, pad = 20, weight = 1)
        self.master.rowconfigure(1, pad = 20, weight = 1)
        self.main_label = ttk.Label(self.master, text = 'No folder_name.txt found for current InventorWork. Please name current InventorWork.',justify = 'center')
        self.main_label.grid(row = 1, column = 1, columnspan = 2)
        self.project_label = ttk.Label(self.master, text = 'Project').grid(column = 1, row = 2, sticky = 'e')
        self.machine_label = ttk.Label(self.master, text = 'Machine').grid(column = 1, row = 3, sticky = 'e')
        self.project_entry = ttk.Entry(self.master, textvariable = self.project).grid(column = 2, row = 2)
        self.machine_entry = ttk.Entry(self.master, textvariable = self.machine).grid(column = 2, row =3)
        self.changebutton = ttk.Button(self.master, text = 'Change', command = self.changeclick2).grid(row = 4, column=2, columnspan = 2)
        for child in self.master.winfo_children():
            child.grid_configure(padx = 5, pady = 5)


    def changeclick2(self):
        if not (self.project.get() and self.machine.get()):
            return False
        if not (check_valid_folder_name(self.project.get()) and check_valid_folder_name(self.machine.get())):
            messagebox.showinfo(message = "Names can only contain alphanumeric characters.")
        else:
            self.yesno = messagebox.askyesno('Restart', "Are you sure you want to restart Inventor and change InventorWork?", parent = self.master)
            if self.yesno:
                no_proj(self.project.get(), self.machine.get(), None)
                self.master.destroy()
                self.changepressed = True
                print('destroyed toplevel')

class window_make_notxt:
    def __init__(self, master):
            
        self.master = master     #sets self.master= master = initializing variable input
                                # intializing varialbe = self.mwindow =
                                #Toplevel(self.mainframe) = new window with mainframe
                                #as parent
        
        self.project_old = StringVar()
        self.machine_old = StringVar()
        self.master.rowconfigure(1, pad = 20, weight = 1)
        self.master.rowconfigure(5, pad = 30, weight = 1)
        self.main_label = ttk.Label(self.master, text = 'No folder_name.txt found, please name the current and new InventorWork folders', justify='center')
        self.main_label.grid(row = 1, column = 1, columnspan = 4)
        self.ciw_label = ttk.Label(self.master, text = 'Current InventorWork', justify = 'center')
        self.ciw_label.grid(row = 2, column = 2)
        self.niw_label = ttk.Label(self.master, text = 'New InventorWork', justify = 'center')
        self.niw_label.grid(row = 2, column = 5) 
        
        self.project_label_old = ttk.Label(self.master, text = 'Project', justify = 'center').grid(column = 1, row = 3)
        self.machine_label_old = ttk.Label(self.master, text = 'Machine', justify = 'center').grid(column = 1, row = 4)
        self.project_entry_old = ttk.Entry(self.master, textvariable = self.project_old).grid(column = 2, row = 3)
        self.machine_entry_old = ttk.Entry(self.master, textvariable = self.machine_old).grid(column = 2, row =4)

        self.separator = ttk.Separator(self.master, orient = VERTICAL).grid(column = 3, row = 2, rowspan = 3, sticky = (N, S))

        self.project_new = StringVar()
        self.machine_new = StringVar()
        self.project_label_new = ttk.Label(self.master, text = 'Project', justify = 'center').grid(column = 4, row = 3)
        self.machine_label_new = ttk.Label(self.master, text = 'Machine', justify = 'center').grid(column = 4, row = 4)
        self.project_entry_new = ttk.Entry(self.master, textvariable = self.project_new).grid(column = 5, row = 3)
        self.machine_entry_new = ttk.Entry(self.master, textvariable = self.machine_new).grid(column = 5, row =4)
        
        self.makebutton = ttk.Button(self.master, text = 'Make', command = self.newclick).grid(row = 5, column=5, columnspan = 4)
        for child in self.master.winfo_children():
            child.grid_configure(padx = 20, pady = 5)

    def newclick(self):
        if not (self.project_new.get() and self.machine_new.get() and self.project_old.get() and self.machine_old.get()):
            return False
        if not (check_valid_folder_name(self.project_new.get()) and check_valid_folder_name(self.machine_new.get()) and check_valid_folder_name(self.project_old.get()) and check_valid_folder_name(self.machine_old.get())):
            messagebox.showinfo(message = "Names can only contain alphanumeric characters.")
        else:
            self.yesno = messagebox.askyesno('Restart', "Are you sure you want to restart Inventor and create new InventorWork?", parent = self.master)
            if self.yesno:
                self.rename_load = Toplevel(self.master)
                self.rename_load.transient()
                self.app = loadwindow(self.rename_load)
                self.rename_load.update()
                print('destroyed toplevel')
                no_proj(self.project_old.get(), self.machine_old.get(), None)
                check = make(self.project_new.get(), self.machine_new.get())
                self.master.destroy()
                if check:
                    return True
                else:
                    return False

class loadwindow:
    def __init__(self, master):
        self.master = master
        self.master.attributes("-topmost", True)
        self.mainlabel = ttk.Label(self.master, text = 'Folder may be in use, trying to rename folder...', justify = 'center')
        self.mainlabel.grid_configure(padx = 40, pady = 40)

        


        

            
root = Tk()
app = igui(root)
root.mainloop()
  
