from calendar import Calendar
from cgitb import text
import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter.messagebox import showinfo
from turtle import onclick
from tkcalendar import DateEntry
from datetime import date
from mail_setup import *
from mail_functions import *
tooltips = data["tooltips"]
result_page_1 = ""


class ShowRemove:
    def RemoveTextDomain(event):
        domain_value_text = domain_text.get()
        if domain_value_text == tooltips["domain"]:
            domain_value.delete(0, END)
            domain_value.config(foreground="#000000")

    def RemoveTextId(event):
        id_value_text =  id_text.get()
        if id_value_text == tooltips["id"]:
            id_value.delete(0, END)
            id_value.config(foreground="#000000")

    def RemoveTextPw(event):
        pw_value_text = pw_text.get()
        if pw_value_text == tooltips["pw"]:
            pw_value.delete(0, END)
            pw_value.config(foreground="#000000")

    def RemoveTextEm(event):
        pw_value_text = em_text.get()
        if pw_value_text == tooltips["em"]:
            em_value.delete(0, END)
            em_value.config(foreground="#000000")


    def ShowTextDomain(event):
        domain_value.config(state=NORMAL)
        domain_value_text = domain_text.get()

        if bool(domain_value_text) == False:
            domain_value.insert(0, tooltips["domain"])
            domain_value.config(foreground="#d3d3d3")

    def ShowTextId(event):
        id_value.config(state=NORMAL)
        id_value_text = id_text.get()
        if bool(id_value_text) == False:
            id_value.insert(0, tooltips["id"])
            id_value.config(foreground="#d3d3d3")

    def ShowTextPw(event):
        pw_value.config(state=NORMAL)
        pw_value_text = pw_text.get()

        if bool(pw_value_text) == False:
            pw_value.insert(0, tooltips["pw"])
            pw_value.config(foreground="#d3d3d3")

    def ShowTextEm(event):
        em_value.config(state=NORMAL)
        em_value_text = em_text.get()

        if bool(em_value_text) == False:
            em_value.insert(0, tooltips["em"])
            em_value.config(foreground="#d3d3d3")       







def StartFunction():
    # >>> Data for log in <<<
    domain_name = domain_text.get()
    id = id_text.get()
    pw = pw_text.get()

    # >>> List of folders can be selected for running test <<<
    #               type = bool
    all_folders_value = all_folders.get()
    inbox_value = inbox.get()
    folders_value = folders.get()
    shared_value = shared.get()
    spam_value = spam.get()
    fetching_value = fetching.get()
    
    folder_dict = {
        "Inbox": inbox_value,
        "Folders": folders_value,
        "Shared": shared_value,
        "Spam": spam_value,
        "Fetching": fetching_value
    }

    # >>> List of folders can be selected for running test <<<
    selected_folders = []
    for folder_name in folder_dict.keys():
        if folder_dict[folder_name] == True:
            selected_folders.append(folder_name)

    # >>> This function is under development <<<
    #               type = int 
    options_value = options_var.get()

    page_unread_dict = {}

    # >>>>> Groupware function at Mail menu <<<<<
    global domain, driver
    domain = "https://%s/ngw/app/#" % domain_name
    driver = StartDriver()
    AccessMail(domain, id, pw)
    
    global folder_page_unread
    folder_page_unread = {}
	
    for selected_folder in selected_folders:
        folder_page_unread.update({selected_folder: {}})
        page_unread_dict = AccessMailFolder(selected_folder)
        folder_page_unread[selected_folder] = page_unread_dict
    
def MailAnalysis(mail_type):
    Type_List = {
        "mail_dict" : "",
        "msg"  : "",
    }
    
    
    mail_dict = Logs.CollectExcelList(mail_type)
    print("mail_dict :",mail_dict)
    mail_list = list(mail_dict.keys())
    print(" mail_list :",mail_list)

    if  mail_type == "suspected_mails":
        msg = "[Suspected Mails] Found [%s] mails in your list." % str(len(mail_list))
    elif mail_type == "frequent_mails":
        msg = "[Frequent Mails] Found [%s] mails from frequent addresses in your list."  % str(len(mail_list))
    elif mail_type == "groupware_mails":
        msg = "[Groupware Mails] Found [%s] mails in your list." % str(len(mail_list))
    elif mail_type == "other_mails":
        msg = "[Others Mails] Found [%s] mails in your list." % str(len(mail_list))
    
    Type_List["mail_dict"] = mail_dict
    Type_List["msg"]       = msg
    return Type_List


    
def Messages(status):
    if status == "pass":
        msg = "Execution is executed successfully."
    else:
        #Logs.CreateFailLog()
        msg = "There is an error while executing. Please run again or report your issue."

    showinfo(title='Information', message=msg)    



def CheckboxAll():
    checkbox_list = [inbox_checkbox, folders_checkbox, shared_checkbox, fetching_checkbox, spam_checkbox]
    for checkbox in checkbox_list:
        if all_folders.get() == True:
            checkbox.select()
        else:
            checkbox.deselect()

def CheckFolders():
    folder_dict = {
        "inbox": {
            "value": inbox.get(),
            "button": inbox_checkbox
        },
        "folders": {
            "value": folders.get(),
            "button": folders_checkbox
        },
        "shared": {
            "value": shared.get(),
            "button": shared_checkbox
        },
        "fetching": {
            "value": fetching.get(),
            "button": fetching_checkbox
        },
        "spam": {
            "value": spam.get(),
            "button": spam_checkbox
        }
    }

    for folder_name in folder_dict.keys():
        if folder_dict[folder_name]["value"] == False and all_folders.get() == True:
            all_checkbox.deselect()

def MarkAsReadFunction():
    suspected_mails_dict = Logs.CollectExcelList("suspected_mails")
    frequent_mails_dict =  Logs.CollectExcelList("frequent_mails")
    groupware_mails_dict =  Logs.CollectExcelList("groupware_mails")
    other_mails_dict =  Logs.CollectExcelList("other_mails")

    selected_mails = {
        "suspected_mails_dict": suspected_mails_dict,
        "frequent_mails_dict": frequent_mails_dict,
        "groupware_mails_dict": groupware_mails_dict,
        "other_mails_dict": other_mails_dict
    }

    selected_mails_dict = {}
    for dict_name in selected_mails.keys():
        print("\n   dict_name: -> " + str(dict_name))
        for selected_checkbox in selected_mails[dict_name].keys():
            selected_mail = dict(selected_mails[dict_name][selected_checkbox])
            print("   Select Mail (dict): => " + str(selected_mail))
            
            mail_title = selected_mail["title"]
            print("   Mail Subject: => " + str(mail_title))
            
            page_number = selected_mail["page_number"]
            print("   Page Number: => " + str(page_number))
            
            page_index = selected_mail["page_index"]
            print("   Current Page Index: => " + str(page_index))
            
            folder_id = selected_mail["folder_id"]
            print("   Folder ID: => " + str(folder_id))
            
            send_date = selected_mail["send_date"]
            print("   Send Date: => " + str(send_date))
            
            folder_name = selected_mail["folder_name"]
            print("   Folder Name: => " + str(folder_name))
            
            #list_total = selected_mail["list_total"] check if list changes after mail is selected

            if folder_id not in selected_mails_dict.keys():    
                selected_mails_dict[folder_id] = {}
                selected_mails_dict[folder_id]["mail_list"] = []
                selected_mails_dict[folder_id]["folder_name"] = folder_name

            try:
                # List of mail checkbox is displayed 
                # (button 'View List' is clicked)
                checkbox = bool(selected_mail["mail_checkbox"])
                
                # if checkbox.get() == True => is selected
                # else: not selected
                if checkbox == True:
                    print("   Mail checkbox is selected")
                
                    detailed_mail = Title.FormatTitle(mail_title, send_date)
                    print("   Selected Mail with date") 

                    selected_mails_dict[folder_id]["mail_list"].append(detailed_mail)
                    print("   Append mail in selected mail list")
    
                else:
                    
                    print("   Mail checkbox is not selected")
            
            except KeyError:
                pass
    
    print("   \nselected_mails_dict: -> " + str(selected_mails_dict))
    for selected_folder_id in selected_mails_dict.keys():
        selected_mail_list = selected_mails_dict[selected_folder_id]["mail_list"]
        print("selected_mail_list: ->" + str(selected_mail_list))
        
        if bool(selected_mail_list) == True:
            folder_url = domain + "/mail/list/%s/" % selected_folder_id
            driver.get(folder_url)
            WaitUntilFolderLoaded(selected_folder_id)
            MarkAsRead_SelectedMails(*selected_mail_list)
    
    '''
        🚩To-Do: list_total = selected_mail["list_total"] check if list changes after mail is selected
                -> Check if list changes and selected mails are at others pages
                -> Check if selected mail is in other folder
    '''

def CheckDateSelect():
    if options_var.get() == 2:
        start_cal.configure(state=NORMAL)
        end_cal.configure(state=NORMAL)
    else:
        start_cal.configure(state=DISABLED)
        end_cal.configure(state=DISABLED)

def SelectDate():
    start_date = start_cal.get_date()
    end_date = end_cal.get_date()

    selected_date = (start_date, end_date)
    # Date format: yyyy-mm-dd

    return selected_date

def QuitExecution():
    
    root.destroy()
    try:
        driver.quit()
    except:
        pass

def ConfigButtons(row_number):
    start_button = ttk.Button(signin, text="Start", width=20, command=StartFunction)
    #start_button = ttk.Button(signin, text="Start", width=20, command=lambda : tkinterApp.show_frame(root, Page1))
    start_button.grid(column=2, row=row_number, columnspan=3, ipadx=6, ipady=2, sticky="W")
 
    quit_button = ttk.Button(signin, text="Quit", width=20, command=QuitExecution)
    quit_button.grid(column=4, row=row_number, columnspan=3, ipadx=6, ipady=2, sticky="W")

def ConfigEmptyLabel(frame, row_number):
    empty_label = ttk.Label(frame, text="")
    empty_label.grid(column=0, row=row_number, ipadx=4, ipady=2, sticky="W")

def ConfigSeparator(frame, row_number):
    separator = ttk.Separator(frame, orient='horizontal')
    separator.grid(column=0, row=row_number, columnspan=6, ipadx=4, ipady=2, sticky="W")

def ConFigLogFrame(frame, row_number):
    log_labelframe = LabelFrame(frame, width=460, height=200)
    log_labelframe.grid(row=row_number, columnspan=6, padx=15, ipadx=90, ipady=20, sticky="W")
    
    log_empty = Label(log_labelframe, text="Log is empty")
    log_empty.grid(row=row_number+1, columnspan=6, padx=15, ipadx=90, ipady=20, sticky="W")

def SaveCheckbox_isSelect(mail_checkbox, mail_text):
    if mail_checkbox.get() == True:
        Logs.WriteInExcel_Checkbox(mail_text)

def ShowMailList(current_row, mail_dict):
    mail_dict = dict(mail_dict["mail_dict"])
    

    for item in list(enumerate(mail_dict.keys())):
        # NOTE item = (0, 'FW: RE: mail app | [Date: 08/17 13:33]')
        
        item_var  = tk.BooleanVar()
        item_text = str(item[1])
        start_row = 0

        item_checkbox = tk.Checkbutton(signin, text=item_text, var=item_var, command=lambda: SaveCheckbox_isSelect(item_var, item_text))
        item_checkbox.grid(column=2, row=row_number, columnspan=4, ipadx=10, ipady=2, sticky="W")
        
        folder_name = mail_dict[item_text]["folder_name"]
        item_folder = ttk.Label(signin, text=folder_name)
        item_folder.grid(column=0, row=row_number, columnspan=2, ipadx=6, ipady=5, sticky="W")

        

        mail_dict[item_text]["checkbox"] = item_checkbox
        mail_dict[item_text]["var"]      = item_var
        mail_dict[item_text]["text"]     = item_text
    
    return mail_dict

def ConfigMailAnalysis(mail_type, current_row):
    '''
        mail_type: suspected_mails, frequent_mails, groupware_mails, other_mails
    '''

    mail_dict = Logs.CollectExcelList(mail_type)
    mail_list = list(mail_dict.keys())

    if mail_type == "suspected_mails":
        msg = "[Suspected Mails] Found [%s] mails in your list." % str(len(mail_list))
    elif mail_type == "frequent_mails":
        msg = "[Frequent Mails] Found [%s] mails from frequent addresses in your list."  % str(len(mail_list))
    elif mail_type == "groupware_mails":
        msg = "[Groupware Mails] Found [%s] mails in your list." % str(len(mail_list))
    elif mail_type == "other_mails":
        msg = "[Others Mails] Found [%s] mails in your list." % str(len(mail_list))
    
    label = ttk.Label(signin, text=msg)
    label.grid(column=0, row=current_row, columnspan=3, ipadx=6, ipady=2, sticky="W")

    view = ttk.Button(signin, text="View List", width=5, command=lambda: ShowMailList(current_row, mail_dict))
    view.grid(column=4, row=current_row, columnspan=3, ipadx=6, ipady=2, sticky="W")

    end_row = current_row + int(len(mail_list))

    return end_row

def ConfigHandlerButtons(current_row):
    move_spam_btn = ttk.Button(signin, text="Move to Spam", width=20, command="")
    move_spam_btn.grid(column=2, row=current_row, columnspan=2, ipadx=6, ipady=2, sticky="W")

    mark_read_btn = ttk.Button(signin, text="Mark as read", width=20, command=MarkAsReadFunction)
    mark_read_btn.grid(column=4, row=current_row, columnspan=2, ipadx=6, ipady=2, sticky="W")

def HandlerBar():
    # Start row of fie
    handler_start_row = signin_start_row+10
    
    # Handler - Suspected mails
    suspected_row = ConfigMailAnalysis("suspected_mails", handler_start_row)

    # Handler - Frequent mails
    frequent_row = ConfigMailAnalysis("frequent_mails", suspected_row+1)

    # Handler - Groupware mails
    groupware_row = ConfigMailAnalysis("groupware_mails", frequent_row+1)

    # Handler - Other mails
    other_row = ConfigMailAnalysis("other_mails", groupware_row+1)

    # Handler - Empty Row
    ConfigEmptyLabel(signin, other_row+1)

    # Handler - Buttons 'Mark as Read' and 'Move to Spam'
    ConfigHandlerButtons(other_row+2)



################## Start #######################
def Distance(start,x):
    return start + 40*x

class tkinterApp(tk.Tk):
    def __init__(self, *args, **kwargs):
        
        tk.Tk.__init__(self, *args, **kwargs)
        
        self.container = tk.Frame(self) 
        self.container.pack(side = "top", fill = "both", expand = True)
        self.container.grid_rowconfigure(0, weight = 1)
        self.container.grid_columnconfigure(0, weight = 1)
        
        self.frames = {} 
        
        for F in (StartPage,Page1,Page2):
            frame = F(self.container, self)
            self.frames[F] = frame
            frame.grid(row = 0, column = 0, sticky ="nsew")
            
        self.show_frame(StartPage)

        

    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()



class StartPage(tk.Frame):
    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()
        
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        
        global domain_value, domain_text
        domain_text  = tk.StringVar()
        height       = 35
        Left_Row_1   = 110
        Start_Y      = 100
       

        canvas = Canvas(self,width=700,height=600)
        canvas.grid(row=0, column=0, sticky="nsew")

        # Title App #
        Title = tk.Label(canvas,text = "MAILBOX MANAGEMENT",font=('Arial',20,'bold'),fg="black")
        canvas.create_window(190, 0, anchor="nw", window=Title)
        
        # Domain #
        domain_1    = tk.Label(canvas,text = "Domain*",font=('Arial',8,'bold'),fg="#6A6A6A")
        canvas.create_window(Left_Row_1,Distance(Start_Y,0),height=height , anchor="nw", window = domain_1 )

        domain_value = tk.Entry(canvas, textvariable = domain_text,width=59,bg="#F9FAFE",relief="flat",highlightthickness=2)
        domain_value.insert(0, tooltips["domain"])
        domain_value.config(foreground="#d3d3d3")
        domain_value.bind("<FocusIn>", ShowRemove.RemoveTextDomain)
        domain_value.bind("<FocusOut>", ShowRemove.ShowTextDomain)
        canvas.create_window(220, Distance(Start_Y,0),height=height, anchor="nw", window=domain_value)

        # ID #
        global id_value, id_text
        id_text = tk.StringVar()
        
        user_id   = tk.Label(canvas,text = "ID*",font=('Arial',8,'bold'),fg="#6A6A6A")
        canvas.create_window(Left_Row_1, Distance(Start_Y,1), anchor="nw", window=user_id,height=height)
        
        id_value = tk.Entry(canvas, textvariable = id_text,width=59,bg="#F9FAFE",relief="flat",highlightthickness=2)
        id_value.insert(0, tooltips["id"])
        id_value.config(foreground="#d3d3d3")
        id_value.bind("<FocusIn>", ShowRemove.RemoveTextId)
        id_value.bind("<FocusOut>", ShowRemove.ShowTextId)
        canvas.create_window(220, Distance(Start_Y,1), anchor="nw", window=id_value,height=height)

        # Password #
        global pw_value, pw_text
        pw_text = tk.StringVar()
        
        uer_password = tk.Label(canvas,text = "Password*",font=('Arial',8,'bold'),fg="#6A6A6A")
        canvas.create_window(Left_Row_1, Distance(Start_Y,2), anchor="nw", window=uer_password,height=height)

        pw_value = tk.Entry(canvas, textvariable = pw_text,width=59,bg="#F9FAFE",relief="flat",highlightthickness=2)
        pw_value.insert(0, tooltips["pw"])
        pw_value.config(foreground="#d3d3d3")
        pw_value.bind("<FocusIn>", ShowRemove.RemoveTextPw)
        pw_value.bind("<FocusOut>", ShowRemove.ShowTextPw)
        canvas.create_window(220, Distance(Start_Y,2), anchor="nw", window=pw_value,height=height)
        
        # Mail
        global em_value, em_text
        em_text = tk.StringVar()
        
        uer_password = tk.Label(canvas,text = "Email*",font=('Arial',8,'bold'),fg="#6A6A6A")
        canvas.create_window(Left_Row_1, Distance(Start_Y,3), anchor="nw", window=uer_password,height=height)

        placeholder_pw = tooltips["em"]
        em_value = tk.Entry(canvas, textvariable = em_text,width=59,bg="#F9FAFE",relief="flat",highlightthickness=2)
        em_value.insert(0, placeholder_pw)
        em_value.config(foreground="#d3d3d3")
        em_value.bind("<FocusIn>", ShowRemove.RemoveTextEm)
        em_value.bind("<FocusOut>", ShowRemove.ShowTextEm)
        canvas.create_window(220, Distance(Start_Y,3), anchor="nw", window=em_value,height=height)

        # Folder #
        global all_checkbox, inbox_checkbox, folders_checkbox, shared_checkbox, fetching_checkbox, spam_checkbox
        global all_folders, inbox, folders, shared, fetching, spam
        
        all_folders = tk.BooleanVar()
        inbox       = tk.BooleanVar()
        folders     = tk.BooleanVar()
        shared      = tk.BooleanVar()
        spam        = tk.BooleanVar()
        fetching    = tk.BooleanVar()

        menu_label = tk.Label(canvas,text = "Folders",font=('Arial',8,'bold'),fg="#6A6A6A")
        canvas.create_window(Left_Row_1, Distance(Start_Y,5), anchor="nw", window=menu_label)

        # Folder row1 #
        all_checkbox = tk.Checkbutton(canvas, text="All", var=all_folders, command=CheckboxAll)
        canvas.create_window(220, Distance(Start_Y,5), anchor="nw", window=all_checkbox)

        
        inbox_checkbox = tk.Checkbutton(canvas, text="Inbox", var=inbox, command=CheckFolders)
        canvas.create_window(320, Distance(Start_Y,5), anchor="nw", window=inbox_checkbox)

        folders_checkbox = tk.Checkbutton(canvas, text="Folders", var=folders, command=CheckFolders)
        canvas.create_window(420, Distance(Start_Y,5), anchor="nw", window=folders_checkbox)

        
        # Folder row2 #
        shared_checkbox = tk.Checkbutton(canvas, text="Shared", var=shared, command=CheckFolders)
        canvas.create_window(220, Distance(Start_Y,6), anchor="nw", window=shared_checkbox)

        fetching_checkbox = tk.Checkbutton(canvas, text="Fetching", var=fetching, command=CheckFolders)
        canvas.create_window(320, Distance(Start_Y,6), anchor="nw", window=fetching_checkbox)

        spam_checkbox = tk.Checkbutton(canvas, text="Spam", var=spam, command=CheckFolders)
        canvas.create_window(420, Distance(Start_Y,6), anchor="nw", window=spam_checkbox)
        
        
        # Options #
        global options_var
        options_var = tk.IntVar()

        year = int(Objects.year)
        month = int(Objects.month)
        day = int(Objects.day)
        default_date = date(year, month, day)

        options_label = tk.Label(canvas, text="Options")
        canvas.create_window(Left_Row_1, Distance(Start_Y,8), anchor="nw", window=options_label)

        auto = tk.Radiobutton(canvas, text="Auto", variable=options_var, value=1, command=CheckDateSelect)
        canvas.create_window(220, Distance(Start_Y,8), anchor="nw", window=auto)
        
        
        custom = tk.Radiobutton(canvas, text="Custom", variable=options_var, value=2, command=CheckDateSelect)
        canvas.create_window(310, Distance(Start_Y,8), anchor="nw", window=custom)
        
        global start_cal, end_cal
        start_cal = DateEntry(canvas, selectmode='day', width=8)
        start_cal.set_date(default_date)
        start_cal.configure(state=DISABLED)
        canvas.create_window(410, Distance(Start_Y,8), anchor="nw", window=start_cal)

        end_cal = DateEntry(canvas, selectmode='day', width=8)
        end_cal.set_date(default_date)
        end_cal.configure(state=DISABLED)
        canvas.create_window(510, Distance(Start_Y,8), anchor="nw", window=end_cal)

        options_var.set(1)
        
        start_button = ttk.Button(canvas, text="Start", width=20, command=lambda:self.get_data_for_page1(controller))
        canvas.create_window(300, Distance(Start_Y,10), anchor="nw", window = start_button)

    def get_data_for_page1(self,controller):
        StartFunction()

        '''
        #suspected_mails#
        Mail_List = MailAnalysis("suspected_mails")
        controller.frames[Page1].Load_UI(Mail_List,controller)
        '''
        #frequent_mails#
        Mail_List = MailAnalysis("frequent_mails")
        controller.frames[Page1].Load_UI(Mail_List,controller)
        
        '''
        #groupware_mails#
        Mail_List = MailAnalysis("groupware_mails")
        controller.frames[Page1].Load_UI(Mail_List,controller)

        
        #other_mails#
        Mail_List = MailAnalysis("other_mails")
        controller.frames[Page1].Load_UI(Mail_List,controller)
        '''




        controller.show_frame(Page1)
       

    
class Page1(tk.Frame):
    def __init__(self, parent, controller):
        
        tk.Frame.__init__(self, parent)
        self.height       = 35
        self.Left_Row_1   = 110
        self.Start_Y      = 100
       
    def Load_UI(self,Mail_List,controller)  :
        
        
        canvas = Canvas(self,width=700,height=600,relief=RIDGE,background="red")
        canvas.grid(row=0, column=0, sticky="nsew")
        

        # Title App #
        Title = tk.Label(canvas,text = "MAILBOX MANAGEMENT",font=('Arial',20,'bold'),fg="black")
        canvas.create_window(190, 0, anchor="nw", window=Title)
        
        # Back #
        back = tk.Label(self,text = "Back to main page",font=('Arial',8,'bold'),fg="#6A6A6A")
        canvas.create_window(self.Left_Row_1,Distance(self.Start_Y,0),height=30 , anchor="nw", window=back )

        # Button Move #
        move_spam_btn = ttk.Button(canvas, text="Move to Spam", width=15, command="")
        canvas.create_window(390,Distance(self.Start_Y,0),height=30 , anchor="nw", window=move_spam_btn)

        # Button Mark #
        mark_read_btn = ttk.Button(canvas, text="Mark as read", width=15, command=MarkAsReadFunction)
        canvas.create_window(490,Distance(self.Start_Y,0),height=30 , anchor="nw", window=mark_read_btn)
       

        Suspected     = ToggledFrame(canvas,controller, text=Mail_List["msg"], borderwidth=1)
        canvas.create_window(self.Left_Row_1, Distance(self.Start_Y,1), anchor="nw", window=Suspected)

        mail_dict = dict(Mail_List["mail_dict"])

        start_row = 0
        for item in list(enumerate(mail_dict.keys())):
            # NOTE item = (0, 'FW: RE: mail app | [Date: 08/17 13:33]')
            
            item_var  = tk.BooleanVar()
            item_text = str(item[1])
            

            item_checkbox = tk.Checkbutton(Suspected.sub_frame, text=item_text, var=item_var, command=lambda: SaveCheckbox_isSelect(item_var, item_text))
            item_checkbox.grid(column=0, row=start_row, columnspan=4, ipadx=10, ipady=2, sticky="W")
            
            folder_name = mail_dict[item_text]["folder_name"]
            item_folder = ttk.Label(Suspected.sub_frame, text=item_text,relief=RIDGE)
            item_folder.grid(column=1, row=start_row, columnspan=2, ipadx=6, ipady=5, sticky="W")

            mail_dict[item_text]["checkbox"] = item_checkbox
            mail_dict[item_text]["var"]      = item_var
            mail_dict[item_text]["text"]     = item_text

            start_row = start_row +1

        
       



class Page2(tk.Frame):
    
    def __init__(self, parent, controller):
        
        tk.Frame.__init__(self, parent)
        height       = 35
        Left_Row_1   = 110
        Start_Y      = 100
        
        canvas = Canvas(self,width=700,height=600)
        canvas.grid(row=0, column=0, sticky="nsew")

        # Title App #
        Title = tk.Label(canvas,text = "MAILBOX MANAGEMENT",font=('Arial',20,'bold'),fg="black")
        canvas.create_window(190, 0, anchor="nw", window=Title)
        
        # Domain #
        user_name = tk.Label(self,text = "Back to main page",font=('Arial',8,'bold'),fg="#6A6A6A")
        canvas.create_window(Left_Row_1,Distance(Start_Y,0),height=height , anchor="nw", window=user_name )

       


class ToggledFrame(tk.Frame):

    def __init__(self, parent,controller, text="", *args, **options):
        tk.Frame.__init__(self, parent, *args, **options)
        
        self.show = tk.IntVar()
        self.show.set(0)

        self.title_frame = ttk.Frame(self)
        self.title_frame.pack(fill="x", expand=1)

        # 
        ttk.Label(self.title_frame, text=text , width=65 ).pack(side="left")
        self.toggle_button = ttk.Checkbutton(self.title_frame, width=10, text='Hide List', command=self.toggle,
                                            variable=self.show, style='Toolbutton')
        self.toggle_button.pack(side="left")
        self.sub_frame = tk.Frame(self)
        
        
        


    def toggle(self):
        if bool(self.show.get()):
            self.sub_frame.pack(fill="x", expand=1)
            self.toggle_button.configure(text='Hide List')
        else:
            self.sub_frame.forget()
            
            self.toggle_button.configure(text='View List')





def MainUI():
    app = tkinterApp()
    app.resizable(True, True)
    app.mainloop()
    
    


MainUI()