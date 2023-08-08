import time, sys, unittest, random, json, requests, openpyxl, testlink
from selenium.webdriver.remote.webelement import WebElement
from datetime import date, datetime
from selenium import webdriver
from appium import webdriver
from random import randint
from openpyxl import load_workbook
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.alert import Alert

from mail_setup import *
mail_dict = data["mail"]
login = data["login"]

def DefineListTotal(list_footer_xpath):
    list_footer = Functions.GetElementText(list_footer_xpath)
    total_number = int(list_footer.replace(",", "").split(" ")[1])

    return total_number

def WaitUntilFolderLoaded(folder_id):
    Waits.WaitUntilPageIsLoaded(mail_dict["page_header"] % folder_id)
    Waits.WaitUntilPageIsLoaded(mail_dict["list_footer"])
    time.sleep(1)

def LogIn(domain_name, userid, userpw):
    driver.get(domain_name + "/sign")
    Logs.Logging("Log in - Access login page")

    Waits.Wait10s_ElementLoaded(login["pw_frame"])

    input_id = Commands.FindElement(login["input_id"])
    input_id.send_keys(userid)
    if input_id.get_attribute("value") != userid:
        input_id.clear()
        input_id.send_keys(userid)
    Logs.Logging("Log in - Input valid ID")

    frame_element = Commands.FindElement(login["pw_frame"])
    driver.switch_to.frame(frame_element)
    input_pw = Commands.FindElement(login["input_pw"])
    input_pw.send_keys(userpw)
    if input_pw.get_attribute("value") != userpw:
        input_pw.clear()
        input_pw.send_keys(userpw)
    driver.switch_to.default_content()
    Logs.Logging("Log in - Input valid password")

    Commands.FindElement(login["submit_btn"]).send_keys(Keys.RETURN)
    Logs.Logging("Log in - Click Submit button")

    login_status = None

    try:
        Waits.WaitElementLoaded(5, login["login_alert"])
        Logs.Logging("Login info is incorrect")
        login_status = False
    except WebDriverException:
        Logs.Logging("Log in successfully")
        login_status = True

    if login_status == True:
        try:
            Waits.WaitElementLoaded(20, login["user_capibility"])
            time.sleep(1)
            Waits.WaitUntilPageIsLoaded(None)
            time.sleep(1)
        except WebDriverException:
            Logs.Logging("Fail to access Groupware after log in")
            login_status[0] = False
    
    return login_status

def CheckGroupwareMail(mail):
    '''<<< Check if mail is sent from groupware menu >>>'''
    
    gwmail_symbol = ["[", "]", "(", ")"]
    # Mail which is sent from groupware menu will includes of these characters
    # and usually starts with "["
    
    if mail[0] == "[":
        check_result = []
        for symbol in gwmail_symbol:
            if symbol in mail:
                check_result.append(True)
            else:
                check_result.append(False)

    if False in check_result:
        return False
    else:
        return True

def CheckReplyForwardMail(mail):
    ''' <<< Check if mail is REPLY/FORWARD >>> '''

    mail = str(mail).lower()
    
    if mail.startswith("fwd:") or mail.startswith("fw:") == True:
        return True
    elif mail.startswith("re:") == True:
        return True
    else:
        return False

   

def MarkAsRead_SelectedMails(*detailed_mail_list):
    print("MarkAsRead_SelectedMails - list(detailed_mail_list): -> " + str(detailed_mail_list))
    for mail in detailed_mail_list:
        print("MarkAsRead_SelectedMails - mail: -> "+ str(mail))
        mail_title = Title.SplitTitle(mail)[0]
        send_date = Title.SplitTitle(mail)[1]
        
        Commands.ClickElement(mail_dict["detailed_mail_checkbox"] % (mail_title, send_date))
        Waits.Wait10s_ElementLoaded(mail_dict["selected_mail_checkbox"] % (mail_title, send_date))
        
        '''try:
            mail in current page
        except:
            pass
        
        move page'''
    
    '''
        🚩To-Do: Make function for multiple pages
    '''

def MailAnalysis(**folder_dict):
    page_unread_dict = {}

    Waits.Wait10s_ElementLoaded(mail_dict["unread_filter"])
    
    unread_number = Functions.GetElementText(mail_dict["unread_counter"]).split("/")[0].strip()
    if int(unread_number) == 0:
        unread_loading = False
    else:
        Commands.ClickElement(mail_dict["unread_filter"])
        unread_loading = True
      
    Waits.WaitUntilPageIsLoaded(mail_dict["list_footer"])
    time.sleep(1)

    if unread_loading == True:
        # <--- Create analysis of unread mail list on each page --->
        
        current_host = Functions.GetElementText(mail_dict["current_email"]).split("@")[1]
        current_folder =  Functions.GetElementText(mail_dict["current_folder"])
        page_total = int(Functions.GetElementText(mail_dict["page_total"]))
        
        
        for current_page in range(page_total):
            current_page+=1
            #current_page_str = Functions.GetElementText(mail_dict["current_page"])
            current_page_str = str(current_page)
            
            # <--- Get number of each type of mails --->
            count_unread_mails = len(Commands.FindElements(mail_dict["unread_msg"]))
            count_internal_mails = len(Commands.FindElements(mail_dict["mail_@title"] % current_host))
            count_external_mails = count_unread_mails - count_internal_mails
            count_important_mails = len(driver.find_elements_by_css_selector("label.important-mail input[type=checkbox]:checked+.lbl"))
            count_alias_mails = len(Commands.FindElements(mail_dict["alias_mail"]))
            count_suspected_mails = len(Commands.FindElements(mail_dict["suspected_mail"]))
            count_past_mails = len(Commands.FindElements(mail_dict["previous_date"]))
            count_gmail_fwd =  len(Commands.FindElements(mail_dict["mail_startswith"] % "Fwd:"))
            count_gw_fwd =  len(Commands.FindElements(mail_dict["mail_startswith"] % "FW:"))
            count_others_fwd =  len(Commands.FindElements(mail_dict["mail_startswith"] % "FWD:"))
            count_gmail_re = len(Commands.FindElements(mail_dict["mail_startswith"] % "Re:"))
            count_gw_re = len(Commands.FindElements(mail_dict["mail_startswith"] % "RE:"))
            count_today_mails = count_unread_mails - count_past_mails
            count_fwd_mails = count_gmail_fwd + count_gw_fwd + count_others_fwd
            count_re_mails = count_gmail_re + count_gw_re

            # "important_mails": {"position1": mail_text, "position2": mail_text}
            important_mails = {}
            alias_mails = {}
            internal_mails = {}
            external_mails = {}
            suspected_mails = {}
            gw_re_fwd_mails = {}
            other_re_fwd_mails = {}
            past_mails = {}
            today_mails = {}

            gw_tasks_mails = {}
            
            print("\n")

            important_input = Commands.FindElements(mail_dict["important_input"])
            for i in range(1, len(important_input)):
                mail = Commands.FindElement(mail_dict["mail_text"] % str(i))
                print("(Important Mail) - mail.text: -> " + str(mail.text))
                
                mail_property_list = important_input[i-1].get_property("attributes")
                mail_position = str(i)

                if len(mail_property_list) == 2 and mail_property_list[1]["name"] == "checked":
                    important_mails[mail_position] = mail.text
                
                i+=1
            
            mail_list = Commands.FindElements(mail_dict["mail_length"])
            for i in range(0, len(mail_list)):
                i+=1
                mail_position = str(i)
                Logs.Logging(" >>> mail_position: %s" % mail_position)
                
                mail = Waits.Wait10s_ElementLoaded(mail_dict["mail_text"] % mail_position)
                print("(Check Mail) - mail.text: -> " + str(mail.text))
                
                send_date = Functions.GetElementText(mail_dict["send_date"] % mail_position)
                
                folder_id = {"column": 11, "cell_value": folder_dict["folder_id"]}
                list_total = {"column": 12, "cell_value": folder_dict["list_total"]}

                try:
                    alias_mail = Commands.FindElement(mail_dict["alias_text"] % mail_position)
                    alias_mails[mail_position] = alias_mail.text

                    alias_dict = {
                        "title": {"column": 1, "cell_value": alias_mail.text},
                        "folder": {"column": 2, "cell_value": current_folder},
                        "page": {"column": 3, "cell_value": current_page_str},
                        "position": {"column": 4, "cell_value": mail_position},
                        "send_date": {"column": 10, "cell_value": send_date},
                        "folder_id": folder_id,
                        "list_total": list_total
                        }

                    if alias_mail.text == important_mails[mail_position]:
                        alias_dict.update(
                            {"important": {"column": 5, "cell_value": "True"},
                            "mail_type": {"column": 6, "cell_value": "frequent_mails"}})
                    else:
                        alias_dict.update(
                            {"mail_type": {"column": 6, "cell_value": "other_mails"}})
                    
                    Logs.WriteInExcel(**alias_dict)
                except WebDriverException:
                    pass
                
                try:
                    internal_mail = Commands.FindElement(mail_dict["internal_text"] % (mail_position, current_host))
                    internal_mails[mail_position] = internal_mail.text
                    
                    internal_dict = {
                        "title": {"column": 1, "cell_value": internal_mail.text},
                        "folder": {"column": 2, "cell_value": current_folder},
                        "page": {"column": 3, "cell_value": current_page_str},
                        "position": {"column": 4, "cell_value": mail_position},
                        "send_date": {"column": 10, "cell_value": send_date},
                        "folder_id": folder_id,
                        "list_total": list_total
                        }

                    gw_mail = DetectGroupwareMail(internal_mail.text)
                    if bool(gw_mail) == True:
                        menu_name = gw_mail[0]
                        status = gw_mail[1]
                        mail_title = internal_mail.text
                        gw_tasks_mails[mail_position] = {"menu": menu_name, "status": status, "title": mail_title}
                        internal_dict.update({
                            "mail_type": {"column": 6, "cell_value": "groupware_mails"},
                            "menu": {"column": 7, "cell_value": menu_name},
                            "status": {"column": 8, "cell_value": status},
                        })
                    else:
                        internal_dict.update({"mail_type": {"column": 6, "cell_value": "frequent_mails"}})
                    
                    Logs.WriteInExcel(**internal_dict)
                except WebDriverException:
                    external_mails[mail_position] = mail.text
                    
                    external_dict = {
                        "title": {"column": 1, "cell_value": mail.text},
                        "folder": {"column": 2, "cell_value": current_folder},
                        "page": {"column": 3, "cell_value": current_page_str},
                        "position": {"column": 4, "cell_value": mail_position},
                        "send_date": {"column": 10, "cell_value": send_date},
                        "folder_id": folder_id,
                        "list_total": list_total
                        }

                    try:
                        mail.text == important_mails[mail_position]
                        external_dict.update(
                            {"important": {"column": 5, "cell_value": "True"},
                            "mail_type": {"column": 6, "cell_value": "frequent_mails"}})
                    except KeyError:
                        external_dict.update(
                            {"mail_type": {"column": 6, "cell_value": "other_mails"}})
                    
                    Logs.WriteInExcel(**external_dict)
                
                try:
                    past_mail = Commands.FindElement(mail_dict["past_mails"] % mail_position)
                    past_mails[mail_position] = past_mail.text
                except WebDriverException:
                    today_mails[mail_position] = mail.text
                
                try:
                    suspected_mail = Commands.FindElement(mail_dict["suspected_text"] % mail_position)
                    suspected_mails[mail_position] = suspected_mail.text
                    
                    suspected_dict = {
                            "title": {"column": 1, "cell_value": suspected_mail.text},
                            "folder": {"column": 2, "cell_value": current_folder},
                            "page": {"column": 3, "cell_value": current_page_str},
                            "position": {"column": 4, "cell_value": mail_position},
                            "mail_type": {"column": 6, "cell_value": "suspected_mails"},
                            "send_date": {"column": 10, "cell_value": send_date},
                            "folder_id": folder_id,
                            "list_total": list_total
                        }
                    
                    Logs.WriteInExcel(**suspected_dict)
                except WebDriverException:
                    pass
                
                re_fw_keys = ["FW:","Fwd:", "FWD:", "RE:", "Re:"]
                for key in re_fw_keys:
                    if str(mail.text).startswith(key) == True:  
                        try:
                            Commands.FindElement(mail_dict["internal_text"] % (mail_position, current_host))
                            gw_re_fwd_mails[mail_position] = mail.text
                        except WebDriverException:
                            other_re_fwd_mails[mail_position] = mail.text
            
            # <--- Append analysis data on current to page unread dict --->
            page_unread_dict[current_page] = {
                "unread_mails": {
                    "total": count_unread_mails,
                    "title": {}
                },
                "internal_mails": {
                    "total": count_internal_mails,
                    "title": internal_mails
                },
                "external_mails": {
                    "total": count_external_mails,
                    "title": external_mails
                },
                "important_mails": {
                    "total": count_important_mails,
                    "title": important_mails
                },
                "alias_mails": {
                    "total": count_alias_mails,
                    "title": alias_mails
                },
                "suspected_mails": {
                    "total": count_suspected_mails,
                    "title": suspected_mails
                },
                "past_mails": {
                    "total": count_past_mails,
                    "title": past_mails
                },
                "today_mails": {
                    "total": count_today_mails,
                    "title": today_mails
                },
                "fwd_mails": {
                    "total": count_fwd_mails,
                    "title": {}
                },
                "re_mails": {
                    "total": count_re_mails,
                    "title": {}
                },
                "gw_re_fwd_mails": {
                    "total": len(gw_re_fwd_mails.keys()),
                    "title": gw_re_fwd_mails
                },
                "other_re_fwd_mails": {
                    "total": len(other_re_fwd_mails.keys()),
                    "title": other_re_fwd_mails
                },
                "groupware_mails": {
                    "total": len(gw_tasks_mails.keys()),
                    "title": gw_tasks_mails
                }
            }


            # <--- Move page to continue scraping data on next page --->
            try:
                Commands.FindElement(mail_dict["nextpage_disabled"])
                Logs.Logging("Current page %s is last page" % current_page)
            except WebDriverException:
                Commands.ClickElement(mail_dict["nextpage_icon"])
                Logs.Logging("Move to next page from page %s" % current_page)
            finally:
                Waits.WaitUntilPageIsLoaded(None)
                time.sleep(1)

        #Logs.Logging(" >>> count_page_unread_dict: " + str(page_unread_dict))
        '''
            🚩To-Do:
            
            important mails: if you think it's not important -> select -> uncheck
            suspected mail: it's recognized as suspected mail -> would you like to mark it as spam mails?
            internal re: / fwd: there are some re:/fwd mails sent from your service and you might need to work on it
            other re: / fwd: 
                + suspected: mark with label
                + re: / fwd: 
            crm: you have [0] assigned tasks from crm
        '''

    print("\npage_unread_dict:\n" + str(page_unread_dict))

    return page_unread_dict

def CheckDraftsFolder():
    msg = ""

    Commands.ClickElement(mail_dict["drafts_folder"])
    Logs.Logging("Access Drafts folder")

    Waits.Wait10s_ElementLoaded(mail_dict["list_footer_drafts"])

    try:
        Commands.FindElement(mail_dict["list_nodata"])
        drafts_list = 0
    except WebDriverException:
        drafts_list = 1
    
    if drafts_list == 1:
        drafts = int(Functions.GetElementText(mail_dict["list_footer_drafts"]).replace(",", "").split(" ")[1])
        Logs.Logging(" >>> count_drafts: " + str(drafts))
        
        msg = "You have [%s] drafts to be resolved"

        ''' 🚩To-Do: Store number of drafts to compare with the previous check
                -> Show msg: You have [%s] new drafts to be resolved comparing to last check'''
    
    return msg

def StartsWith(title, key):
    if str(title).startswith(key) == True:
        return True
    else:  
        return False
        
def EndsWith(title, key):
    if str(title).endswith(key) == True:
        return True
    else:
        return False

def Contains(title, key):
    if "," in key:
        # If there are 2 conditions to be checked, it will be splitted by the comma
        #    E.g: mail_title = hanh1's Project(hanh123) Ticket Notice [Person in Charge-hanh2]
        #        ==> key = "Ticket Notice,Person in Charge" 

        # ===> Purpose: Check with 2 conditions
        key_list = str(key).split(",")
        if key_list[0] and key_list[1] in title:
            return True
        else:
            return False
    else:
        # ===> Purpose: Check with single condition
        if key in title:
            return True
        else:
            return False

def ValidateMethod(title, str_method, key):
    # str_method = gw_keyword_dict.keys()
    if str_method == "startswith":
        key_bool = StartsWith(title, key)
    elif str_method == "endswith":
        key_bool = EndsWith(title, key)
    else:
        key_bool = Contains(title, key)
    
    return key_bool

def DetectGroupwareMail(mail_title):
    gw_keyword_dict = {
        "startswith": {
            "approval": {
                "en": ["(Request)", "(Complete)", "(Reject)", "(Referrer)"],
                "ko": "(결재요청)",
                "vi": "(Yêu cầu bởi)"
            },
            "timecard": {
                "en": ["[Request] HR > Approval", "[Approved] HR > Approval", "[Cancelled] HR > Approval"]
            },
            "vacation": {
                "en": ["[Request]", "[Approved]", "[Cancelled]"]
            },
            "circular": {
                "en": "[Circular]",
                "ko": "[회람판]",
                "vi": "[Công văn]"
            },
            "clouddisk": {
                "en": "Weblink from"
            },
            "calendar": {
                "vi": "Đây là bảng thông báo lịch của anh/chị"
            }
        },

        "endswith": {
            "resource": {
                "en": ["(Approved)", "(Cancel)", "(Modify)"],
                "ko": ["(승인 완료)", "(예약취소)", "(예약 수정)"],
                "vi": ["(Đã được phê duyệt)", "(Hủy bỏ)", "(Sửa)"]
            },
            "clouddisk": {
                "ko": "님이 보내신 웹링크 입니다."
            },
            "to-do": {
                "ko": "의 ToDo 공지"
            },
            "calendar": {
                "en": " Schedule",
                "ko": "님의 일정알림입니다."
            }
        },

        "contains": {
            "resource_other": {
                "en": "Date (Use Period)",
                "ko": "사용일(사용기간)",
                "vi": "Ngày sử dụng"
            },
            "resource_permission": {
                "en": "Permission System",
                "ko": "승인제",
                "vi": "Đặt trước cần phê duyệt"
            },
            "resource_meeting": {
                "en": "You are invited to a conference.",
                "ko": "회의 참석을 안내 합니다.",
                "vi": "Bạn được mời đến hộp."
            },
            "project_work": {
                "en": "Project,Work Notice"
            },
            "project_ticket": {
                "en": "Ticket Notice,Person in Charge",
                "ko": "님의 프로젝트,티켓 알림입니다",
                "vi": "Là thông báo dự án,Người phụ trách"
            },
            "clouddisk": {
                "vi": "Là weblink,khách đã gửi."
            },
            "archive": {
                "en": "Archive,Report Notice",
                "vi": "Thông tin,lưu trữ"
            },
            "to-do": {
                "en": "To-Do,Notice",
                "vi": "Thông tin,báo cáo"
            },
            "crm_helpdesk": {
                "en": "You have been assigned as Rep for Ticket",
                "ko": "티켓,의 담당자로 지정 되었습니다.",
                "vi": "Đã chỉ định người phụ trách cho trao đổi"
            }
        }
    }

    mail_tuple = ()

    for str_method in gw_keyword_dict.keys():
        for menu_name in dict(gw_keyword_dict[str_method]).keys():
            lang_list =  dict(gw_keyword_dict[str_method][menu_name]).keys()
            for language in lang_list:
                search_keyword = gw_keyword_dict[str_method][menu_name][language]
                
                if type(search_keyword) == list:
                    # Multiple keywords
                    # ===> Menu with multiple types of mail (normally for status Request, Approve, Cancel)
                    for key in search_keyword:
                        if ValidateMethod(mail_title, str_method, key) == True:
                            mail_tuple = (menu_name, key)
                            Logs.Logging("Mail is sent from gw menu >>> Menu: %s | Status: %s" % mail_tuple)
                            key_found = menu_found = break_mloop = True
                            break
                        else:
                            key_found = menu_found = break_mloop = False
                
                else:
                    # Single keyword
                    if ValidateMethod(mail_title, str_method, search_keyword) == True:
                        mail_tuple = (menu_name, search_keyword)
                        print("str_method: %s | search key: %s" % (str(str_method), str(search_keyword)))
                        Logs.Logging("Mail is sent from gw menu >>> Menu: %s | Status: %s" % mail_tuple)
                        key_found = menu_found = break_mloop = True
                    else:
                        key_found = menu_found = break_mloop = False

                if key_found == True:
                    break
            
            if menu_found == True:
                break
        
        if break_mloop == True:
            break
    
    return mail_tuple

def AccessMail(domain_name, user_id, user_pw):
    LogIn(domain_name, user_id, user_pw)
    driver.get(domain_name + "/mail/list/all/")
    Waits.Wait10s_ElementLoaded(mail_dict["list_footer"])

def AccessFolderByName(folder_name):
    '''
        🚩To-Do: Detect language and translate language for folder_name [ko,vi]
                other language > refuse to run
    '''

    Logs.MsgLogging("\n[%s]" % folder_name)

    Waits.WaitUntilPageIsLoaded(None)

    page_name = {
            "Inbox": "mail_Maildir",
            "Fetching": "mail_External",
            "Spam": "mail_Spam"
        }

    if folder_name in list(page_name.keys()):
        page_id = page_name[folder_name]
        default_folder = True
    else:
        default_folder = False
        page_id = folder_name

    if bool(default_folder) == True:
        Commands.ClickElement(mail_dict["mail_folder"] % folder_name)
        Logs.Logging("Access folder [%s]" % folder_name)
        Waits.WaitUntilPageIsLoaded(mail_dict["page_header"] % page_id)
    else:
        parent_folder_name = Functions.GetElementText(mail_dict["parent_folder"] % folder_name)
        
        try:
            Commands.FindElement(mail_dict["active_parent_folder"] % folder_name)            
            Logs.Logging("Parent folder [%s] is active" % parent_folder_name)
        except WebDriverException:
            Logs.Logging("Parent folder [%s] is not open" % parent_folder_name)

            Commands.ClickElement(mail_dict["open_parent_folder"] % folder_name)
            Waits.WaitElementLoaded(5, mail_dict["active_parent_folder"]  % folder_name)
            Logs.Logging("Click open sub-menu list")
        finally:
            Commands.ClickElement(mail_dict["folder_text"] % folder_name)
            Logs.Logging("Access folder [%s]" % folder_name)
            Waits.WaitUntilPageIsLoaded(mail_dict["subfolder_active"] % folder_name)
    
    Waits.WaitUntilPageIsLoaded(mail_dict["list_footer"])
    time.sleep(1)

    folder_id = Functions.DefineCurrentURL().split("/mail/list/")[1].replace("/", "")
    total_number = DefineListTotal(mail_dict["list_footer"])
    
    folder_dict = {
        "folder_id": folder_id,
        "list_total": total_number
    }

    return folder_dict

def CollectFolderName(custom_folder):
    folder_list = []

    try:
        Waits.WaitElementLoaded(5, mail_dict["sub_menu"] % custom_folder)
    except WebDriverException:
        Commands.ClickElement(mail_dict["open_subfolder"] % custom_folder)
        Logs.Logging("Open folder sub-menu [%s] to view folder list" % custom_folder)
        
        Waits.WaitElementLoaded(3, mail_dict["sub_menu"] % custom_folder)

    try:
        Waits.WaitElementLoaded(5, mail_dict["subfolders_active"] % custom_folder)
        folders = Commands.FindElements(mail_dict["subfolders_active"] % custom_folder)
        for folder in folders:
            folder_list.append(folder.text)
    except WebDriverException:
        Logs.Logging("You don't have folders in this sub-menu [%s]" % custom_folder) # Alert msg

    return folder_list

def AccessMailFolder(folder_name):
    page_unread_dict = None
    custom_folders = ["Folders", "Shared"]

    if folder_name in custom_folders:
        folder_list = CollectFolderName(folder_name)
        if bool(folder_list) == True:
            for folder in folder_list:
                folder_dict = AccessFolderByName(folder)
                page_unread_dict = MailAnalysis(**folder_dict)
    else:
        folder_dict = AccessFolderByName(folder_name)
        page_unread_dict = MailAnalysis(**folder_dict)
    
    return page_unread_dict

def StartDriver():
    global driver
    driver = Driver.StartWebdriver()

    return driver