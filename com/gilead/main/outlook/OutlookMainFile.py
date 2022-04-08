'''
Created on Aug 24, 2018

@author: bkotamahanti
'''
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException,\
    WebDriverException
from definitions import OUTLOOK_PROCESS, email_subject, ROOT_DIR, OUTLOOK_IMAGES_DIR
from pyscreeze import ImageNotFoundException
import pyautogui
import pyautogui as pg
from selenium.webdriver.common.keys import Keys


class OutlookMainClass():
    __process = OUTLOOK_PROCESS
    __images_dir = ROOT_DIR+OUTLOOK_IMAGES_DIR
    def __init__(self, utilsRef, loggerRef):
        self.utils = utilsRef
        self.logger = loggerRef
        self.driver = ""
    
    
    def connectOutlookToWebDriver(self):
        self.driver = webdriver.Remote(command_executor="http://localhost:9999",
                         desired_capabilities={
                             "app" :  r"C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\OUTLOOK.exe",#Office 365 path
#                              "app" :  r"C:\\Program Files (x86)\\Microsoft Office\\Office16\\OUTLOOK.exe", Office 2016 path
                                               "debugConnectToRunningApp": 'true',
                                               "args": "-port 345"
                             
                             })
        
    
    def killProcess(self):
        self.utils.killProcess(self.__process)
    
    def openApplication(self):
        self.utils.openApplication(self.__process)
    
    def startOutlook_via_StartMenu(self):
        self.utils.startProgramByName_via_StartMenu("Outlook")
    
    def isOutlookProcessExist(self):
        return self.utils.isProcessExist(self.__process)
    
    
    def getLoggedInUserEmailId(self):
        parent_element = self.driver.find_element_by_class_name("rctrl_renwnd32")
        name = parent_element.get_attribute("Name")
        name_list = list(name)
        for index in range(0, len(name_list) ):
            if name_list[index] == '-':
                break
        email = ""
        
        while(name_list[index+1] != '-'):
#         for index in range(start_index_of_email+1, len(name_list)):
#             if( name_list[index] != '-'):
            email =  email + "".join(name_list[index+1])
            index += 1

        return email.strip()    
        
        
    
    
    
    def clickInboxElementIfExpanded(self):
        Mail_folders_element = self.driver.find_element_by_name("Mail Folders")
        elements = Mail_folders_element.find_elements_by_class_name("NetUIWBTreeDisplayNode")
        for element in elements:
            if "Inbox" in element.get_attribute("Name"):
                inbox_element = element
                break
      
        list_of_childelements = inbox_element.find_elements_by_class_name("NetUIWBTreeDisplayNode")
        if len(list_of_childelements) > 1:
            inbox_element.click()
            pyautogui.click(clicks=1)
    
    def getOutlookMainWindowDoctopPaneElement(self):
        parent_element = self.driver.find_element_by_class_name("rctrl_renwnd32")
        MsoDockTop_element = parent_element.find_element_by_name("MsoDockTop")
        return MsoDockTop_element
    
    
    def clickOnEmailDraftFolder(self):
        Mail_folders_element = self.driver.find_element_by_name("Mail Folders")
        elements = Mail_folders_element.find_elements_by_class_name("NetUIWBTreeDisplayNode")
        for element in elements:
            if "Drafts" in element.get_attribute("Name"):
                draft_element = element
                draft_element.click()
    
    
    def checkOutlookMainWindowExchangeServerStatusElement(self):
        doc_bottom = self.driver.find_element_by_name("MsoDockBottom")
        statusbar_element = doc_bottom.find_element_by_class_name("NetUInetpane")
        statusbar_elements = statusbar_element.find_elements_by_class_name("NetUISimpleButton")
        for bar_element in statusbar_elements:
            if "Connected to: Microsoft Exchange" in bar_element.get_attribute("Name"):
                return True
        return False
    
    def getTextOfOutlookConnectedToMSExchange(self):
        doc_bottom = self.driver.find_element_by_name("MsoDockBottom")
        statusbar_element = doc_bottom.find_element_by_class_name("NetUInetpane")
        statusbar_elements = statusbar_element.find_elements_by_class_name("NetUISimpleButton")
        for bar_element in statusbar_elements:
            if "Connected to: Microsoft Exchange" in bar_element.get_attribute("Name"):
                return "Connected to: Microsoft Exchange"
        return None
        
    def getOutlookUntitledWindowDoctopPaneElement(self):
        compose_mail_window_element = self.driver.find_element_by_class_name("rctrl_renwnd32")
        dockTop_element = compose_mail_window_element.find_element_by_name("MsoDockTop")
        return dockTop_element
    
    def isOutlookElementExist(self, webelement , **property1):
        return self.utils.isElementPresent( webelement,  **property1)
    
    def getTodaysEmailGroupElement(self):
        table_view_element = self.getOutlookTableViewElement()
        table_element_today = table_view_element.find_element_by_name("Group By: Expanded: Date: Today")
        return table_element_today
    

    def isOutlookSentEmailElementExist(self):
        table_element_today = self.getTodaysEmailGroupElement()
        elements = table_element_today.find_elements_by_class_name("LeafRow")
        for element in elements:
            print(element.get_attribute("Name"))
            if "This is test Subject" in element.get_attribute("Name"):
                return True
            else:
                continue
        return False

    def getOutlookTableViewElement(self):
        parent_element = self.driver.find_element_by_class_name("rctrl_renwnd32")
#         table_view_element = parent_element.find_element_by_name("Table View")
        table_view_element = parent_element.find_element_by_class_name("SuperGrid")
        return table_view_element
    
    
    def isOutlookDraftedEmailElementExist(self):
        table_view_element = self.getOutlookTableViewElement()
        elements = table_view_element.find_elements_by_class_name("LeafRow")
        for element in elements:
            print(element.get_attribute("Name"))
            if "This is test Subject" in element.get_attribute("Name"):
                return True
            else:
                continue
        return False
    
    
    def clickEmailDraftedInOutlookMainWindow(self):
        table_view_element = self.getOutlookTableViewElement()
        elements = table_view_element.find_elements_by_class_name("LeafRow")
        for element in elements:
            print(element.get_attribute("Name"))
            if "This is test Subject" in element.get_attribute("Name"):
                element.click()
                (a, b) = pyautogui.position()
                pyautogui.rightClick(a, b, 2)
                pyautogui.hotkey('D')
                self.utils.sleepUntil(4)
                pop_up_element = self.driver.find_element_by_class_name("rctrl_renwnd32").find_element_by_class_name("#32770")
                pop_up_element.find_element_by_name("Yes").click()
#                 pyautogui.hotkey('d')
                break
            else:
                continue
        return False
    
    
    def clickNewMail(self):
        MsoDockTop_element = self.getOutlookMainWindowDoctopPaneElement()
        MsoDockTop_element.find_element_by_name("New Email").click()
    
    def clickEmailReceivedInOutlookMainWindow(self):
        table_element_today = self.getTodaysEmailGroupElement()
        elements = table_element_today.find_elements_by_class_name("LeafRow")
        for element in elements:
            print(element.get_attribute("Name"))
            if "This is test Subject" in element.get_attribute("Name"):
                element.click()
                element.submit()
                break
            else:
                continue
    def deleteEmailReceived(self):
        ribbon_pane_element = self.driver.find_element_by_class_name("rctrl_renwnd32").find_element_by_name("MsoDockTop").find_element_by_class_name("NetUInetpane")
        ribbons_tab_element  = ribbon_pane_element.find_element_by_name("Ribbon Tabs")
        ribbons_tab_element.find_element_by_name("Message").click()
        self.utils.sleepUntil(3)
        lower_ribbon_pane_element = ribbon_pane_element.find_element_by_name("Lower Ribbon")
        delete_group_element = lower_ribbon_pane_element.find_element_by_name("Message").find_element_by_name("Delete")
        elements = delete_group_element.find_elements_by_class_name("NetUIRibbonButton")
        for element in elements:
            if element.get_attribute("Name") == "Delete" :
                element.click()
                break

    def clickFileTabInOutlookMainWindow(self):
        self.driver.find_element_by_name("File Tab").click()

    def clickExitMenuOption(self):
        self.clickFileTabInOutlookMainWindow()
        file_menu_element = self.driver.find_element_by_name("Backstage view").find_element_by_name("File")
        self.utils.sleepUntil(3)
        file_menu_element.find_element_by_name("Exit").click()
    
    def clickFileMenuOptions(self):
        self.clickFileTabInOutlookMainWindow()
        file_menu_element = self.driver.find_element_by_name("Backstage view").find_element_by_name("File")
        self.utils.sleepUntil(3)
        file_menu_element.find_element_by_name("Options").click()
    
    def clickAddInOption(self):
        options_pane=self.driver.find_element_by_name("Outlook Options").find_element_by_class_name("NetUIHWNDElement")
        categories_list=options_pane.find_element_by_name("Outlook Options").find_element_by_name("Categories")
        categories_list.find_element_by_name("Add-ins").click()

    def clickFileMenuOfficeAccount(self):
        self.clickFileTabInOutlookMainWindow()
        file_menu_element = self.driver.find_element_by_name("Backstage view").find_element_by_name("File")
        self.utils.sleepUntil(3)
        file_menu_element.find_element_by_name("Office Account").click()
    
    def clickAboutOutlook(self):
        container=self.driver.find_element_by_name("Backstage view").find_element_by_name("Office Account").find_element_by_class_name("NetUISlabContainer")
        container.find_element_by_class_name("NetUIStickyButton").click()
    

    def clickFileTabInUntitledWindow(self):
        MsoDockTop_element = self.getOutlookUntitledWindowDoctopPaneElement()
        MsoDockTop_element.find_element_by_name("File Tab").click()
        self.logger.info("Clicked on New Email - > File Tab")


    def getFileMenuOptionsOfUntitledWindow(self):
        compose_mail_window_element = self.driver.find_element_by_class_name("rctrl_renwnd32")
        file_menu_element = compose_mail_window_element.find_element_by_name("File")
        return file_menu_element

    def clickCloseMenuOptionOfUntitledWindow(self):
        self.clickFileTabInUntitledWindow()
        file_menu_element = self.getFileMenuOptionsOfUntitledWindow()
        file_menu_element.find_element_by_name("Close").click()
        self.logger.info("Clicked on Untitled message window - > File Tab- > Close")
    

    def getEmailFormRegionElement(self):
        compose_mail_window_element = self.driver.find_element_by_class_name("rctrl_renwnd32")
        form_region = compose_mail_window_element.find_element_by_class_name("AfxWndW")
        return form_region

    def enterEmail__id(self, email_id):
        form_region = self.getEmailFormRegionElement() 
        to_element = form_region.find_element_by_name("To")
        to_element.send_keys(email_id)
        # to_element.send_keys(Keys.ENTER)
        # to_element.send_keys(Keys.TAB)
        # to_element.send_keys(Keys.TAB)

    def enterEmail_subject(self, email_subject):
        form_region = self.getEmailFormRegionElement()
        #         subject_element = form_region.find_element_by_xpath("//*[@AutomationId='4101' and @Name='Subject']")
        #         subject_element = form_region.find_element_by_xpath("//*[@AutomationId='4101' and @Name='Subject']")
        #subject_elements = form_region.find_elements_by_name("Subject")
        subject_elements = form_region.find_elements_by_name("Subject")
        #subject_element.send_keys(Keys.TAB)
        #subject_elements.send_keys(Keys.TAB)

        for subject_element in subject_elements:
            if "4101" == subject_element.get_attribute("AutomationId"):
            #if "Edit(50004)" == subject_element.get_attribute("ControlType"):
            #if "RichEdit20WPT" == subject_element.get_attribute("ClassName"):
                subject_element.send_keys(email_subject)
            
        
    
    
    def enterEmail_message_and_click_on_Send(self, email_message):
        form_region = self.getEmailFormRegionElement() 
        message_element = form_region.find_element_by_class_name("_WwG")
        message_element.send_keys(email_message)
        form_region.find_element_by_name("Send").click()
    
    
    def clickSendEmail(self):
        form_region = self.getEmailFormRegionElement() 
        form_region.find_element_by_name("Send").click()
        
        
    def enterEmail_message(self, email_message):
        form_region = self.getEmailFormRegionElement() 
        message_element = form_region.find_element_by_class_name("_WwG")
        message_element.send_keys(email_message)
    
    
    def clickSaveIconOnUntitledMessageWindow(self):
        dockTop_element = self.getOutlookUntitledWindowDoctopPaneElement()
        OuickAccessToolBar_element = dockTop_element.find_element_by_name("Quick Access Toolbar")
        OuickAccessToolBar_element.find_element_by_name("Save").click()
    
    
    def clickEmailSend(self):
        form_region = self.getEmailFormRegionElement()
        send_element = form_region.find_element_by_name("Send") 
#         send_element = form_region.find_element_by_xpath("//*[@AutomationId='4256' and @Name='Send']")
        send_element.click()

    """ Outlook Data File Test Case """

    def clickNewItems(self):
        MsoDockTop_element = self.getOutlookMainWindowDoctopPaneElement()
        MsoDockTop_element.find_element_by_name("New Items").click()

    def clickMoreItems(self):
        MsoDockTop_element = self.getOutlookMainWindowDoctopPaneElement()
        MsoDockTop_element.find_element_by_name("More Items").click()

    def checkDataFile(self):
        try:
            MsoDockTop_element = self.getOutlookMainWindowDoctopPaneElement()
            assert MsoDockTop_element.find_element_by_name("Outlook Data File...").click()
            self.utils.sleepUntil(7)
            return False
            self.logger.info("Outlook Data File is exit")
        except NoSuchElementException:
            return True

    def startProgramByName_via_StartMenu(self, searchText):
        self.utils.startProgramByName_via_StartMenu(searchText)
     

    def check_if_ImageExist(self,state):  
        try:
            if state == 1:
                #a, b, _, _ = pg.locateOnScreen(self.__images_dir + '//file_opening_message.PNG', 10) or pg.locateOnScreen(self.__images_dir + '//Nofile.PNG', 10)
                #self.logger.info("Able to locate image {}/{} at location ( {}, {} )".format(self.__images_dir, '//generic_working_properly.PNG', a, b) )
                self.utils.sleepUntil(2)
                self.logger.info("Currently we don't have any .pst file in system")
                self.close_window()
                return True 
            elif state == 2:
                _, _, _, _ = pg.locateOnScreen(self.__images_dir + 'outlook_pst_disabled_message.PNG', 10)
                self.utils.sleepUntil(2)
                return True
        except ImageNotFoundException as e:
            e.print_stack_trace()
            return False

    def close_window(self):
        pg.hotkey('alt', 'f4')
        self.logger.info("Closed the window with alt+f4 key")

    def isImageVisible(self, scr1): #, scr2, scrollDown=False, state=1):
        a = ""
        b = ""
        try:
            a, b, _, _ = pg.locateOnScreen(scr1, 10)
            return True
            self.logger.info("image {} is located at ( {} , {} )".format(scr1, a, b))
        except:
#             self.close_window()
            self.logger.debug("image {} is  not located".format(scr1))
            return False
    
    
    def doubleClickOnImage(self, src1, offset_x, offset_y,scrollDown=False):
        if scrollDown:
            pg.press('pagedown')
            pg.press('pagedown')
            
        a = ""
        b = ""
        a, b, _, _ = pg.locateOnScreen(src1, 10)
        self.utils.sleepUntil(3)
        pg.click(x=a+offset_x, y=b+offset_y, clicks=1, duration=2)
        self.logger.info("Clicked on image {} at location ( {}, {} )".format(src1, a+offset_x, b+offset_y))

    def clickFileMenu(self):
        self.clickFileTabInOutlookMainWindow()
        file_menu_element = self.driver.find_element_by_name("Backstage view").find_element_by_name("File")
        self.utils.sleepUntil(3)
        file_menu_element.find_element_by_name("Open & Export").click()

    def clickNextbutton(self):
        self.driver.find_element_by_name("Next >").click()
        self.utils.sleepUntil(3)
        self.driver.find_element_by_name("Outlook Data File (.pst)").click()
        self.utils.sleepUntil(3)
        self.driver.find_element_by_name("Next >").click()
        self.utils.sleepUntil(3)

    def close_message_window(self):
        pg.hotkey('Esc')
        self.logger.info("Closed the window with Esc key")
            
        






