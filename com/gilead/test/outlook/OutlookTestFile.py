'''
Created on Sep 12, 2018

@author: bkotamahanti
'''
from com.gilead.main.outlook.OutlookMainFile import OutlookMainClass
import inspect
from definitions import email_content, email_subject, OUTLOOK_PROCESS, ROOT_DIR, OUTLOOK_IMAGES_DIR 
from selenium.common.exceptions import NoSuchElementException,\
    WebDriverException
from inspect import getframeinfo, currentframe

class OutlookTestClass():
    __images_dir = ROOT_DIR+OUTLOOK_IMAGES_DIR
    
    def __init__(self, utilsRef, loggerRef):
        self.logger = loggerRef
        self.utils = utilsRef
        self.outlook_maincode = OutlookMainClass(utilsRef, loggerRef)
        
    
    '''TC: W10DA-19 '''            
    def outlook_compose_mail_check_mail_inbox_close_testcase_id_W10DA_19(self): 
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.utils.sleepUntil(5)
            self.logger.info("Launched the WebDriver of Outlook app... ")
            
            
            #TC Step 1
            self.logger.info("Checking and killing the process OUTLOOK.exe if any before executing the test case")
            self.outlook_maincode.killProcess()
            
            #TC Step 2
#             self.outlook_maincode.startOutlook_via_StartMenu()
            self.outlook_maincode.openApplication()
            self.utils.sleepUntil(8)
            assert self.outlook_maincode.isOutlookProcessExist(), "Outlook process OUTLOOK.EXE didn't exist"
            
            ''' Connect the already running Outlook process to winium Driver instead of opening new one.'''
            self.outlook_maincode.connectOutlookToWebDriver()
            self.utils.sleepUntil(8)
            assert self.outlook_maincode.isOutlookElementExist(self.outlook_maincode.getOutlookMainWindowDoctopPaneElement(),**{'Name':'New Email'}), "Outlook app didnt open properly"
            
            self.logger.info("New Email Element is preset")
            
            email_recepient_id = self.outlook_maincode.getLoggedInUserEmailId()
            
            #TC Step 3
            '''commenting below lines to save execution time and it is already covered in same test case below.'''
            ''' For future reference keeping the code.
            self.outlook_maincode.clickNewMail()
            self.logger.info("Clicked on New Email element in outlook")
            self.utils.sleepUntil(5)
            
            assert self.outlook_maincode.isOutlookElementExist(self.outlook_maincode.getOutlookUntitledWindowDoctopPaneElement(), **{'Name':'File Tab'}), "Untitled message window didnt show up"
            
            #TC Step 4
            self.outlook_maincode.clickCloseMenuOptionOfUntitledWindow()
            self.utils.sleepUntil(5)
            assert ( self.outlook_maincode.isOutlookElementExist(self.outlook_maincode.driver, **{'Name':'Untitled - Message (HTML) '})== False ), "Untitled message window didnt close yet"
            self.logger.info("New Compose mail closed successfully")
            assert self.outlook_maincode.isOutlookElementExist(self.outlook_maincode.getOutlookMainWindowDoctopPaneElement(), **{'Name':'New Email'}), "Outlook app didnt open properly"
            self.logger.info("Outlook Main window is showing up now...")
            '''
            #TC Step 5
            self.logger.info("Click New Email again...This time we would try to enter To, Subject and some message")
            self.outlook_maincode.clickNewMail()
            self.logger.info("Clicked New Email element again")
            self.utils.sleepUntil(5)
            assert self.outlook_maincode.isOutlookElementExist(self.outlook_maincode.getOutlookUntitledWindowDoctopPaneElement(), **{'Name':'File Tab'}), "Untitled message window didn't show up"
            
            #TC Step 6
            self.outlook_maincode.enterEmail__id(email_recepient_id)
            self.logger.info("Entered email id ...")
            
            self.outlook_maincode.enterEmail_subject(email_subject)
            self.logger.info("Entered email Subject ...")
            
            self.outlook_maincode.enterEmail_message_and_click_on_Send(email_content)
            self.logger.info("Entered email Content and clicked on Send.....")
            self.utils.sleepUntil(2)
            
            #TC Step 7
            self.utils.sleepUntil(30)
            
            assert self.outlook_maincode.isOutlookSentEmailElementExist(), "Didn't receive email in Inbox...something went wrong"
            self.logger.info("Email is received successfully")
            
            self.outlook_maincode.clickEmailReceivedInOutlookMainWindow()
            self.logger.info("Clicked on received Email")
            
            self.utils.sleepUntil(8)
            self.outlook_maincode.deleteEmailReceived()
            self.utils.sleepUntil(5)
            
            #TC Step 8
            self.outlook_maincode.clickExitMenuOption()
            self.logger.info("clicked on File->Exit of Outlook main window")
            
            #TC Step 9
            self.utils.sleepUntil(8)
            assert self.outlook_maincode.isOutlookProcessExist()==False, "Outlook process OUTLOOK.EXE still exist"
    
            return True, "", func.co_name
        
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
            
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of Outlook app") 
            self.outlook_maincode.killProcess() 
            
            
            
    '''TC ID : W10DA_20 ''' 
    def test_outlook_compose_mail_save_draft_testcase_id_W10DA_20(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the WebDriver of Outlook app... ")
            
            self.utils.sleepUntil(5)
            #TC Step 1
            self.logger.info("Checking and killing the process OUTLOOK.exe if any before executing the test case")
            self.outlook_maincode.killProcess()
            
            #TC Step 2
#             self.outlook_maincode.startOutlook_via_StartMenu()
            self.outlook_maincode.openApplication()
            self.utils.sleepUntil(5)
            assert self.outlook_maincode.isOutlookProcessExist(), "Outlook process OUTLOOK.EXE didn't start properly"
            
            ''' Connect the already running Outlook process to winium Driver instead of opening new one.'''
            self.outlook_maincode.connectOutlookToWebDriver()
            self.utils.sleepUntil(5)
            assert self.outlook_maincode.isOutlookElementExist(self.outlook_maincode.getOutlookMainWindowDoctopPaneElement(),**{'Name':'New Email'}), "Outlook app didnt open properly"
            
            self.logger.info("New Email Element is preset")
            
            email_recepient_id = self.outlook_maincode.getLoggedInUserEmailId()
            
            #TC Step 3
            self.outlook_maincode.clickNewMail()
            self.logger.info("Clicked on New Email element in outlook")
            self.utils.sleepUntil(5)
            
            assert self.outlook_maincode.isOutlookElementExist(self.outlook_maincode.getOutlookUntitledWindowDoctopPaneElement(), **{'Name':'File Tab'}), "Untitled message window didnt show up"
            
            #TC Step 4
            self.outlook_maincode.enterEmail__id(email_recepient_id)
            self.logger.info("Entered email id ...")
            
            self.outlook_maincode.enterEmail_subject(email_subject)
            self.logger.info("Entered email Subject ...")
            
            self.outlook_maincode.enterEmail_message(email_content)
            self.logger.info("Entered email Content .....")
            self.utils.sleepUntil(5)
            
            self.outlook_maincode.clickSaveIconOnUntitledMessageWindow()
            self.utils.sleepUntil(5)
            assert self.outlook_maincode.isOutlookElementExist(self.outlook_maincode.getOutlookUntitledWindowDoctopPaneElement(), **{'Name':'File Tab'}), "Untitled message window didnt show up"
            
            #TC Step 5
            self.outlook_maincode.clickCloseMenuOptionOfUntitledWindow()
            self.utils.sleepUntil(5)
            
            assert self.outlook_maincode.isOutlookElementExist(self.outlook_maincode.getOutlookMainWindowDoctopPaneElement(),**{'Name':'New Email'}), "Outlook app didnt show up properly"
            
            self.outlook_maincode.clickInboxElementIfExpanded()
            self.utils.sleepUntil(5)
            
            #TC Step 6
            self.outlook_maincode.clickOnEmailDraftFolder()
            self.utils.sleepUntil(5)
            
#             assert self.outlook_maincode.isOutlookDraftedEmailElementExist(), "Didn't receive email in Draft...something went wrong"
#             self.logger.info("Email is drafted successfully")
#             
#             self.outlook_maincode.clickEmailDraftedInOutlookMainWindow()
#             self.logger.info("Clicked on drafted Email")
#             self.utils.sleepUntil(5)
#              
#             #TC Step 7
            self.outlook_maincode.clickExitMenuOption()
            self.logger.info("clicked on File->Exit of Outlook main window")
#             
#             #TC Step 8
            self.utils.sleepUntil(5)
            assert self.outlook_maincode.isOutlookProcessExist()==False, "Outlook process OUTLOOK.EXE still exist"
#     
            return True, "", func.co_name
        
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
            
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of Outlook app")
            self.outlook_maincode.killProcess()  
    
    '''TC ID: W10DA-18 '''
    def outlook_check_connectedTo_microsoft_exchange_testcase_id_W10DA_18(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info(
                "Starting execution of function {} :".format(func.co_name))

            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the WebDriver of Outlook app... ")
            self.utils.sleepUntil(5)
            # TC Step 1
            self.logger.info(
                "Checking and killing the process OUTLOOK.exe if any before executing the test case")
            self.outlook_maincode.killProcess()

            # TC Step 2
            
#             self.outlook_maincode.startOutlook_via_StartMenu()
            self.outlook_maincode.openApplication()
            self.utils.sleepUntil(8)
            assert self.outlook_maincode.isOutlookProcessExist(), "Outlook process OUTLOOK.EXE didn't start properly"

            ''' Connect the already running Outlook process to winium Driver instead of opening new one.'''
            self.outlook_maincode.connectOutlookToWebDriver()
            self.utils.sleepUntil(5)
            assert self.outlook_maincode.isOutlookElementExist(self.outlook_maincode.getOutlookMainWindowDoctopPaneElement(), **{'Name': 'New Email'}), "Outlook app didnt open properly"
            
            # TC Step 3
            self.utils.sleepUntil(25)
            assert self.outlook_maincode.checkOutlookMainWindowExchangeServerStatusElement(), "Outlook not connected to Microsoft Exchange"
            self.logger.info("Outlook connected to Microsoft Exchange")
            text = self.outlook_maincode.getTextOfOutlookConnectedToMSExchange()
            assert  text == "Connected to: Microsoft Exchange", "The expected 'Connected to: Microsoft Exchange' didnt match with  actual "+text
            self.utils.sleepUntil(8)

            # TC Step 4
            self.outlook_maincode.clickExitMenuOption()
            self.logger.info("clicked on File->Exit of Outlook main window")

            #TC Step 5
            self.utils.sleepUntil(8)
            assert self.outlook_maincode.isOutlookProcessExist() == False, "Outlook process OUTLOOK.EXE still exist"
            return True, "", func.co_name

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of Outlook app")
            self.outlook_maincode.killProcess()

            
    ''' TC ID: W10DA-132 '''

    def outlook_check_datafile_testcase_id_W10DA_132(self): 
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.utils.sleepUntil(5)
            self.logger.info("Launched the WebDriver of Outlook app... ")
            
            #TC Step 1
            self.logger.info("Checking and killing the process OUTLOOK.exe if any before executing the test case")
            self.outlook_maincode.killProcess()
            
            #TC Step 2
            self.outlook_maincode.openApplication()
            self.utils.sleepUntil(8)
            assert self.outlook_maincode.isOutlookProcessExist(), "Outlook process OUTLOOK.EXE didn't exist"
            
            ''' Connect the already running Outlook process to winium Driver instead of opening new one.'''
            self.outlook_maincode.connectOutlookToWebDriver()
            self.utils.sleepUntil(8)
            assert self.outlook_maincode.isOutlookElementExist(self.outlook_maincode.getOutlookMainWindowDoctopPaneElement(),**{'Name':'New Email'}), "Outlook app didnt open properly"
    
            #TC Step 3
            self.outlook_maincode.clickNewItems()
            self.utils.sleepUntil(5)

            #TC Step 4
            self.outlook_maincode.clickMoreItems()
            self.utils.sleepUntil(5)

            #TC Step 5
            self.outlook_maincode.checkDataFile()
            self.utils.sleepUntil(5)
            self.logger.info(" No Outlook Data File in list")
            return True, "", func.co_name
        
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            self.logger.info("Outlook Data file exists..")
            return False, e, func.co_name  
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of Outlook app") 
            self.outlook_maincode.killProcess() 

    ''' TC ID: W10DA-133 '''

    def outlook_check_if_datafile_exists_testcase_id_W10DA_133(self): 
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            # TC Step 1
            self.outlook_maincode.startProgramByName_via_StartMenu(".pst")
            self.logger.info("Started system settings via start menu")
            self.utils.sleepUntil(5)
            
            # TC Step 2
            assert self.outlook_maincode.check_if_ImageExist(1), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'windows_checking.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.logger.info("Currently we don't have any .pst file in system")
            self.outlook_maincode.close_message_window()
    
            return True, "", func.co_name  
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name

    ''' TC ID: W10DA-134 '''

    def outlook_check_datafile_importexport_testcase_id_W10DA_134(self): 
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.utils.sleepUntil(5)
            self.logger.info("Launched the WebDriver of Outlook app... ")
            
            #TC Step 1
            self.logger.info("Checking and killing the process OUTLOOK.exe if any before executing the test case")
            self.outlook_maincode.killProcess()
            
            #TC Step 2
            self.outlook_maincode.openApplication()
            self.utils.sleepUntil(8)
            assert self.outlook_maincode.isOutlookProcessExist(), "Outlook process OUTLOOK.EXE didn't exist"
            
            ''' Connect the already running Outlook process to winium Driver instead of opening new one.'''
            self.outlook_maincode.connectOutlookToWebDriver()
            self.utils.sleepUntil(8)
            assert self.outlook_maincode.isOutlookElementExist(self.outlook_maincode.getOutlookMainWindowDoctopPaneElement(),**{'Name':'New Email'}), "Outlook app didnt open properly"
            
            #TC Step 3
            self.outlook_maincode.clickFileMenu()
            self.logger.info("clicked on File->Exit of Outlook main window")

            #TC Step 4
            assert self.outlook_maincode.isImageVisible(self.__images_dir + 'outlook_importexport.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'outlook_file.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.outlook_maincode.doubleClickOnImage(self.__images_dir + 'outlook_importexport.PNG', 40, 8 )
            self.utils.sleepUntil(5)

            #TC Step 5
            self.outlook_maincode.clickNextbutton()

            #TC Step 5
            assert self.outlook_maincode.check_if_ImageExist(2), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'windows_checking.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.outlook_maincode.close_window()

            return True, "", func.co_name
        
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            self.logger.info("Outlook Data file exists..")
            return False, e, func.co_name  
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of Outlook app") 
            self.outlook_maincode.killProcess() 
    
    
    ''' TC ID: W10DA-110 '''        
    def outlook_check_version_testcase_id_W10DA_110(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.utils.sleepUntil(5)
            self.logger.info("Launched the WebDriver of Outlook app... ")
            
            
            #TC Step 1
            self.logger.info("Checking and killing the process OUTLOOK.exe if any before executing the test case")
            self.outlook_maincode.killProcess()
            
            #TC Step 2
#             self.outlook_maincode.startOutlook_via_StartMenu()
            self.outlook_maincode.openApplication()
            self.utils.sleepUntil(8)
            assert self.outlook_maincode.isOutlookProcessExist(), "Outlook process OUTLOOK.EXE didn't exist"
            
            ''' Connect the already running Outlook process to winium Driver instead of opening new one.'''
            self.outlook_maincode.connectOutlookToWebDriver()
            self.utils.sleepUntil(8)
            assert self.outlook_maincode.isOutlookElementExist(self.outlook_maincode.getOutlookMainWindowDoctopPaneElement(),**{'Name':'New Email'}), "Outlook app didnt open properly"
            
            self.logger.info("New Email Element is preset")
            
            self.outlook_maincode.clickFileMenuOfficeAccount()
            self.logger.info("clicked on File->office Account of Outlook main window")
            
            #TC Step 9
            self.utils.sleepUntil(8)
            
            assert self.outlook_maincode.isImageVisible( self.__images_dir + 'About_Outlook.PNG'), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'About_Outlook.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            
            self.utils.sleepUntil(8)
            self.outlook_maincode.killProcess() 
            
            #TC Step 8
            self.logger.info("clicked on File->Exit of Outlook main window")
            
            #TC Step 9
            self.utils.sleepUntil(8)
            assert self.outlook_maincode.isOutlookProcessExist()==False, "Outlook process OUTLOOK.EXE still exist"
    
            return True, "", func.co_name
        
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
            
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of Outlook app") 
            self.outlook_maincode.killProcess() 
    
    ''' TC ID: W10DA-111 '''      
    def outlook_offline_settings_testcase_id_W10DA_111(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info(
                "Starting execution of function {} :".format(func.co_name))

            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the WebDriver of Outlook app... ")
            self.utils.sleepUntil(5)
            # TC Step 1
            self.logger.info(
                "Checking and killing the process OUTLOOK.exe if any before executing the test case")
            self.outlook_maincode.killProcess()

            # TC Step 2
            
#             self.outlook_maincode.startOutlook_via_StartMenu()
            self.outlook_maincode.openApplication()
            self.utils.sleepUntil(8)
            assert self.outlook_maincode.isOutlookProcessExist(), "Outlook process OUTLOOK.EXE didn't start properly"

            ''' Connect the already running Outlook process to winium Driver instead of opening new one.'''
            self.outlook_maincode.connectOutlookToWebDriver()
            self.utils.sleepUntil(5)
            assert self.outlook_maincode.isOutlookElementExist(self.outlook_maincode.getOutlookMainWindowDoctopPaneElement(), **{'Name': 'New Email'}), "Outlook app didnt open properly"
            
            # TC Step 3
            self.utils.sleepUntil(25)
            assert self.outlook_maincode.checkOutlookMainWindowExchangeServerStatusElement(), "Outlook not connected to Microsoft Exchange"
            self.logger.info("Outlook connected to Microsoft Exchange")
            text = self.outlook_maincode.getTextOfOutlookConnectedToMSExchange()
            assert  text == "Connected to: Microsoft Exchange", "The expected 'Connected to: Microsoft Exchange' didnt match with  actual "+text
            self.utils.sleepUntil(8)

            # TC Step 4
            self.outlook_maincode.clickExitMenuOption()
            self.logger.info("clicked on File->Exit of Outlook main window")

            #TC Step 5
            self.utils.sleepUntil(8)
            assert self.outlook_maincode.isOutlookProcessExist() == False, "Outlook process OUTLOOK.EXE still exist"
            return True, "", func.co_name

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of Outlook app")
            self.outlook_maincode.killProcess()
    
    ''' TC ID: W10DA-112 ''' 
    def outlook_addins_disabled_testcase_id_W10DA_112(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.utils.sleepUntil(5)
            self.logger.info("Launched the WebDriver of Outlook app... ")
            
            
            #TC Step 1
            self.logger.info("Checking and killing the process OUTLOOK.exe if any before executing the test case")
            self.outlook_maincode.killProcess()
            
            #TC Step 2
#             self.outlook_maincode.startOutlook_via_StartMenu()
            self.outlook_maincode.openApplication()
            self.utils.sleepUntil(8)
            assert self.outlook_maincode.isOutlookProcessExist(), "Outlook process OUTLOOK.EXE didn't exist"
            
            ''' Connect the already running Outlook process to winium Driver instead of opening new one.'''
            self.outlook_maincode.connectOutlookToWebDriver()
            self.utils.sleepUntil(8)
            assert self.outlook_maincode.isOutlookElementExist(self.outlook_maincode.getOutlookMainWindowDoctopPaneElement(),**{'Name':'New Email'}), "Outlook app didnt open properly"
            
            self.logger.info("New Email Element is preset")
            
            #TC Step 8
            self.outlook_maincode.clickFileMenuOptions()
            self.logger.info("clicked on File->options of Outlook main window")
            
            #TC Step 9
            self.utils.sleepUntil(8)
            self.outlook_maincode.clickAddInOption()
            
            self.utils.sleepUntil(8)
            self.outlook_maincode.killProcess() 
            assert self.outlook_maincode.isOutlookProcessExist()==False, "Outlook process OUTLOOK.EXE still exist"
    
            return True, "", func.co_name
        
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
            
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of Outlook app") 
            self.outlook_maincode.killProcess() 
    
              
    ''' TC ID: W10DA-113 '''
    def test_outlook_account_settings_testcase_id_W10DA_113(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.utils.sleepUntil(5)
            self.logger.info("Launched the WebDriver of Outlook app... ")
            
            
            #TC Step 1
            self.logger.info("Checking and killing the process OUTLOOK.exe if any before executing the test case")
            self.outlook_maincode.killProcess()
            
            #TC Step 2
#             self.outlook_maincode.startOutlook_via_StartMenu()
            self.outlook_maincode.openApplication()
            self.utils.sleepUntil(8)
            assert self.outlook_maincode.isOutlookProcessExist(), "Outlook process OUTLOOK.EXE didn't exist"
            
            ''' Connect the already running Outlook process to winium Driver instead of opening new one.'''
            self.outlook_maincode.connectOutlookToWebDriver()
            self.utils.sleepUntil(8)
            assert self.outlook_maincode.isOutlookElementExist(self.outlook_maincode.getOutlookMainWindowDoctopPaneElement(),**{'Name':'New Email'}), "Outlook app didnt open properly"
            
            self.logger.info("New Email Element is preset")
            
            email_recepient_id = self.outlook_maincode.getLoggedInUserEmailId()
            
         
            #TC Step 5
            self.logger.info("Click New Email again...This time we would try to enter To, Subject and some message")
            self.outlook_maincode.clickNewMail()
            self.logger.info("Clicked New Email element again")
            self.utils.sleepUntil(5)
            assert self.outlook_maincode.isOutlookElementExist(self.outlook_maincode.getOutlookUntitledWindowDoctopPaneElement(), **{'Name':'File Tab'}), "Untitled message window didn't show up"
            
            #TC Step 6
            self.outlook_maincode.enterEmail__id(email_recepient_id)
            self.logger.info("Entered email id ...")
            
            self.outlook_maincode.enterEmail_subject(email_subject)
            self.logger.info("Entered email Subject ...")
            
            self.outlook_maincode.enterEmail_message_and_click_on_Send(email_content)
            self.logger.info("Entered email Content and clicked on Send.....")
            self.utils.sleepUntil(2)
            
            #TC Step 7
            self.utils.sleepUntil(30)
            
            assert self.outlook_maincode.isOutlookSentEmailElementExist(), "Didn't receive email in Inbox...something went wrong"
            self.logger.info("Email is received successfully")
            
            self.outlook_maincode.clickEmailReceivedInOutlookMainWindow()
            self.logger.info("Clicked on received Email")
            
            self.utils.sleepUntil(8)
            self.outlook_maincode.deleteEmailReceived()
            self.utils.sleepUntil(5)
            
            #TC Step 8
            self.outlook_maincode.clickExitMenuOption()
            self.logger.info("clicked on File->Exit of Outlook main window")
            
            #TC Step 9
            self.utils.sleepUntil(8)
            assert self.outlook_maincode.isOutlookProcessExist()==False, "Outlook process OUTLOOK.EXE still exist"
    
            return True, "", func.co_name
        
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
            
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of Outlook app") 
            self.outlook_maincode.killProcess() 
    
    ''' TC ID: W10DA-114 '''    
    def test_outlook_account_settings_encryption_testcase_id_W10DA_114(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the WebDriver of Outlook app... ")
            
            self.utils.sleepUntil(5)
            #TC Step 1
            self.logger.info("Checking and killing the process OUTLOOK.exe if any before executing the test case")
            self.outlook_maincode.killProcess()
            
            #TC Step 2
#             self.outlook_maincode.startOutlook_via_StartMenu()
            self.outlook_maincode.openApplication()
            self.utils.sleepUntil(5)
            assert self.outlook_maincode.isOutlookProcessExist(), "Outlook process OUTLOOK.EXE didn't start properly"
            
            ''' Connect the already running Outlook process to winium Driver instead of opening new one.'''
            self.outlook_maincode.connectOutlookToWebDriver()
            self.utils.sleepUntil(5)
            assert self.outlook_maincode.isOutlookElementExist(self.outlook_maincode.getOutlookMainWindowDoctopPaneElement(),**{'Name':'New Email'}), "Outlook app didnt open properly"
            
            self.logger.info("New Email Element is preset")
            
            email_recepient_id = self.outlook_maincode.getLoggedInUserEmailId()
            
            #TC Step 3
            self.outlook_maincode.clickNewMail()
            self.logger.info("Clicked on New Email element in outlook")
            self.utils.sleepUntil(5)
            
            assert self.outlook_maincode.isOutlookElementExist(self.outlook_maincode.getOutlookUntitledWindowDoctopPaneElement(), **{'Name':'File Tab'}), "Untitled message window didnt show up"
            
            #TC Step 4
            self.outlook_maincode.enterEmail__id(email_recepient_id)
            self.logger.info("Entered email id ...")
            
            self.outlook_maincode.enterEmail_subject(email_subject)
            self.logger.info("Entered email Subject ...")
            
            self.outlook_maincode.enterEmail_message(email_content)
            self.logger.info("Entered email Content .....")
            self.utils.sleepUntil(5)
            
            self.outlook_maincode.clickSaveIconOnUntitledMessageWindow()
            self.utils.sleepUntil(5)
            assert self.outlook_maincode.isOutlookElementExist(self.outlook_maincode.getOutlookUntitledWindowDoctopPaneElement(), **{'Name':'File Tab'}), "Untitled message window didnt show up"
            
            #TC Step 5
            self.outlook_maincode.clickCloseMenuOptionOfUntitledWindow()
            self.utils.sleepUntil(5)
            
            assert self.outlook_maincode.isOutlookElementExist(self.outlook_maincode.getOutlookMainWindowDoctopPaneElement(),**{'Name':'New Email'}), "Outlook app didnt show up properly"
            
            self.outlook_maincode.clickInboxElementIfExpanded()
            self.utils.sleepUntil(5)
            
            #TC Step 6
            self.outlook_maincode.clickOnEmailDraftFolder()
            self.utils.sleepUntil(5)
            
#             assert self.outlook_maincode.isOutlookDraftedEmailElementExist(), "Didn't receive email in Draft...something went wrong"
#             self.logger.info("Email is drafted successfully")
#             
#             self.outlook_maincode.clickEmailDraftedInOutlookMainWindow()
#             self.logger.info("Clicked on drafted Email")
#             self.utils.sleepUntil(5)
#              
#             #TC Step 7
            self.outlook_maincode.clickExitMenuOption()
            self.logger.info("clicked on File->Exit of Outlook main window")
#             
#             #TC Step 8
            self.utils.sleepUntil(5)
            assert self.outlook_maincode.isOutlookProcessExist()==False, "Outlook process OUTLOOK.EXE still exist"
#     
            return True, "", func.co_name
        
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
            
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of Outlook app")
            self.outlook_maincode.killProcess()  