'''
Created on Sep 12, 2018

@author: bkotamahanti
'''
from com.gilead.main.onenote.OneNoteMainFile import OneNoteMainClass
import inspect
from selenium.common.exceptions import NoSuchElementException,\
    WebDriverException
from definitions import ONENOTE_TEST_DIR, ROOT_DIR, ONENOTE_BINARY_PATH
import shutil
import os
from com.gilead.main.outlook.OutlookMainFile import OutlookMainClass

class OneNoteTestClass():
    __root_dir = ROOT_DIR
    __onenote_test_dir = ONENOTE_TEST_DIR
    __onenote_binary_path = ONENOTE_BINARY_PATH
    def __init__(self, utilsRef, loggerRef):
        self.logger = loggerRef
        self.utils = utilsRef
        self.onenote_maincode = OneNoteMainClass(utilsRef, loggerRef)
        self.outlook_maincode = OutlookMainClass(utilsRef, loggerRef)
        self.onenote_maincode.killProcess()
        
        if(os.path.exists(self.__root_dir+self.__onenote_test_dir+"OneNoteNewFile")):
            shutil.rmtree(self.__root_dir+self.__onenote_test_dir+"OneNoteNewFile")
        if(os.path.exists(self.__root_dir+self.__onenote_test_dir+"OneNoteNewFile2")):
            shutil.rmtree(self.__root_dir+self.__onenote_test_dir+"OneNoteNewFile2")
    
    '''TC ID: W10DA-12 '''        
    def onenote_open_close_testcase_id_W10DA_12(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            '''
#              Sometimes when we enter OneNote in start meni it is Opening Outlook if Outlook is open already.
#                Hence Killing before we execute test case 
            self.outlook_maincode.killProcess()'''
            
            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the WebDriver of OneNote app... ")
            
            # TC Step 1
            self.logger.info(
                "Checking and killing the process ONENOTE.EXE if any before executing the test case")
            self.onenote_maincode.killProcess()
            
            #TC Step 2
#             self.onenote_maincode.startOneNote_via_StartMenu()
            self.onenote_maincode.openApplication(self.__onenote_binary_path)
            self.utils.sleepUntil(5)
            assert self.onenote_maincode.isOneNoteProcessExist(), "OneNote process ONENOTE.EXE didn't start properly"
            self.logger.info("OneNote app started and process ONENOTE.EXE exist in taskmanager")
            
            ''' Connect the already running Outlook process to winium Driver instead of opening new one.'''
            self.onenote_maincode.connectOneNoteToWebDriver()
            self.logger.info("connected Onenote app to WebDriver ")
            self.utils.sleepUntil(3)
            assert self.onenote_maincode.isOneNoteElementExist(self.onenote_maincode.getOneNoteMainWindowDoctopPaneElement(),**{'Name':'File Tab'}), "OneNote app didn't open properly"
            self.logger.info("Checked the basic window elements present in OneNote GUI")
            
            #TC Step 3
            self.onenote_maincode.clickCloseButtonOfOneNoteApp()
            self.utils.sleepUntil(5)
            
            assert self.onenote_maincode.isOneNoteProcessExist()==False, "OneNote process ONENOTE.EXE didn't close properly"
            self.logger.info("Closed OneNote App")
            return True, "", func.co_name

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of OneNote app")
            self.onenote_maincode.killProcess()
    
    
    def onenote_open_save_reopen_testcase_id_W10DA_13(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            '''
#           Sometimes when we enter OneNote in start meni it is Opening Outlook if Outlook is open already.
#           Hence Killing before we execute test case 
            self.outlook_maincode.killProcess()'''
            
            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the WebDriver of OneNote app... ")
            
            # TC Step 1
            self.logger.info(
                "Checking and killing the process ONENOTE.EXE if any before executing the test case")
            self.onenote_maincode.killProcess()
            
            #TC Step 2
#             self.onenote_maincode.startOneNote_via_StartMenu()
            self.onenote_maincode.openApplication(self.__onenote_binary_path)
            self.utils.sleepUntil(5)
            assert self.onenote_maincode.isOneNoteProcessExist(), "OneNote process ONENOTE.EXE didn't start properly"
            self.logger.info("OneNote app started and process ONENOTE.EXE exist in taskmanager")
            
            ''' Connect the already running Outlook process to winium Driver instead of opening new one.'''
            self.onenote_maincode.connectOneNoteToWebDriver()
            self.logger.info("connected Onenote app to WebDriver ")
            self.utils.sleepUntil(3)
            assert self.onenote_maincode.isOneNoteElementExist(self.onenote_maincode.getOneNoteMainWindowDoctopPaneElement(),**{'Name':'File Tab'}), "OneNote app didn't open properly"
            self.logger.info("Checked the basic window elements present in OneNote GUI")
            
            #TC Step 3
            self.onenote_maincode.clickNewOneNote()
            self.logger.info("Cliked New File")
            assert self.onenote_maincode.isOneNoteElementExist(self.onenote_maincode.getOneNoteNewGroupSavingfeaturesElement(), **{'Name' : 'Browse'}), "The New File template element This PC is not visible"
            self.logger.info("Able to find Browse element on New File template")
            self.utils.sleepUntil(2)
            
            #TC Step 4
            self.onenote_maincode.clickBrowseElementInNewTemplate()
            self.logger.info("Clicked First Browse element")
            self.utils.sleepUntil(2)
            assert self.onenote_maincode.isOneNoteElementExist(self.onenote_maincode.getOneNoteNewGroupPickAFolderElement(), **{'ClassName' : 'NetUIStickyButton'}), "Not able to find Browse button to browse and save file"
            self.logger.info("Onenote  2nd Browse element exists in New template")
            #TC Step 5
            self.onenote_maincode.clickSecondBrowseElementInNewTemplate()
            self.logger.info("Clicked 2nd Browse element")
            self.utils.sleepUntil(2)
            
            self.onenote_maincode.enterFileNameInNewBrowseWindow(self.__root_dir+self.__onenote_test_dir+"OneNoteNewFile")
            self.utils.sleepUntil(2)
            self.logger.info("Entered Filename as:"+self.__root_dir+self.__onenote_test_dir+"OneNoteNewFile")
            
            assert self.onenote_maincode.isOneNoteElementExist(self.onenote_maincode.getOneNoteNotebookDropDownElement(), **{'Name': 'OneNoteNewFile'})
            self.logger.info("OneNoteNewFile element is present")
            #TC Step 6
            self.onenote_maincode.enterText("Hello World!! this is sample text")
            self.utils.sleepUntil(2)
            self.logger.info("Entered Text in OneNoteNewFile")
            
            self.onenote_maincode.clickCloseButtonOfOneNoteApp()
            self.utils.sleepUntil(5)
            self.logger.info("Close OneNoteNewFile")
            assert self.onenote_maincode.isOneNoteProcessExist()==False, "OneNote process ONENOTE.EXE didn't close properly"
            self.logger.info("Closed OneNote App")
            
            #TC Step 7
#             self.onenote_maincode.startOneNote_via_StartMenu()
            self.onenote_maincode.openApplication(self.__onenote_binary_path)
            self.logger.info("Started OneNote again")
            self.utils.sleepUntil(5)
            assert self.onenote_maincode.isOneNoteProcessExist(), "OneNote process ONENOTE.EXE didn't start properly"
            self.logger.info("OneNote app started and process ONENOTE.EXE exist in taskmanager")
            assert self.onenote_maincode.isOneNoteElementExist(self.onenote_maincode.getOneNoteNotebookDropDownElement(), **{'Name': 'OneNoteNewFile'})
            self.logger.info("OneNoteNewFile element exists")
            assert self.onenote_maincode.isOneNoteOutlineElementExist('Hello World!! This is sample text'), "Not able to find Text element 'Hello World!! This is sample text' in New File"
            self.logger.info("Verified Entered text is present in OneNoteNewFile")
            self.onenote_maincode.clickCloseButtonOfOneNoteApp()
            self.logger.info("Closed OneNote App")
            self.utils.sleepUntil(5)
            
            assert self.onenote_maincode.isOneNoteProcessExist()==False, "OneNote process ONENOTE.EXE didn't close properly"
            
            self.onenote_maincode.removeFile(self.__root_dir+self.__onenote_test_dir+"OneNoteNewFile")
            self.logger.info("Removed already created file from "+self.__root_dir+self.__onenote_test_dir+"OneNoteNewFile")
            return True, "", func.co_name 
            
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of OneNote app")
            self.onenote_maincode.killProcess()   
    
    def onenote_open_save_delete_reopen_testcase_id_W10DA_14(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))

            '''
#           Sometimes when we enter OneNote in start meni it is Opening Outlook if Outlook is open already.
#           Hence Killing before we execute test case 
            self.outlook_maincode.killProcess()'''
       
            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the WebDriver of OneNote app... ")
            
            # TC Step 1
            self.logger.info(
                "Checking and killing the process ONENOTE.EXE if any before executing the test case")
            self.onenote_maincode.killProcess()
            
            #TC Step 2
#             self.onenote_maincode.startOneNote_via_StartMenu()
            self.onenote_maincode.openApplication(self.__onenote_binary_path)
            self.utils.sleepUntil(5)
            assert self.onenote_maincode.isOneNoteProcessExist(), "OneNote process ONENOTE.EXE didn't start properly"
            self.logger.info("OneNote app started and process ONENOTE.EXE exist in taskmanager")
            
            ''' Connect the already running Outlook process to winium Driver instead of opening new one.'''
            self.onenote_maincode.connectOneNoteToWebDriver()
            self.logger.info("connected Onenote app to WebDriver ")
            self.utils.sleepUntil(3)
            assert self.onenote_maincode.isOneNoteElementExist(self.onenote_maincode.getOneNoteMainWindowDoctopPaneElement(),**{'Name':'File Tab'}), "OneNote app didn't open properly"
            self.logger.info("Checked the basic window elements present in OneNote GUI")
             
             
            #TC Step 3
            self.onenote_maincode.clickNewOneNote()
            self.logger.info("Clicked New File")
            assert self.onenote_maincode.isOneNoteElementExist(self.onenote_maincode.getOneNoteNewGroupSavingfeaturesElement(), **{'Name' : 'Browse'}), "The New File template element This PC is not visible"
            self.logger.info("Able to find Browse element on New File template")
            self.utils.sleepUntil(2)
             
            #TC Step 4
            self.onenote_maincode.clickBrowseElementInNewTemplate()
            self.logger.info("Clicked First Browse element")
            self.utils.sleepUntil(2)
            assert self.onenote_maincode.isOneNoteElementExist(self.onenote_maincode.getOneNoteNewGroupPickAFolderElement(), **{'ClassName' : 'NetUIStickyButton'}), "Not able to find Browse button to browse and save file"
            self.logger.info("Onenote  2nd Browse element exists in New template")
            #TC Step 5
            self.onenote_maincode.clickSecondBrowseElementInNewTemplate()
            self.logger.info("Clicked 2nd Browse element")
            self.utils.sleepUntil(2)
             
            self.onenote_maincode.enterFileNameInNewBrowseWindow(self.__root_dir+self.__onenote_test_dir+"OneNoteNewFile2")
            self.utils.sleepUntil(2)
            self.logger.info("Entered Filename as:"+self.__root_dir+self.__onenote_test_dir+"OneNoteNewFile2")
             
            assert self.onenote_maincode.isOneNoteElementExist(self.onenote_maincode.getOneNoteNotebookDropDownElement(), **{'Name': 'OneNoteNewFile2'})
            self.logger.info("OneNoteNewFile2 element is present")
            #TC Step 6
            self.onenote_maincode.enterText("Hello World!! this is sample text")
            self.utils.sleepUntil(2)
            self.logger.info("Entered Text in OneNoteNewFile")
            
            #TC Step 7
            self.onenote_maincode.clickCloseButtonOfOneNoteApp()
            self.utils.sleepUntil(5)
            self.logger.info("Close OneNoteNewFile2")
            assert self.onenote_maincode.isOneNoteProcessExist()==False, "OneNote process ONENOTE.EXE didn't close properly"
            self.logger.info("Closed OneNote App")
            
            
            #TC Step 8
            '''Removing onenote file before opening '''
            self.logger.info("Removing onenote file before opening")
            self.onenote_maincode.removeFile(self.__root_dir+self.__onenote_test_dir+"OneNoteNewFile2")
            self.logger.info("Removed just created file from: "+self.__root_dir+self.__onenote_test_dir+"OneNoteNewFile2")
            
            #TC Step 9
#             self.onenote_maincode.startOneNote_via_StartMenu()
            self.onenote_maincode.openApplication(self.__onenote_binary_path)
            self.logger.info("Started OneNote again")
            self.utils.sleepUntil(5)
            assert self.onenote_maincode.isOneNoteProcessExist(), "OneNote process ONENOTE.EXE didn't start properly"
            self.logger.info("OneNote app started and process ONENOTE.EXE exist in taskmanager")
            
            #TC Step 10
            self.onenote_maincode.clickOpenOneNoteMenuOption()
            self.logger.info("Clicked File->Open menu option")
            self.utils.sleepUntil(2)
            self.onenote_maincode.clickScollDownButton()
            self.logger.info("Clicked Scroll down button to see Browse button")
            
            self.onenote_maincode.clickBrowseElementOfOpenTemplate()
            self.logger.info("Clicked Browse element")
            self.utils.sleepUntil(2)
            
            self.onenote_maincode.enterFileNameInOpenBrowseWindow(self.__root_dir+self.__onenote_test_dir+"OneNoteNewFile2")
            self.utils.sleepUntil(2)
            self.logger.info("Entered Filename as:"+self.__root_dir+self.__onenote_test_dir+"OneNoteNewFile2")
            
            self.onenote_maincode.clickOpenButtonOfOpenBrowseWindow()
            self.logger.info("Clicked Browse window->Open button")
            
            assert self.onenote_maincode.isOneNoteElementExist(self.onenote_maincode.getOneNoteMicrosoftOneNoteDialog(), **{'Name' : 'OK'}), "The OneNote error dialog didn't pop-up"
            self.logger.info("Pop-up is displayed")
            
            self.onenote_maincode.clickOKOnMicrosoftOneNoteDialog()
            self.logger.info("Clicked Ok button in file not found error pop-up")
            self.onenote_maincode.clickOneNoteAppBackButton()
            self.logger.info("Clicked OneNote back button")
            self.onenote_maincode.clickCloseButtonOfOneNoteApp()
            self.logger.info("Clicked OneNote Close button")
            
            return True, "", func.co_name
           
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of OneNote app") 
            self.onenote_maincode.killProcess()  
    