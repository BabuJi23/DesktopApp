'''
Created on Oct 19, 2018

@author: pparamasivan
'''
import inspect
from pywinauto import MatchError
from selenium.common.exceptions import NoSuchElementException
from com.gilead.main.snipping.SnippingMainFile import SnippingMainClass


class SnippingTestClass():

    def __init__(self, utilsRef, loggerRef):
        self.logger = loggerRef
        self.utils = utilsRef
        self.snipping_maincode = SnippingMainClass(utilsRef, loggerRef)
        

    '''TC: W10DA-27 '''
    def snipping_open_close_testcase_id_W10DA_27(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info(
                "Starting execution of function {} :".format(func.co_name))
            
            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the WebDriver of Snipping Tool app... ")
            
            # TC Step 1 : kill Snipping Tool process "SnippingTool.exe" if it
            # is running
            self.logger.info("Killing Snipping Tool process before executing test case")
            self.snipping_maincode.killProcess()

            # TC Step 2 : Open Snipping Tool application via Windows Start Menu
            self.logger.info("Starting SnippingTool.exe process")
            self.snipping_maincode.startSnipping_via_StartMenu()
                        
            self.snipping_maincode.connectSnippingToWebDriver();
            self.logger.info("Snipping App Tool Web driver connected successfully")
            
            assert self.snipping_maincode.isSnippingProcessExist(
            ), "SnippingTool.exe is not started"
            self.logger.info("Snipping process started successfully")

            # TC Step 3 : Click on window close 'X' button
            self.snipping_maincode.clickSnippingToolCloseButton()
            self.logger.info("Snipping App Tool Closed")

            # TC Step 4 : Check in Task Manager if "SnippingTool.exe" is not running"
            self.utils.sleepUntil(2)
            assert self.snipping_maincode.isSnippingProcessExist()==False, "Snipping Tool: SnippingTool.exe process not running"
            
            self.logger.info("Snipping Tool app opened..")
            return True, "", func.co_name
        
        except (AssertionError, AttributeError, TypeError, MatchError) as e:
            return False, e, func.co_name
        
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of Snipping Tool app")
            self.snipping_maincode.killProcess()
            
    '''TC: W10DA-28 '''
    def snipping_new_cancel_option_close_testcase_id_W10DA_28(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info(
                "Starting execution of function {} :".format(func.co_name))
            
            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the WebDriver of Snipping Tool app... ")
            
            # TC Step 1 : kill  process "SnippingTool.exe" if it is running
            self.logger.info("Killing Snipping Tool process before executing test case")
            self.snipping_maincode.killProcess()

            # TC Step 2 : Open Snipping Tool application via Windows Start Menu
            self.logger.info("Starting SnippingTool.exe process")
            self.snipping_maincode.startSnipping_via_StartMenu()
                        
            self.snipping_maincode.connectSnippingToWebDriver();
            self.logger.info("Snipping App Tool Web driver connected successfully")
            
            assert self.snipping_maincode.isSnippingProcessExist(
            ), "SnippingTool.exe is not started"
            self.logger.info("Snipping process started successfully")

            #Step 3 Click on Cancel button - Cancel button grayed out
            self.snipping_maincode.isCancelButtonEnabled()
            assert self.snipping_maincode.isCancelButtonEnabled()==False, "Cancel button Link is enabled"              
            self.logger.info("Verified the Cancel button is Grayed Out")
            
            #Step4  Click Option button - Tools option window open
            self.snipping_maincode.clickSnippingToolOptionButton()
            self.logger.info("Snipping Tool Options link Clicked")
            self.utils.sleepUntil(2)    
                 
            #Step5  Click OK button - window close
            self.snipping_maincode.clickSnippingToolOptionOkButton()
            self.logger.info("Snipping Tool Options dialog OK button Clicked") 
                       
            #Step6  Click New button -check Cancel button enabled
            self.snipping_maincode.clickSnippingToolNewButton()
            self.logger.info("Snipping Tool New link was enabled and Clicked")
            self.utils.sleepUntil(3)
            
            assert self.snipping_maincode.isCancelButtonEnabled(), "Cancel button Link is not enabled"  
            self.logger.info("Verified the Snipping Tool Cancel link was enabled")

            # TC Step 7 : Click 'X' button to close the Snipping Tool window
            self.snipping_maincode.clickSnippingToolCloseButton()
            self.logger.info("Snipping App Tool Closed")

            # TC Step 8 : Check in Task Manager if "SnippingTool.exe" is not running"
            self.utils.sleepUntil(2)
            assert self.snipping_maincode.isSnippingProcessExist()==False, "Snipping Tool: SnippingTool.exe process not running"
            
            self.logger.info("Snipping Tool app opened..")
            return True, "", func.co_name
        
        except (AssertionError, AttributeError, TypeError, MatchError) as e:
            return False, e, func.co_name
        
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of Snipping Tool app")
            self.snipping_maincode.killProcess()
             
            
