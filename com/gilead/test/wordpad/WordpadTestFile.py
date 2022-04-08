'''
Created on Oct 30, 2018
@author: pparamasivan
'''

from com.gilead.main.wordpad.WordpadMainFile import WordpadMainClass
import inspect
from definitions import  WORDPAD_FILE_LOCATION, WORDPADFILE
from selenium.common.exceptions import NoSuchElementException,\
    WebDriverException

class WordpadTestClass():

    __file_location = WORDPAD_FILE_LOCATION
    __filename = WORDPAD_FILE_LOCATION + WORDPADFILE
    __wordpadfile = __filename + ".rtf"
       
    def __init__(self, utilsRef, loggerRef):
        self.utils = utilsRef
        self.logger = loggerRef
        self.wordpad_maincode = WordpadMainClass(utilsRef, loggerRef)
        self.utils.removeFile(self.__wordpadfile)
        
    
    '''TC ID: W10DA-50 '''        
    def wordpad_open_save_close_testcase_id_W10DA_050(self):        
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))

            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the Winium Driver for Wordpad application. ")
            
            # TC Step 1
            self.logger.info(
                "Checking and killing the process WORDPAD.EXE if any before executing the test case")
            self.wordpad_maincode.killProcess()
            
            #TC Step 2. Open MS Wordpad
            self.wordpad_maincode.connectWordpadToWebDriver()
            self.logger.info("connected Microsoft Wordpad app to WebDriver ") 
            
            self.wordpad_maincode.startWordpadApp("WORDPAD.EXE")
            self.utils.sleepUntil(3)
            assert self.wordpad_maincode.isWordpadProcessExist(), "The Wordpad process didn't start"
            self.logger.info("Wordpad process started successfully")
            self.utils.sleepUntil(3)

            #TC Step 3. Enter some text
            self.logger.info("Entering some text in Document-Wordpad")
            self.wordpad_maincode.enterTextInWordpadDocument()
            self.utils.sleepUntil(3)
            self.logger.info("Entering text is done...")
 
            #TC Step 4. Select Menu > File > Exit or title bar close button 
            self.wordpad_maincode.clickCloseButtonOfWordpadApp()
            self.logger.info("The wordpad title bar 'Close' button clicked")
            self.utils.sleepUntil(2)
            
            #TC Step 5. Click save button
            self.wordpad_maincode.clickSaveButtonOfWordpadPopUp()
            self.logger.info("clicked Save button ")
            self.utils.sleepUntil(2)
            
            #TC Step 6. Type a file name and then click Save 
            self.wordpad_maincode.enterFileNameInSaveAsWindow(self.__filename)
            self.utils.sleepUntil(3)
            self.wordpad_maincode.clickSaveButtonAfterFileNameEntered()
            self.logger.info("After file entry clicked Save button ")
            
            #TC Step 8,9,10 Open Wordpad again  and Check wordpad running 
            self.wordpad_maincode.startWordpadApp("WORDPAD.EXE")
            self.utils.sleepUntil(3)
            self.logger.info("Started wordpad process again.")
            self.wordpad_maincode.openWordpadDocumentFile(self.__wordpadfile)
            self.utils.sleepUntil(3)
            self.wordpad_maincode.clickWordpadFileOpenButton()
            self.utils.sleepUntil(3)
            assert self.wordpad_maincode.isWordpadProcessExist(), "WORDPAD.EXE didn't start properly"           
            
            #TC Step 11 Verify the file content - Saved content should displayed 
            self.logger.info("Started validating content text.")
            content_text = self.wordpad_maincode.getWordpadDocumentContentText()
            self.logger.info("Doc content is ====> " + content_text)
            
            assert 'This is a Wordpad text' in content_text, "Wordpad doc content is not present"
            self.logger.info("Verified the Wordpad Document Content text")
            
            #TC Step 12  Select Menu > File > Exit - Wordpad closed.
            self.wordpad_maincode.clickCloseWordpadAppButton()
            self.logger.info(" clicked Wordpad application 'Close' button ")
               
            #TC Step 13
            self.wordpad_maincode.isWordpadProcessExist() == False, "Wordpad process still running"
            return True, "", func.co_name

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of Wordpad application")
            
 
    '''TC ID: W10DA-51 '''        
    def wordpad_dont_save_testcase_id_W10DA_051(self):        
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))

            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the Winium Driver for Wordpad application. ")
            
            # TC Step 1
            self.logger.info(
                "Checking and killing the process WORDPAD.EXE if any before executing the test case")
            self.wordpad_maincode.killProcess()
            
            #TC Step 2. Open MS Wordpad
            self.wordpad_maincode.connectWordpadToWebDriver()
            self.logger.info("connected Microsoft Wordpad app to WebDriver ") 

            self.wordpad_maincode.startWordpadApp("WORDPAD.EXE")
            self.utils.sleepUntil(3)
            assert self.wordpad_maincode.isWordpadProcessExist(), "The Wordpad process didn't start"
            self.logger.info("Wordpad process started successfully")
            self.utils.sleepUntil(3)

            #TC Step 3. Enter some text
            self.logger.info("Entering some text in Document-Wordpad")
            self.wordpad_maincode.enterTextInWordpadDocument()
            self.utils.sleepUntil(3)
            self.logger.info("Entering text is done...")
 
            #TC Step 4. Select Menu > File > Exit  - Pop up with three buttons 'Save', 'Don't Save' and 'Cancel' should appear
            self.wordpad_maincode.clickCloseButtonOfWordpadApp()
            self.logger.info("The wordpad title bar 'Close' button clicked")
            self.utils.sleepUntil(3)
            
            #TC Step 5.Click 'Don't Save' button - wordpad closed
            self.wordpad_maincode.clickDontSaveButtonInWordpadPopUp()
            self.logger.info("clicked Don't Save button ")
            self.utils.sleepUntil(3)
            
            #TC Step 6. Check in Task Manager if Wordpad process is running - No Wordpad process "wordpad.exe" in Task Manager.
            self.utils.sleepUntil(3)
            assert self.wordpad_maincode.isWordpadProcessExist() == False, "Wordpad process still running"
            
            return True, "", func.co_name

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of Wordpad application")
            self.wordpad_maincode.killProcess()

            
    '''TC ID: W10DA-52 '''                    
    def wordpad_delete_file_validate_error_testcase_id_W10DA_052(self):        
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))

            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the Winium Driver for Wordpad application. ")
            
            self.logger.info(
                "Checking and killing the process WORDPAD.EXE if any before executing the test case")
            if self.utils.isFileExists(self.__wordpadfile):
                self.utils.removeFile(self.__wordpadfile)
            self.wordpad_maincode.killProcess()
            
            #TC Step 1. Open MS Wordpad
            self.wordpad_maincode.connectWordpadToWebDriver()
            self.logger.info("connected Microsoft Wordpad app to WebDriver ") 

            self.wordpad_maincode.startWordpadApp("WORDPAD.EXE")
            self.utils.sleepUntil(3)
            assert self.wordpad_maincode.isWordpadProcessExist(), "The Wordpad process didn't start"
            self.logger.info("Wordpad process started successfully")
            self.utils.sleepUntil(3)

            #TC Step 2. Enter some text
            self.logger.info("Entering some text in Document-Wordpad")
            self.wordpad_maincode.enterTextInWordpadDocument()
            self.utils.sleepUntil(3)
            self.logger.info("Entering text is done...")
 
            #TC Step 3. Select Menu > File > save 
            self.wordpad_maincode.clickWordpadToolbarSaveButton()
            self.logger.info("The wordpad title bar 'Save' button clicked")
            self.utils.sleepUntil(3)

            #TC Step 4. Type a file name and then click Save 
            self.wordpad_maincode.enterFileNameInSaveAsWindow(self.__filename)
            self.utils.sleepUntil(3)
            self.wordpad_maincode.clickSaveButtonAfterFileNameEntered()
            self.logger.info("After file entry clicked Save button ")

            #TC Step 5 File > Exit Or Close the application
            self.wordpad_maincode.clickCloseWordpadAppButton()
            self.logger.info(" clicked Wordpad application 'Close' button ")
            self.utils.sleepUntil(3)        
 
            #TC Step 6 Open Windows Explorer, give the file path where the file is saved and delete the file
            self.utils.removeFile(self.__wordpadfile)
            self.utils.sleepUntil(5)
            
            #TC Step 7 Open Wordpad and click file name under Recent - error message should display

            self.wordpad_maincode.startWordpadApp("WORDPAD.EXE")
            self.utils.sleepUntil(3)
            self.logger.info("Started wordpad process again.")
            self.wordpad_maincode.openWordpadDocumentFile(self.__wordpadfile)
            self.utils.sleepUntil(3)
            self.wordpad_maincode.clickWordpadFileOpenButton()
            self.utils.sleepUntil(3)

            assert self.wordpad_maincode.isWordpadElementExist(self.wordpad_maincode.getWordpadPopUpWindowElement(), **{'Name' : 'OK'}), "The Wordpad error dialog didn't pop-up"
            self.logger.info("Wordpad error Pop-up is displayed")
            
            self.utils.sleepUntil(3)
            assert self.wordpad_maincode.isWordpadElementTextExist(self.wordpad_maincode.getWordpadPopUpWindowElement(), "File not found" , **{'ClassName' : 'Element'} )
            self.logger.info("Wordpad error Pop-up displayed 'File not found' message")
         
            #TC Step 8 Click OK and then close Word clicking 'X' button
            self.wordpad_maincode.clickOKOnWordpadOpenDialog()
            self.logger.info("Clicked OK button in file not found error pop-up")
            
            self.utils.sleepUntil(3)
            self.wordpad_maincode.clickCancelOnWordpadOpenDialog()
            self.logger.info("Clicked Cancel button in Open dialog pop-up")
                            
            self.wordpad_maincode.clickCloseButtonOfWordpadApp()
            self.utils.sleepUntil(3)
               
            self.wordpad_maincode.isWordpadProcessExist() == False, "Microsoft Wordpad process still running"
            return True, "", func.co_name

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of Word application")
            self.wordpad_maincode.killProcess()
