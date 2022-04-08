'''
Created on Oct 24, 2018
@author: pparamasivan
'''

from com.gilead.main.access.AccessMainFile import AccessMainClass
import inspect
from definitions import ACCESS_FILE_LOCATION
from selenium.common.exceptions import NoSuchElementException,\
    WebDriverException
#import pypyodbc as pyodbc


class AccessTestClass():

    __file_location = ACCESS_FILE_LOCATION
        
    def __init__(self, utilsRef, loggerRef):
        self.utils = utilsRef
        self.logger = loggerRef
        self.access_maincode = AccessMainClass(utilsRef, loggerRef)
        self.utils.removeFile(self.__file_location)
        
    
    '''TC ID: W10DA-53 '''        
    def access_open_save_close_testcase_id_W10DA_53(self):        
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))

            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the Winium Driver for Microsoft access application. ")
            
            # TC Step 1
            self.logger.info(
                "Checking and killing the process MSACCESS.EXE if any before executing the test case")
            self.access_maincode.killProcess()
              
            #TC Step 2,3,4,5,6,8
            
            self.access_maincode.createAccessDatabase(self.__file_location)
            self.utils.sleepUntil(3)
            assert self.utils.isFileExists(self.__file_location) , "Access DB {} is not created ".format(self.__file_location)
            self.logger.info("Microsoft Access DB is created at {}".format(self.__file_location))
        
            
            self.utils.openFile(self.__file_location)
            self.utils.sleepUntil(3)
            assert self.access_maincode.isAccessProcessExist(), "Microsoft Access MSACCESS.EXE didn't start properly"            
            ''' Connect the already running Access process to Winium Driver instead of opening new one.'''
            self.access_maincode.connectAccessToWebDriver()
            self.logger.info("connected Access application to WebDriver")
         
            assert self.access_maincode.isAccessElementExist(self.access_maincode.getAccessDatabaseElement(),**{'Name':'Table1'}), "Access Element is not present"
            self.logger.info("Checked the Access DB elements present in Access GUI")          
                 
            #TC Step 9
            self.access_maincode.clickCloseButtonOfAccessApp()
            self.logger.info("Access application Close button was clicked")
            self.utils.sleepUntil(3)
               
            #TC Step 10
            self.access_maincode.isAccessProcessExist() == False, "Microsoft Access process still running"
            return True, "", func.co_name

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of Access application")
            self.access_maincode.killProcess()


            
    '''TC ID: W10DA-54 '''        
    def access_dont_save_close_testcase_id_W10DA_54(self):        
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))

            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the Winium Driver for Microsoft access application. ")
            
            # TC Step 1 kill Access process "msaccess.exe" if it is running
            self.logger.info(
                "Checking and killing the process MSACCESS.EXE if any before executing the test case")
            self.access_maincode.killProcess()

            #TC Step 2 Open MS Access
            self.access_maincode.startMicroSoftAccessApp("MSACCESS.EXE")
            self.utils.sleepUntil(8)

            assert self.access_maincode.isAccessProcessExist() , "Microsoft Access process didn't start"
            self.logger.info("Microsoft Access process started")

            self.access_maincode.connectAccessToWebDriver()
            self.logger.info("connected Access application to WebDriver")

            #Step 3 Click Blank desktop database
            self.access_maincode.clickAccessBlankDatabaseLink()
            self.logger.info("Access DB Blank Database link was clicked")
            
            #Step 4 Click Create button
            self.utils.sleepUntil(5)
            self.access_maincode.clickAccessDatabaseCreateButton()
            self.logger.info("Access DB Create button was clicked")
            self.utils.sleepUntil(3)
                        
            #Step 5 Enter data Name, Department, Details
            assert self.access_maincode.isAccessElementExist(self.access_maincode.getAccessDatabaseTableElement(),**{'Name':'Tables'}), "Access DB  table Element is not present"
            self.logger.info("Checked the Access DB Table element present in Access GUI")                
            self.access_maincode.enterDataInAccessDatabase()
            self.logger.info("Access DB data was entered")

            #TC Step 6  Select Menu > File > Close -- Click 'X' button to close the window
            self.utils.sleepUntil(5)
            self.access_maincode.clickCloseButtonOfAccessApp()
            self.logger.info("Access DB 'Close' button was clicked.")
              
            #Step 7 Click No for warning pop up dialog
            self.utils.sleepUntil(5) 
            self.access_maincode.clickAccessDialogNoButton()
            self.logger.info("Access DB 'No' button was clicked, not to save data")   
       

            #TC Step 8 Check in Task Manager if "msaccess.exe" is running 
            self.access_maincode.isAccessProcessExist() == False, "Microsoft Access process still running"
            return True, "", func.co_name

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of Access application") 
            self.access_maincode.killProcess()   

            
    '''TC ID: W10DA-55 '''                    
    def access_delete_file_validate_error_testcase_id_W10DA_55(self):        
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))

            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the Winium Driver for Microsoft Access application. ")
            
            # TC Step 1
            self.logger.info(
                "Checking and killing the process MSACCESS.EXE if any before executing the test case")
            self.access_maincode.killProcess()
              
            #TC Step 2,3,4,5,6,7
            self.access_maincode.createAccessDatabase(self.__file_location)
            self.utils.sleepUntil(3)
            assert self.utils.isFileExists(self.__file_location) , "Access DB {} is not created ".format(self.__file_location)
            self.logger.info("Microsoft Access DB is created at {}".format(self.__file_location))

            self.utils.openFile(self.__file_location)
            self.utils.sleepUntil(3)
            assert self.access_maincode.isAccessProcessExist(), "Microsoft Access MSACCESS.EXE didn't start properly"            
            ''' Connect the already running Access process to Winium Driver instead of opening new one.'''
            self.access_maincode.connectAccessToWebDriver()
            self.logger.info("connected Access application to WebDriver")
         
            assert self.access_maincode.isAccessElementExist(self.access_maincode.getAccessDatabaseElement(),**{'Name':'Table1'}), "Access Element is not present"
            self.logger.info("Checked the Access DB elements present in Access GUI")          

            #TC Step 8 Open Windows Explorer, give the file path where the file is saved and delete the file
            self.access_maincode.killProcess()
            self.utils.removeFile(self.__file_location)
            self.utils.sleepUntil(8)
            
            #TC Step 9 Open Word and click file name under file location-  Error message should display 
            self.access_maincode.startMicroSoftAccessApp("MSACCESS.EXE")
            self.utils.sleepUntil(5)

            self.access_maincode.clickOpenOtherFilesFolderLink()
            self.access_maincode.clickOpenWindowBrowseButton()
            
            self.access_maincode.enterFileNameInOpenBrowseWindow(self.__file_location)
            self.utils.sleepUntil(5)
            self.logger.info("Entered Filename Open browse window as:{}".format(self.__file_location))
            
            self.access_maincode.clickOpenButtonOfOpenBrowseWindow()
            self.logger.info("Clicked Browse window->Open button")

            self.utils.sleepUntil(5)
            assert self.access_maincode.isAccessElementExist(self.access_maincode.getAccessPopUpWindowElement(), **{'Name' : 'OK'}), "The Access error dialog didn't pop-up"
            self.logger.info("Access error Pop-up is displayed")
            
            self.utils.sleepUntil(5)
            self.access_maincode.isAccessElementTextExist(self.access_maincode.getAccessPopUpWindowElement(), "File not found" , **{'ClassName' : 'Element'} )
            self.logger.info("Access error Pop-up displayed 'File not found' message")
         
            #TC Step 10 Click OK and then close Word clicking 'X' button
            self.access_maincode.clickOKOnAccessOpenDialog()
            self.logger.info("Clicked OK button in file not found error pop-up")
            
            self.utils.sleepUntil(5)
            self.access_maincode.clickCancelOnAccessOpenDialog()
            self.logger.info("Clicked Cancel button in Open dialog pop-up")
            
            self.utils.sleepUntil(5)
            self.access_maincode.clickAccessAppBackButton()
            self.logger.info("Clicked Word back button")

            self.access_maincode.clickCloseButtonAccessApp()
            self.logger.info("Access application Close button was clicked")
                                    
            #TC Step 11 Check in Task Manager if "msaccess.exe" is running 
            self.access_maincode.isAccessProcessExist() == False, "Microsoft Access process still running"
            return True, "", func.co_name

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of Word application") 
            self.access_maincode.killProcess()                               
            