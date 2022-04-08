'''
Created on Oct 29, 2018

@author: bkotamahanti
'''

import inspect
from definitions import  PUBLISHER_BINARY_PATH, PUBLISHER_TEXT_TO_ENTER, PUBLISHER_FILE_LOCATION1, PUBLISHER_FILE_LOCATION2
from selenium.common.exceptions import NoSuchElementException,\
    WebDriverException
from com.gilead.main.publisher.PublisherMainFile import PublisherMainClass

class PublisherTestClass():
    
    __publisher_binary_path = PUBLISHER_BINARY_PATH
    __publisher_text_to_enter = PUBLISHER_TEXT_TO_ENTER
    __file_location1 = PUBLISHER_FILE_LOCATION1
    __file_location2 = PUBLISHER_FILE_LOCATION2
    
    def __init__(self, utilsRef, loggerRef):
        self.utils = utilsRef
        self.logger = loggerRef
        self.publisher_maincode = PublisherMainClass(utilsRef, loggerRef)
        self.utils.removeFile(self.__file_location1)
        self.utils.removeFile(self.__file_location2)
    
    '''TC_ID : W10DA_039 '''
    def publisher_open_save_close_testcase_id_W10DA_039(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))

            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the WebDriver of MS-Publisher app... ")
            
            # TC Step 1
            self.logger.info(
                "Checking and killing the process MSPUB.EXE if any before executing the test case")
            self.publisher_maincode.killProcess()
            
            #TC Step 2
            self.publisher_maincode.openApplication(self.__publisher_binary_path)
            self.utils.sleepUntil(2)
            
            assert self.publisher_maincode.isPublisherProcessExist(), "Publisher process MSPUB.EXE didn't exist"
            self.logger.info("Publisher process started in task manager")
            
            #TC Step 3
            ''' Connect the already running Publisher process to winium Driver instead of opening new one.'''
            self.publisher_maincode.connectPublisherToWebDriver()
            assert self.publisher_maincode.isPublisherElementExist(self.publisher_maincode.driver,**{'ClassName':'MSWinPub'}), "Publisher app didnt open properly"
            self.logger.info("Validated element on Publisher GUI successfully")
            
            #TC Step 4
            self.publisher_maincode.clickOnBlankTemplate()
            self.logger.info("Clicked blank template")
            self.utils.sleepUntil(3)
            
            #TC Step 5
            self.publisher_maincode.enterTextOnBlankTemplate(self.__publisher_text_to_enter)
            self.logger.info("Successfully entered text in blank template")
            
            #TC Step 5
            self.publisher_maincode.saveFile(self.__file_location1)
            self.utils.sleepUntil(2)
            
            #TC Step 6
            assert self.utils.isFileExists(self.__file_location1), "Publisher File is not saved to location {}".format(self.__file_location1)
            self.logger.info("Successfully Saved file to location {}".format(self.__file_location1 ))
            
            #TC Step 7
            self.publisher_maincode.killProcess()
            self.logger.info("Closed Publisher app")
            self.utils.sleepUntil(4)
            
            self.publisher_maincode.openApplication(self.__file_location1)
            self.logger.info("Opened Publisher file {}".format(self.__file_location1))
            self.utils.sleepUntil(3)
            '''Not able to validate the text of file as value is not saved in the element because of that element is not returning the text '''
            
            assert self.publisher_maincode.isPublisherElementExist(self.publisher_maincode.driver,**{'ClassName':'MSWinPub'}), "Publisher app didnt open properly"
            self.logger.info("Validated element on Publisher GUI successfully")
            
            #TC Step 8
            self.publisher_maincode.killProcess()
            self.logger.info("Closed Publisher app")
            
            self.utils.sleepUntil(5)
            
            assert self.publisher_maincode.isPublisherProcessExist()== False , "Publisher process MSPUB.EXE still exist"
            return True, "", func.co_name

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of Publisher app")
            self.publisher_maincode.killProcess()
    
    '''TC_ID : W10DA_040 '''
    def publisher_open_dontsave_close_testcase_id_W10DA_040(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))

            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the WebDriver of MS-Publisher app... ")
            
            # TC Step 1
            self.logger.info(
                "Checking and killing the process MSPUB.EXE if any before executing the test case")
            self.publisher_maincode.killProcess()
            
            #TC Step 2
            self.publisher_maincode.openApplication(self.__publisher_binary_path)
            self.utils.sleepUntil(2)
            
            assert self.publisher_maincode.isPublisherProcessExist(), "Publisher process MSPUB.EXE didn't exist"
            self.logger.info("Publisher process started in task manager")
            
            #TC Step 3
            ''' Connect the already running Publisher process to winium Driver instead of opening new one.'''
            self.publisher_maincode.connectPublisherToWebDriver()
            assert self.publisher_maincode.isPublisherElementExist(self.publisher_maincode.driver,**{'ClassName':'MSWinPub'}), "Publisher app didnt open properly"
            self.logger.info("Validated element on Publisher GUI successfully")
            
            #TC Step 4
            self.publisher_maincode.clickOnBlankTemplate()
            self.logger.info("Clicked blank template")
            self.utils.sleepUntil(3)
            
            #TC Step 5
            self.publisher_maincode.enterTextOnBlankTemplate(self.__publisher_text_to_enter)
            self.logger.info("Successfully entered text in blank template")
            
            #TC Step 6
            self.publisher_maincode.clickCloseButtonOfPublisherApp()
            self.utils.sleepUntil(2)
            self.logger.info("Clicked Close button of Publisher app")
            
            #TC Step 7
            self.publisher_maincode.clickDontSaveInPopUp()
            self.logger.info("Clicked Dont save in MS Pop Up")
            self.utils.sleepUntil(4)
            
            assert self.publisher_maincode.isPublisherProcessExist()== False , "Publisher process MSPUB.EXE still exist"
            self.logger.info("Publisher app closed successfully after clicking dont save")
            
            return True, "", func.co_name

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of Publisher app")
            self.publisher_maincode.killProcess()
        
    '''TC_ID : W10DA_041 '''
    def publisher_open_save_delete_open_close_testcase_id_W10DA_041(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))

            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the WebDriver of MS-Publisher app... ")
            
            # TC Step 1
            self.logger.info(
                "Checking and killing the process MSPUB.EXE if any before executing the test case")
            self.publisher_maincode.killProcess()
            
            #TC Step 2
            self.publisher_maincode.openApplication(self.__publisher_binary_path)
            self.utils.sleepUntil(2)
            
            assert self.publisher_maincode.isPublisherProcessExist(), "Publisher process MSPUB.EXE didn't exist"
            self.logger.info("Publisher process started in task manager")
            
            #TC Step 3
            ''' Connect the already running Publisher process to winium Driver instead of opening new one.'''
            self.publisher_maincode.connectPublisherToWebDriver()
            assert self.publisher_maincode.isPublisherElementExist(self.publisher_maincode.driver,**{'ClassName':'MSWinPub'}), "Publisher app didnt open properly"
            self.logger.info("Validated element on Publisher GUI successfully")
            
            #TC Step 4
            self.publisher_maincode.clickOnBlankTemplate()
            self.logger.info("Clicked blank template")
            self.utils.sleepUntil(3)
            
            #TC Step 5
            self.publisher_maincode.enterTextOnBlankTemplate(self.__publisher_text_to_enter)
            self.logger.info("Successfully entered text in blank template")
            
            self.publisher_maincode.saveFile(self.__file_location2)
            self.utils.sleepUntil(2)
            
            #TC Step 6
            assert self.utils.isFileExists(self.__file_location2), "Publisher File is not saved to location {}".format(self.__file_location2)
            self.logger.info("Successfully Saved file to location {}".format(self.__file_location2 ))
            
            #TC Step 7
            self.publisher_maincode.killProcess()
            self.logger.info("Closed Publisher app")
            self.utils.sleepUntil(4)
            
            
            #TC Step 8
            self.utils.removeFile( self.__file_location2 )
            self.utils.sleepUntil(5)
            self.logger.info("removed file from {}".format(self.__file_location2))
            
            
            #TC Step 9
            self.publisher_maincode.openApplication(self.__publisher_binary_path)
            self.utils.sleepUntil(2)
            
            assert self.publisher_maincode.isPublisherProcessExist(), "Publisher process MSPUB.EXE didn't exist"
            self.logger.info("Publisher process started in task manager")
            
            
            self.publisher_maincode.clickOnBlankTemplate()
            self.logger.info("Clicked blank template")
            self.utils.sleepUntil(3)
            
            
            #TC Step 10
            self.publisher_maincode.openPublisherFile( self.__file_location2 )
            self.logger.info("Entered filename and Clicked Browse window->Open button")
            
            self.utils.sleepUntil(3)
            assert self.publisher_maincode.isPublisherElementExist(self.publisher_maincode.getPublisherPopUpWindowElement(), **{'Name' : 'OK'}), "The Publisher error dialog didn't pop-up"
            self.logger.info("Publisher error Pop-up is displayed")
            
            self.utils.sleepUntil(5)
            self.publisher_maincode.isPublisherElementTextExist(self.publisher_maincode.getPublisherPopUpWindowElement(), "File not found" , **{'ClassName' : 'Element'} )
            self.logger.info("Publisher error Pop-up displayed 'File not found' msg")
            
            
            #TC Step 11
            self.publisher_maincode.clickOKOnPublisherOpenDialog()
            self.logger.info("Clicked Ok button in file not found error pop-up")
            
            self.utils.sleepUntil(5)
            self.publisher_maincode.clickCancelOnPublisherOpenDialog()
            self.logger.info("Clicked Cancel button in Open dialog pop-up")
            self.utils.sleepUntil(5)
            
            #TC Step 12
            self.publisher_maincode.killProcess()
            self.logger.info("Killed Power point process")
            self.utils.sleepUntil(5)
            
            assert self.publisher_maincode.isPublisherProcessExist()== False , "Publisher process MSPUB.EXE still exist"
            return True, "", func.co_name

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of Publisher app")
            self.publisher_maincode.killProcess()
    