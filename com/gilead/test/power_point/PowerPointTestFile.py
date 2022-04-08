'''
Created on Oct 3, 2018

@author: bkotamahanti
'''
from com.gilead.main.power_point.PowerPointMainFile import PowerPointMainClass
import inspect
from definitions import POWERPOINT_FILE_LOCATION1, POWERPOINT_FILE_LOCATION2
from selenium.common.exceptions import NoSuchElementException,\
    WebDriverException

class PowerPointTestClass():

    __file_location1 = POWERPOINT_FILE_LOCATION1
    __file_location2 = POWERPOINT_FILE_LOCATION2
    
    
    def __init__(self, utilsRef, loggerRef):
        self.utils = utilsRef
        self.logger = loggerRef
        self.powerpoint_maincode = PowerPointMainClass(utilsRef, loggerRef)
        self.utils.removeFile(self.__file_location1)
        
    
    '''TC ID: W10DA-36 '''        
    def powerpoint_open_save_close_testcase_id_W10DA_36(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))

            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the WebDriver of Power Point app... ")
            
            # TC Step 1
            self.logger.info(
                "Checking and killing the process POWERPOINT.EXE if any before executing the test case")
            self.powerpoint_maincode.killProcess()
            
            #TC Step 2, 3, 4, 5, 6, 7,8
            self.powerpoint_maincode.createPowerPointPresentation(self.__file_location1)
            self.utils.sleepUntil(3)
            assert self.utils.isFileExists(self.__file_location1) , "Powerpoint ppt {} is not created ".format(self.__file_location1)
            self.logger.info("Powerpoint is created at {}".format(self.__file_location1))
            
            #TC Step 9
            self.utils.openFile(self.__file_location1)
            self.utils.sleepUntil(8)
            assert self.powerpoint_maincode.isPowerPointProcessExist(), "Powerpoint process POWERPOINT.EXE didn't start properly"            
            ''' Connect the already running Powerpoint process to winium Driver instead of opening new one.'''
            self.powerpoint_maincode.connectPowerPointToWebDriver()
            self.logger.info("connected Powerpoint app to WebDriver ")
            
#             assert self.powerpoint_maincode.isPowerPointElementExist(self.powerpoint_maincode.getPowerPointMainWindowSlideElement("scriptCreatedPowerPoint"),**{'Name':'Slide 1 - Adding a Bullet Slide'}), "PowerPoint Element is not present"
            self.logger.info("Checked the basic powerpoint elements present in powerpoint GUI")
            
            #TC Step 10
            self.powerpoint_maincode.killProcess()
            self.utils.sleepUntil(8)
            
            assert self.powerpoint_maincode.isPowerPointProcessExist()==False, "PowerPoint process POWERPOINT.EXE didn't close properly"
            self.logger.info("Closed PowerPoint App")
            return True, "", func.co_name

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of Powerpoint app")
            self.powerpoint_maincode.killProcess()
    
    
    '''TC ID: W10DA-37 '''   
    def powerpoint_open_dont_save_close_testcase_id_W10DA_37(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))

            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the WebDriver of Power Point app... ")
            
            # TC Step 1
            self.logger.info(
                "Checking and killing the process POWERPOINT.EXE if any before executing the test case")
            self.powerpoint_maincode.killProcess()
            
            #TC Step 2
#             self.powerpoint_maincode.startPowerPoint_via_StartMenu()
            self.powerpoint_maincode.openPowerpoint("POWERPNT.EXE")
            self.utils.sleepUntil(8)
            assert self.powerpoint_maincode.isPowerPointProcessExist() , "Powerpoint ppt process didnt start"
            self.logger.info("Powerpoint process started")
            
            
            self.powerpoint_maincode.connectPowerPointToWebDriver()
            self.logger.info("connected Powerpoint app to WebDriver ")
            
            
            #TC Step 3 
            self.powerpoint_maincode.clickBlankPresentation()
            self.logger.info("Clicked Blank Presentaion layout")
            self.utils.sleepUntil(3)
            
            #TC Step 4
            self.powerpoint_maincode.enterTextInTitleTextBox("Hello World!!")
            self.logger.info("Entered Hello World in Title Text Box")
#             self.powerpoint_maincode.enterTextInTextBox("Hello World!!this is text area")
#             self.logger.info("Entered text in Text Box as well")
            
            #TC Step 5
            self.powerpoint_maincode.clickCloseButtonOfPowerPointApp()
            self.logger.info("Clicked Close button of Power point app")
            
            #TC Step 6
            self.powerpoint_maincode.clickDontSaveInMsPopUp()
            self.logger.info("Clicked Dont save in MSPop Up")
            self.utils.sleepUntil(8)

            #TC Step 7
            self.powerpoint_maincode.isPowerPointProcessExist() == False, "Power point process still running"
            return True, "", func.co_name

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of OneNote app")
            self.powerpoint_maincode.killProcess()
            
            
    '''TC ID: W10DA-38 '''        
    def powerpoint_open_save_delete_open_testcase_id_W10DA_38(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))

            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the WebDriver of Power Point app... ")
            
            # TC Step 1
            self.logger.info(
                "Checking and killing the process POWERPOINT.EXE if any before executing the test case")
            self.powerpoint_maincode.killProcess()
            
            #TC Step 2,3,4
            self.powerpoint_maincode.createPowerPointPresentation(self.__file_location2)
            self.utils.sleepUntil(5)
            assert self.utils.isFileExists(self.__file_location2) , "Powerpoint ppt {} is not created ".format(self.__file_location2)
            self.logger.info("Validated Powerpoint ppt is present at {}".format(self.__file_location2))
            
            self.utils.openFile(self.__file_location2)
            self.logger.info("Opened the powerpoint created at {}".format(self.__file_location2))
            
            self.utils.sleepUntil(8)
            assert self.powerpoint_maincode.isPowerPointProcessExist(), "Powerpoint process POWERPOINT.EXE didn't start properly" 
            self.logger.info("Powerpoint process POWERPOINT.EXE is present ")           
            
            ''' Connect the already running Powerpoint process to winium Driver instead of opening new one.'''
            self.powerpoint_maincode.connectPowerPointToWebDriver()
            self.logger.info("connected Powerpoint app to WebDriver ")
            
#             assert self.powerpoint_maincode.isPowerPointElementExist(self.powerpoint_maincode.getPowerPointMainWindowSlideElement("samplePowerPoint2"),**{'Name':'Slide 1 - Adding a Bullet Slide'}), "PowerPoint Element is not present"
            self.logger.info("Checked the basic powerpoint elements present in powerpoint GUI created at {}".format(self.__file_location2))
            
            #TC Step 5
            self.powerpoint_maincode.killProcess()
            self.utils.sleepUntil(2)
            
            #TC Step 6
            self.utils.removeFile(self.__file_location2)
            self.utils.sleepUntil(5)
            
            #TC Step 7
#             self.powerpoint_maincode.startPowerPoint_via_StartMenu()
            self.powerpoint_maincode.openPowerpoint("POWERPNT.EXE")
            self.utils.sleepUntil(5)
            
            self.powerpoint_maincode.clickOpenOtherPresentationsButton()
            self.powerpoint_maincode.clickOpenWindowBrowseButton()
            
            self.powerpoint_maincode.enterFileNameInOpenBrowseWindow(self.__file_location2)
            self.utils.sleepUntil(5)
            self.logger.info("Entered Filename Open browse window as:{}".format(self.__file_location2))
            
            self.powerpoint_maincode.clickOpenButtonOfOpenBrowseWindow()
            self.logger.info("Clicked Browse window->Open button")
            
            
            self.utils.sleepUntil(5)
            assert self.powerpoint_maincode.isPowerPointElementExist(self.powerpoint_maincode.getPowerPointPopUpWindowElement(), **{'Name' : 'OK'}), "The Powerpoint error dialog didn't pop-up"
            self.logger.info("Powerpoint error Pop-up is displayed")
            
            self.utils.sleepUntil(5)
            self.powerpoint_maincode.isPowerPointElementTextExist(self.powerpoint_maincode.getPowerPointPopUpWindowElement(), "File not found" , **{'ClassName' : 'Element'} )
            self.logger.info("Powerpoint error Pop-up displayed 'File not found' msg")
            
            #TC Step 8
            self.powerpoint_maincode.clickOKOnPowepointOpenDialog()
            self.logger.info("Clicked Ok button in file not found error pop-up")
            
            self.utils.sleepUntil(5)
            self.powerpoint_maincode.clickCancelOnPowerpointOpenDialog()
            self.logger.info("Clicked Cancel button in Open dialog pop-up")
            
            self.utils.sleepUntil(5)
            self.powerpoint_maincode.clickPowerpointAppBackButton()
            self.logger.info("Clicked Powerpoint back button")
            
            self.utils.sleepUntil(5)
            self.powerpoint_maincode.clickCloseButtonOfPowerPointAppWhenNoPresentation()
            self.logger.info("Clicked Close button of Powerpoint")
            
            self.utils.sleepUntil(8)
            assert self.powerpoint_maincode.isPowerPointProcessExist()==False, "PowerPoint process POWERPOINT.EXE didn't close properly"
            self.logger.info("Closed PowerPoint App")
            
            return True, "", func.co_name

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of Powerpoint app")
            self.powerpoint_maincode.killProcess()
    
