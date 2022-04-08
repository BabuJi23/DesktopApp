'''
Created on Sep 19, 2018

@author: bkotamahanti
'''
from com.gilead.main.acrobat.AcrobatReaderMainFile import AcrobatReaderMainClass
import inspect
from selenium.common.exceptions import NoSuchElementException,\
    WebDriverException
from definitions import ROOT_DIR, ACROBAT_READER_TEST_DIR, ACROBAT_READER_MAIN_DIR, ACROBAT_READER_IMAGES_DIR, \
    ACROBAT_READER_TEST_FILE
from definitions import email_recepient_id
from com.gilead.main.outlook.OutlookMainFile import OutlookMainClass

class AcrobatReaderTestClass(object):
    __root_dir = ROOT_DIR
    __acrobat_test_dir = ACROBAT_READER_TEST_DIR
    __acrobat_main_dir = ACROBAT_READER_MAIN_DIR
    __image_dir = ROOT_DIR + ACROBAT_READER_IMAGES_DIR 
    
    
    def __init__(self, utilsRef, loggerRef):
        self.utils = utilsRef
        self.logger = loggerRef
        self.acrobat_maincode = AcrobatReaderMainClass(utilsRef, loggerRef )
        self.acrobat_maincode.killProcess()
        self.outlook_maincode = OutlookMainClass(utilsRef, loggerRef)
    
    def acrobat_open_read_close_testcase_id_W10DA_09(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))

            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the WebDriver of AcrobatReader app... ")
            
            # TC Step 1
            self.logger.info(
                "Checking and killing the process AcroRd32.exe if any before executing the test case")
            self.acrobat_maincode.killProcess()
            self.utils.sleepUntil(10)
            #TC Step 2
            self.acrobat_maincode.startAcrobatReader_via_StartMenu()
            self.utils.sleepUntil(5)
            assert self.acrobat_maincode.isAcrobatReaderProcessExist(), "Acrobat Reader process AcroRd32.exe didn't start properly"
            self.logger.info("Acrobat Reader app started and process AcroRd32.exe exist in taskmanager")
            
            ''' Connect the already running Acrobat Reader process to winium Driver instead of opening new one.'''
            self.acrobat_maincode.connectAcrobatReaderToWebDriver()
            self.logger.info("connected Acrobat reader app to WebDriver ")
            
            assert self.acrobat_maincode.isAcrobatReaderElementExist(self.acrobat_maincode.getAcrobatReaderMainWindowMenuElement(), **{'Name': 'File'}), "Acrobat Reader app didn't open properly"
            self.logger.info("Checked the basic window elements present in Acrobat Reader GUI")
            
            #TC Step 3
            self.acrobat_maincode.clickAcrobatReaderOpenMenuOption()
            self.logger.info("Clicked AcrobatReaderOpenMenuOption")
            self.utils.sleepUntil(3)
            
            #TC Step 4
            self.acrobat_maincode.enterFileNameAndClickToOpen(ACROBAT_READER_TEST_FILE)
            
            assert self.acrobat_maincode.isAcrobatReaderElementExist(self.acrobat_maincode.driver, **{'Name': 'GIAT_test.pdf - Adobe Acrobat Reader DC'}), "Acrobat reader GIAT_test.pdf didnt open properly"
            self.logger.info("Able to find the GIAT_test.pdf file element")
            #TC Step 5
            self.acrobat_maincode.clickAcrobatReaderClose()
            self.logger.info("Closed Acrobat Reader successfully")
            self.utils.sleepUntil(35)
            
            assert self.acrobat_maincode.isAcrobatReaderProcessExist() == False, "Acrobat Reader process AcroRd32.exe didn't stop yet"
            self.logger.info("Acrobat Reader process AcroRd32.exe doesn't exist in taskmanager anymore which is expected")
            
            return True, "", func.co_name            
        
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of AcrobatReader app")
            self.acrobat_maincode.killProcess()
    
    def acrobat_attach_file_email_testcase_id_W10DA_10(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))

            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the WebDriver of AcrobatReader app... ")
            
            # TC Step 1
            self.logger.info(
                "Checking and killing the process AcroRd32.exe if any before executing the test case")
            self.acrobat_maincode.killProcess()
            
            self.utils.sleepUntil(5)
            #TC Step 5
            self.acrobat_maincode.startAcrobatReader_via_StartMenu()
            self.utils.sleepUntil(3)
            assert self.acrobat_maincode.isAcrobatReaderProcessExist(), "Acrobat Reader process AcroRd32.exe didn't start properly"
            self.logger.info("Acrobat Reader app started and process AcroRd32.exe exist in taskmanager")
            
            ''' Connect the already running Acrobat Reader process to winium Driver instead of opening new one.'''
            self.acrobat_maincode.connectAcrobatReaderToWebDriver()
            self.logger.info("connected Acrobat reader app to WebDriver ")
            
            assert self.acrobat_maincode.isAcrobatReaderElementExist(self.acrobat_maincode.getAcrobatReaderMainWindowMenuElement(), **{'Name': 'File'}), "Acrobat Reader app didn't open properly"
            self.logger.info("Checked the basic window elements present in Acrobat Reader GUI")
            
            #TC Step 3
            self.acrobat_maincode.clickAcrobatReaderOpenMenuOption()
            self.logger.info("Clicked AcrobatReaderOpenMenuOption")
            self.utils.sleepUntil(3)
            
            
            #TC Step 4
            self.acrobat_maincode.enterFileNameAndClickToOpen(ACROBAT_READER_TEST_FILE)
            
            assert self.acrobat_maincode.isAcrobatReaderElementExist(self.acrobat_maincode.driver, **{'Name': 'GIAT_test.pdf - Adobe Acrobat Reader DC'}), "Acrobat reader GIAT_test.pdf didnt open properly"
            self.logger.info("Able to find the GIAT_test.pdf file element")
            
            
            '''Before clicking on Send Email Attach option close outlook to avoid test case failures if they are open before by test case'''
            self.outlook_maincode.killProcess()
            
            #TC Step5
            self.acrobat_maincode.clickShareFileAttachToEmailMenuOption()
            self.logger.info("Clicked on File-> Share File-> Attach to Email option")
            
#             self.utils.sleepUntil(2)
#             assert self.acrobat_maincode.isArcobatReaderElementExist(self.acrobat_maincode.driver,  **{'Name':'Send Email'}), "Acrobat reader Send Email dialog didn't appear"
            
            self.acrobat_maincode.clickSendEmailDialogContinueButton(self.__image_dir+"send_attachment_image.PNG", self.__image_dir+"continue_button_image.PNG")
            self.logger.info("Clicked on Send As Attachement Dialog -> Continue button")
            
            self.utils.sleepUntil(3)
            
            assert self.acrobat_maincode.isAcrobatReaderElementExist(self.acrobat_maincode.driver, **{'Name': 'GIAT_test.pdf - Message (HTML) '}), "Email didn't open"
            self.logger.info("Outlook Email element is present")
            
            '''connect webdriver to Outlook '''
            self.outlook_maincode.connectOutlookToWebDriver()
            self.logger.info("Connected outlook to webdriver..")
            
            #TC Step 6
            self.outlook_maincode.enterEmail__id(email_recepient_id)
            self.logger.info("Entered email id ...")
            
            self.utils.sleepUntil(2)
            self.outlook_maincode.clickSendEmail()
            self.logger.info("Clicked email Send button ...")
            
            self.utils.sleepUntil(2)
            
            #TC Step 7
            self.acrobat_maincode.killProcess()
            
            self.logger.info("Killed Acrobat reader application")
            self.utils.sleepUntil(35)
            
            assert self.acrobat_maincode.isAcrobatReaderProcessExist() == False, "Acrobat Reader process AcroRd32.exe didn't stop yet"
            self.logger.info("Acrobat Reader process AcroRd32.exe doesn't exist in taskmanager anymore which is expected")
            
            
            
            ''' Webdriver is taking some time to close acrobat reader especially when switching bw 2 applications that are running. in this case webdriver
            connected to outlook and then it is switching back to acrobat reader which is taking time to identify the elements.Hence not following below approach and killing the process directly.
            But keeping code for safe side.'''
            
            ''' connect webdriver to Acrobat Reader again'''
#             self.acrobat_maincode.connectAcrobatReaderToWebDriver()
#             assert self.acrobat_maincode.isArcobatReaderElementExist(self.acrobat_maincode.driver,  **{'Name':'GIAT_test.pdf - Adobe Acrobat Reader DC'}), "Acrobat reader GIAT_test.pdf didnt open properly"
#             self.logger.info("Back to GIAT_test.pdf file ")
#             
#             self.acrobat_maincode.clickAcrobatReaderClose()
#             self.logger.info("Closed Acrobat Reader...")
#             
            return True, "", func.co_name            
        
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of AcrobatReader app")
            self.acrobat_maincode.killProcess()
    
    def acrobat_print_cancel_testcase_id_W10DA_11(self): 
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))

            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the WebDriver of AcrobatReader app... ")
            
            self.logger.info(
                "Checking and killing the process AcroRd32.exe if any before executing the test case")
            self.acrobat_maincode.killProcess()
            self.utils.sleepUntil(10)
            
            #TC Step 1
            self.acrobat_maincode.startAcrobatReader_via_StartMenu()
            self.utils.sleepUntil(5)
            assert self.acrobat_maincode.isAcrobatReaderProcessExist(), "Acrobat Reader process AcroRd32.exe didn't start properly"
            self.logger.info("Acrobat Reader app started and process AcroRd32.exe exist in taskmanager")
            
            ''' Connect the already running Acrobat Reader process to winium Driver instead of opening new one.'''
            self.acrobat_maincode.connectAcrobatReaderToWebDriver()
            self.logger.info("connected Acrobat reader app to WebDriver ")
            
            assert self.acrobat_maincode.isAcrobatReaderElementExist(self.acrobat_maincode.getAcrobatReaderMainWindowMenuElement(), **{'Name': 'File'}), "Acrobat Reader app didn't open properly"
            self.logger.info("Checked the basic window elements present in Acrobat Reader GUI")
            
            #TC Step 2
            self.acrobat_maincode.clickAcrobatReaderOpenMenuOption()
            self.logger.info("Clicked AcrobatReaderOpenMenuOption")
            self.utils.sleepUntil(3)
            
            self.acrobat_maincode.enterFileNameAndClickToOpen(ACROBAT_READER_TEST_FILE)
            
            assert self.acrobat_maincode.isAcrobatReaderElementExist(self.acrobat_maincode.driver, **{'Name': 'GIAT_test.pdf - Adobe Acrobat Reader DC'}), "Acrobat reader GIAT_test.pdf didnt open properly"
            self.logger.info("Able to find the GIAT_test.pdf file element")
            
            #TC Step 3
            self.acrobat_maincode.clickAcrobatReaderPrintMenuOption()
            self.utils.sleepUntil(3)
            self.logger.info("Clicked Print menu option of Acrobat Reader")
            
            assert self.acrobat_maincode.isAcrobatReaderElementExist(self.acrobat_maincode.getAcrobatReaderMainWindowElement(), **{'Name': 'Print'}), "Acrobat Reader Print window didn't open properly"
            self.logger.info("Checked the basic Print window element present in Acrobat Reader GUI")
            
            #TC Step 4
            self.acrobat_maincode.clickPrintwindowCancelButton()
            self.logger.info("Clicked Print window cancel button...")
            
            assert self.acrobat_maincode.isAcrobatReaderElementExist(self.acrobat_maincode.driver, **{'Name': 'GIAT_test.pdf - Adobe Acrobat Reader DC'}), "Acrobat reader GIAT_test.pdf didnt open properly"
            self.logger.info("Able to find the GIAT_test.pdf file element")
            
            #TC Step 5
            self.acrobat_maincode.clickAcrobatReaderClose()
            self.logger.info("Closed Acrobat reader application")
            self.utils.sleepUntil(35)
            
            
            #TC Step 6
            assert self.acrobat_maincode.isAcrobatReaderProcessExist() == False, "Acrobat Reader process AcroRd32.exe didn't stop yet"
            self.logger.info("Acrobat Reader process AcroRd32.exe doesn't exist in taskmanager anymore which is expected")
            
            
            return True, "", func.co_name            
        
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of AcrobatReader app")
            self.acrobat_maincode.killProcess()
    