'''
Created on Oct 23, 2018

@author: bkotamahanti
'''

import inspect
from definitions import EXCEL_FILE_LOCATION1, EXCEL_FILE_LOCATION2, EXCEL_BINARY_PATH
from com.gilead.main.excel.ExcelMainFile import ExcelMainClass
from selenium.common.exceptions import NoSuchElementException,\
    WebDriverException

class ExcelTestClass():
    __file_location1 = EXCEL_FILE_LOCATION1
    __file_location2 = EXCEL_FILE_LOCATION2
    __excel_binary_path= EXCEL_BINARY_PATH
    
    def __init__(self, utilsRef, loggerRef):
        self.utils = utilsRef
        self.logger = loggerRef
        self.excel_maincode = ExcelMainClass(utilsRef, loggerRef)
        self.utils.removeFile(self.__file_location1)

    
    '''TC ID: W10DA-47 '''
    def excel_open_save_close_testcase_id_W10DA_047(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))

            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the WebDriver of MS-Excel app... ")
            
            self.utils.sleepUntil(5)
            # TC Step 1
            self.logger.info(
                "Checking and killing the process EXCEL.EXE if any before executing the test case")
            self.excel_maincode.killProcess()
            
            #TC Step 2, 3, 4, 5, 6, 7,8
            self.excel_maincode.createExcelSheet(self.__file_location1)
            self.logger.info("File created at {}".format(self.__file_location1))
            
            self.utils.sleepUntil(5)
            assert self.utils.isFileExists(self.__file_location1) , "Excel file {} is not created ".format(self.__file_location1)
            self.logger.info("Verified successfully Excel file is located at {}".format(self.__file_location1))
            
            #TC Step 9
            self.utils.sleepUntil(5)
            self.utils.openFile(self.__file_location1)
            self.utils.sleepUntil(8)
            
            assert self.excel_maincode.isExcelProcessExist(), "Excel process EXCEL.EXE didn't start properly"   
            
            ''' Connect the already running Excel process to winium Driver instead of opening new one.'''
            self.excel_maincode.connectExcelToWebDriver()
            self.utils.sleepUntil(5)
            self.logger.info("connected Excel app to WebDriver ")
            
             
            assert self.excel_maincode.isExcelElementExist(self.excel_maincode.driver,**{'ClassName':'NetUIOfficeCaption'}), "Excel GUI Element is not present"
            self.logger.info("Checked the basic Excel elements present in Excel GUI")
            
            
            self.utils.sleepUntil(3)
            sheetnames = self.excel_maincode.getExcelSheetNames(self.__file_location1)
            self.utils.sleepUntil(5)
            assert ['Numbers', 'Names'] == sheetnames, "The sheet names {} didn't match".format(sheetnames)   
            
            worksheet = self.excel_maincode.getWorkSheet(self.__file_location1, "Numbers")
            
            self.utils.sleepUntil(5)
            assert worksheet['C3']._value ==2 , "Data didnt match in excel sheet"
#             TC Step 10
            self.excel_maincode.killProcess()
            self.utils.sleepUntil(8)
             
            assert self.excel_maincode.isExcelProcessExist()==False, "Excel process Excel.EXE didn't close properly"
            self.logger.info("Closed Excel App")
            return True, "", func.co_name

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of Excel app")
            self.excel_maincode.killProcess()
    
    '''TC ID: W10DA-48 '''
    def excel_open_dont_save_close_testcase_id_W10DA_048(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))

            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the WebDriver of MS-Excel app... ")
            self.utils.sleepUntil(5)
            # TC Step 1
            self.logger.info(
                "Checking and killing the process EXCEL.EXE if any before executing the test case")
            self.excel_maincode.killProcess()
            
            #TC Step 2, 3, 4, 5, 6, 7,8
            self.excel_maincode.createExcelSheet(self.__file_location1)
            self.logger.info("File created at {}".format(self.__file_location1))
            
            self.utils.sleepUntil(5)
            assert self.utils.isFileExists(self.__file_location1) , "Excel file {} is not created ".format(self.__file_location1)
            self.logger.info("Verified successfully Excel file is located at {}".format(self.__file_location1))
            
            #TC Step 9
            self.utils.openFile(self.__file_location1)
            self.utils.sleepUntil(5)
            
            assert self.excel_maincode.isExcelProcessExist(), "Excel process EXCEL.EXE didn't start properly"   
            ''' Connect the already running Excel process to winium Driver instead of opening new one.'''
            self.excel_maincode.connectExcelToWebDriver()
            self.logger.info("connected Excel app to WebDriver ")
            
            self.utils.sleepUntil(5)
             
            assert self.excel_maincode.isExcelElementExist(self.excel_maincode.driver,**{'ClassName':'NetUIOfficeCaption'}), "Excel GUI Element is not present"
            self.logger.info("Checked the basic Excel elements present in Excel GUI")
            
#             worksheet = self.excel_maincode.getWorkSheet(self.__file_location1, "Numbers")
            self.excel_maincode.enterDataInWorkSheet("Numbers")
            self.logger.info("Entered new text in Excel file MSExcelDocument.xlsx")
            self.utils.sleepUntil(5)
            
            ''' App is not closing properly when clicking on Close button. Hot key used to don't save functionality is tested'''
            #self.excel_maincode.clickExcelCloseButton()
            self.excel_maincode.clickExcelDontSaveMenuOption()
            self.logger.info("Clicked close button of Excel and Don't Save pop up displayed")
            
            self.utils.sleepUntil(4)
            
            assert self.excel_maincode.isExcelElementExist(self.excel_maincode.getExcelPromptDialogElement(),**{"Name": "Don't Save"}), "Save Dont Save Dialog didn't open"
            self.logger.info("Dialog prompt is shown up successfully ")
            
#             TC Step 10
            self.excel_maincode.killProcess()
            self.utils.sleepUntil(8)
             
            assert self.excel_maincode.isExcelProcessExist()==False, "Excel process Excel.EXE didn't close properly"
            self.logger.info("Closed Excel App")
            return True, "", func.co_name

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of Excel app")
            self.excel_maincode.killProcess()
            
    
    
    '''TC ID: W10DA-49 '''
    def excel_delete_testcase_id_W10DA_049(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))

            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the WebDriver of MS-Excel app... ")
            self.utils.sleepUntil(5)
            # TC Step 1
            self.logger.info(
                "Checking and killing the process EXCEL.EXE if any before executing the test case")
            self.excel_maincode.killProcess()
            
            #TC Step 2, 3, 4, 5, 6, 7,8
            self.excel_maincode.createExcelSheet(self.__file_location2)
            self.logger.info("File created at {}".format(self.__file_location2))
            
            self.utils.sleepUntil(5)
            assert self.utils.isFileExists(self.__file_location2) , "Excel file {} is not created ".format(self.__file_location2)
            self.logger.info("Verified successfully Excel file is located at {}".format(self.__file_location2))
            
            #TC Step 9
            self.utils.openFile(self.__file_location2)
            self.utils.sleepUntil(5)
            
            assert self.excel_maincode.isExcelProcessExist(), "Excel process EXCEL.EXE didn't start properly"   
            
            ''' Connect the already running Excel process to winium Driver instead of opening new one.'''
            self.excel_maincode.connectExcelToWebDriver()
            self.logger.info("connected Excel app to WebDriver ")
            
            self.utils.sleepUntil(5)
             
            assert self.excel_maincode.isExcelElementExist(self.excel_maincode.driver,**{'ClassName':'NetUIOfficeCaption'}), "Excel GUI Element is not present"
            self.logger.info("Checked the basic Excel elements present in Excel GUI")
            
            sheetnames = self.excel_maincode.getExcelSheetNames(self.__file_location2)
            self.utils.sleepUntil(5)
            assert ['Numbers', 'Names'] == sheetnames, "The sheet names {} didn't match".format(sheetnames)  
            
            
            self.logger.info("The sheet names {} matched".format(sheetnames)) 
            
            self.excel_maincode.killProcess()
            self.utils.sleepUntil(5)
            self.utils.removeFile(self.__file_location2)
            self.utils.sleepUntil(3)
            self.logger.info("removed the file {} successfully ".format(self.__file_location2))
            
            self.excel_maincode.openExcel(self.__excel_binary_path)
            self.utils.sleepUntil(5)
            assert self.excel_maincode.isExcelProcessExist(), "Excel process EXCEL.EXE didn't start properly"
            
            self.excel_maincode.clickBlankExcelWorkbook()
            self.logger.info("Clicked Blank Excel layout")
            self.utils.sleepUntil(5)
            
            self.excel_maincode.clickExcelOpenMenuOption()
            self.logger.info("Clicked Excel Open menu option")
            
            self.utils.sleepUntil(8)
            
            self.excel_maincode.clickOpenBrowseButton()
            self.logger.info("Clicked Open Browse option")
            
            self.utils.sleepUntil(8)
            
            self.excel_maincode.enterFileNameInOpenBrowseWindow(self.__file_location2)
            self.utils.sleepUntil(5)
            self.logger.info("Entered Filename Open browse window as:{}".format(self.__file_location2))
            
            self.excel_maincode.clickOpenButtonOfOpenBrowseWindow()
            self.logger.info("Clicked OpenDialog ->Open button")
            
            
            self.utils.sleepUntil(5)
            assert self.excel_maincode.isExcelElementExist(self.excel_maincode.getExcelPopUpWindowElement(), **{'Name' : 'OK'}), "The Excel error dialog didn't pop-up"
            self.logger.info("Excel error Pop-up is displayed")
            
            self.utils.sleepUntil(5)
            self.excel_maincode.isExcelElementTextExist(self.excel_maincode.getExcelPopUpWindowElement(), "File not found" , **{'ClassName' : 'Element'} )
            self.logger.info("Excel Pop-up displayed 'File not found' msg")
            
            #TC Step 8
            self.excel_maincode.clickOKOnExcelOpenDialog()
            self.logger.info("Clicked Ok button in file not found error pop-up")
            
            self.utils.sleepUntil(5)
            self.excel_maincode.clickCancelOnExcelOpenDialog()
            self.logger.info("Clicked Cancel button in Open dialog pop-up")
            
            self.utils.sleepUntil(5)
            self.excel_maincode.clickExceltAppBackButton()
            self.logger.info("Clicked Excel back button")
            
            self.utils.sleepUntil(5)
            ''' App is not closing properly when clicking on Close button. Hence killing the process as the main functionality is tested'''
#             self.excel_maincode.clickCloseButtonOfExcelAppWhenNoPresentation()
#             self.logger.info("Clicked Close button of Excel")
            self.excel_maincode.killProcess()
            self.logger.info("Excel process is closed")
            self.utils.sleepUntil(8)
            assert self.excel_maincode.isExcelProcessExist()==False, "Excel process EXCEL.EXE didn't close properly"
            self.logger.info("Closed Excel App")
            
            return True, "", func.co_name

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of Excel app")
            self.excel_maincode.killProcess()
    