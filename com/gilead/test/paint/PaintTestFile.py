'''
Created on Sep 12, 2018

@author: bkotamahanti
'''

from definitions import ROOT_DIR, PAINT_TEST_DIR
import os
from com.gilead.main.paint.PaintMainFile import PaintMainClass
import inspect
from selenium.common.exceptions import NoSuchElementException,\
    WebDriverException

class PaintTestClass():
    __root_dir = ROOT_DIR
    __paint_test_dir = PAINT_TEST_DIR
    def __init__(self, utilsRef, loggerRef):
        self.logger = loggerRef
        self.utils = utilsRef
        if os.path.exists(self.__root_dir+self.__paint_test_dir+"SavedPaintFile.png"):
            os.remove(self.__root_dir+self.__paint_test_dir+"SavedPaintFile.png")
        if(os.path.exists(self.__root_dir+self.__paint_test_dir+"shot1.png")):
            os.remove(self.__root_dir+self.__paint_test_dir+"shot1.png")
        if(os.path.exists(self.__root_dir+self.__paint_test_dir+"shot2.png")):
            os.remove(self.__root_dir+self.__paint_test_dir+"shot2.png")    
        
        self.paint_maincode = PaintMainClass(utilsRef, loggerRef)
    
    
    '''TC: W10DA-5 '''
    def paint_open_do_paint_close_save_app_open_testcase_id_W10DA_5(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the WebDriver of Paint app... ")
            
            #TC Step 1
            self.logger.info("Checking and killing the process mspaint.exe before executing the test case")
            self.paint_maincode.killProcess()
            #TC Step 2 
#             self.paint_maincode.startPaint_via_StartMenu()
            self.paint_maincode.openPaint("mspaint.exe")
            self.utils.sleepUntil(2)
            assert self.paint_maincode.isPaintProcessExist(), "Paint process mspaint.exe didn't exist"
            
            ''' Connect the already running paint process to winium Driver instead of opening new one.'''
            self.paint_maincode.connectPaintToWebDriver()
            assert (self.paint_maincode.getPaintAppName()=="Untitled - Paint"), "Untitled Paint didn't open"
            self.logger.info("Paint app opened..")
            
            #TC Step 3 
            self.paint_maincode.maximizePaintApp()
            self.logger.info("maximized Paint window...")
            
            #TC Step 4
            self.paint_maincode.drawSomethingOnEmptyPaintBoard()
            self.logger.info("Take screenshot before saving....")
            
            ''' Take screenshot after drawing something. we will use to compare it later'''
            coordinates = (450, 350, 350, 300)
            self.paint_maincode.takeScreenshot(*coordinates, self.__root_dir+self.__paint_test_dir+"shot1.jpg")
            
            #TC Step 5
            self.paint_maincode.clickExitMenuOptionInPaintApp()
            assert (self.paint_maincode.getTextFromPopUpWindow() == "Do you want to save changes to Untitled?"), "pop-up didn't display"
            self.logger.info("pop-up displayed with Save, Don't Save, Cancel options")
            
            #TC Step 6
            self.paint_maincode.clickSaveOnPopUp()
            self.logger.info("clicked Save button on pop-up window")
            
            ''' validating Save As dialog properties '''
            properties_list = [("#32770","Name"), ("AppControlHost", "Name")]
            #properties_list = [("File name:", "Name"), ("Save as type:", "Name")]
            self.logger.info("Identifying the properties {} of Save as dialog".format(properties_list))
            x,y=self.paint_maincode.getAttributesValuesOfSaveAsDialog(*properties_list)
            assert (x,y) == ("Save As","File name:"), "Values of attributes ({} , {}) didn't match in Save As Dialog with ({}, {}). Something went wrong..".format("Save As","File name:",x,y)
            self.logger.info("Save As dialog displayed...")
            
            #TC Step 7
            self.paint_maincode.enterFileNameAndClickSave(self.__root_dir+self.__paint_test_dir+"SavedPaintFile")
            self.utils.sleepUntil(2)
            assert ( self.paint_maincode.isPaintProcessExist()==False ), "Paint process mspaint.exe still exists even after closing the app"
            self.logger.info("mspaint.exe process is not running any more after closing Paint app.")

            #TC Step 8
#             self.paint_maincode.startPaint_via_StartMenu()
            self.paint_maincode.openPaint("mspaint.exe")
            
            assert self.paint_maincode.isPaintProcessExist(), "Paint process mspaint.exe didn't exist"
            assert (self.paint_maincode.getPaintAppName()=="Untitled - Paint"), "Untitled Paint didn't open"
            self.logger.info("Paint app opened..")
            
            #TC Step 9
            self.paint_maincode.clickOpenMenuOptionViaWebdriver()
            self.utils.sleepUntil(2)
            self.paint_maincode.openPreviousSavedFile("SavedPaintFile.png")
                                                      
            #TC Step 10
            assert (self.paint_maincode.getPaintAppName()=="SavedPaintFile.png - Paint"), "FileName didn't change after Saving File"
            ''' take 2nd screenshot after opening file to compare with prev. image screenshot'''
            self.paint_maincode.takeScreenshot(*coordinates, self.__root_dir+self.__paint_test_dir+"shot2.jpg")
             
            '''compare the images '''
            assert self.paint_maincode.areImagesSame(self.__root_dir+self.__paint_test_dir+"shor1.jpg",self.__root_dir+self.__paint_test_dir+"shot2.jpg"), "Images are not same..."
            self.logger.info("Closing the Paint application...")
            self.paint_maincode.clickExitMenuOptionInPaintApp()
            self.logger.info("Clicked File close button...")
            return True, "", func.co_name
            
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of Paint app")
            self.paint_maincode.killProcess()
            
    
    '''TC: W10DA-6 '''        
    def paint_open_do_paint_close_dont_save_testcase_id_W10DA_6(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the WebDriver of Paint app... ")
            
            #TC Step 1
            self.logger.info("Checking and killing the process mspaint.exe before executing the test case")
            self.paint_maincode.killProcess()
            #TC Step 2 
            
#             self.paint_maincode.startPaint_via_StartMenu()
            self.paint_maincode.openPaint("mspaint.exe")
            self.utils.sleepUntil(2)
            assert self.paint_maincode.isPaintProcessExist(), "Paint process mspaint.exe didn't exist"
            
            ''' Connect the already running paint process to winium Driver instead of opening new one.'''
            self.paint_maincode.connectPaintToWebDriver()
            assert (self.paint_maincode.getPaintAppName()=="Untitled - Paint"), "Untitled Paint didn't open"
            self.logger.info("Paint app opened..")
            
            #TC Step 3
            self.paint_maincode.drawSomethingOnEmptyPaintBoard()
            
            #TC Step 4
            self.paint_maincode.clickExitMenuOptionInPaintApp()
            assert (self.paint_maincode.getTextFromPopUpWindow() == "Do you want to save changes to Untitled?"), "pop-up didn't display"
            self.logger.info("pop-up displayed with Save, Don't Save, Cancel options")
            
            #TC Step 5
            self.paint_maincode.clickDontSaveOnPopUp()
            self.logger.info("clicked Don't Save button on pop-up window")
            self.utils.sleepUntil(2)
            assert ( self.paint_maincode.isPaintProcessExist() == False ), "Paint process still running even after clicking on Dont Save button in Pop up"
            return True, "", func.co_name
        
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
            
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of Paint app")
            self.paint_maincode.killProcess()
            
            
    '''TC: W10DA-7 '''        
    def paint_save_delete_open_file_testcase_id_W10DA_7(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the WebDriver of Paint app... ")
            
            #TC Step 1
            self.logger.info("Checking and killing the process mspaint.exe before executing the test case")
            self.paint_maincode.killProcess()
           
#             self.paint_maincode.startPaint_via_StartMenu()
            self.paint_maincode.openPaint("mspaint.exe")
            self.utils.sleepUntil(2)
            assert self.paint_maincode.isPaintProcessExist(), "Paint process mspaint.exe didn't exist"
            
            ''' Connect the already running paint process to winium Driver instead of opening new one.'''
            self.paint_maincode.connectPaintToWebDriver()
            assert (self.paint_maincode.getPaintAppName()=="Untitled - Paint"), "Untitled Paint didn't open"
            self.logger.info("Paint app opened..")
            
            #TC Step 2
            self.paint_maincode.drawSomethingOnEmptyPaintBoard()
            
            
            #TC Step 3
            self.paint_maincode.clickSaveMenuOptionInPaintApp()
            self.utils.sleepUntil(2)
            
            ''' validating Save As dialog properties '''
            properties_list = [("#32770","Name"), ("AppControlHost", "Name")]
            self.logger.info("Identifying the properties {} of Save as dialog".format(properties_list))
            x,y=self.paint_maincode.getAttributesValuesOfSaveAsDialog(*properties_list)
            assert (x,y) == ("Save As","File name:"), "Values of attributes ({} , {}) didn't match in Save As Dialog with ({}, {}). Something went wrong..".format("Save As","File name:",x,y)
            self.logger.info("Save As dialog displayed...")
            
            #TC Step 4
            self.paint_maincode.enterFileNameAndClickSave(self.__root_dir+self.__paint_test_dir+"SamplePaintFile.png")
            self.utils.sleepUntil(2)
            
            #TC Step 5
            self.logger.info("Closing the Paint application...")
            self.paint_maincode.clickExitMenuOptionInPaintApp()
            self.logger.info("Clicked File close button...")
            self.utils.sleepUntil(2)
            
            assert ( self.paint_maincode.isPaintProcessExist() == False ), "Paint process still running even after closing the app"
            
            #TC Step 6
            self.logger.info("deleting the SamplePaintFile.png file ")
            self.paint_maincode.removePaintFile(self.__root_dir+self.__paint_test_dir+"SamplePaintFile.png")
            
            #TC Step 7
            #self.paint_maincode.startPaint_via_StartMenu()
            self.paint_maincode.openPaint("mspaint.exe")
            self.utils.sleepUntil(2)
            assert self.paint_maincode.isPaintProcessExist(), "Paint process mspaint.exe didn't exist"
            assert (self.paint_maincode.getPaintAppName()=="Untitled - Paint"), "Untitled Paint didn't open"
            self.logger.info("Paint app opened..")
            
            #TC Step 8
            self.paint_maincode.clickOpenMenuOptionViaWebdriver()
            self.utils.sleepUntil(2)
            self.paint_maincode.openPreviousSavedFile(self.__root_dir+self.__paint_test_dir+"SamplePaintFile.png")
            '''
            ********************Need to work on how to recognize the image across all resolutions. Commenting for now..************************
            x_coordinate, y_coordinate = self.paint.getCoordinatesByLocatingGivenImageOnScreen("paint_file_not_found_image.PNG")
            assert (x_coordinate != 0 and y_coordinate !=0) ,"either of the coordinates are empty" 
            '''
            #TC Step 9
            self.utils.sleepUntil(2)
            self.logger.info("clicking OK ")

            self.paint_maincode.clickOK()
            self.logger.info("Clicked on Ok button")
            
            self.paint_maincode.clickCancel()
            self.logger.info("clicked Cancel")
            self.paint_maincode.clickExitMenuOptionInPaintApp()
            self.logger.info("Closed Paint app..")
            
            return True, "", func.co_name
        
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
            
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of Paint app")  
            self.paint_maincode.killProcess()