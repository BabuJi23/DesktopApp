import inspect
import os
from definitions import ROOT_DIR, NOTEPAD_TEST_DIR
from pywinauto import MatchError
from selenium.common.exceptions import NoSuchElementException
from com.gilead.main.notepad.NotepadMainFile import NotepadMainClass


class NotepadTestClass():
    __root_dir = ROOT_DIR
    __notepad_test_dir = NOTEPAD_TEST_DIR
    
    def __init__(self, utilsRef, loggerRef):
        self.logger = loggerRef
        self.utils = utilsRef
        if os.path.exists(self.__root_dir+self.__notepad_test_dir+"SavedTextFile.txt"):
            os.remove(self.__root_dir+self.__notepad_test_dir+"SavedTextFile.txt")
        
        if os.path.exists(self.__root_dir+self.__notepad_test_dir+"sample.txt"):
            os.remove(self.__root_dir+self.__notepad_test_dir+"sample.txt")
        
        self.notepad_maincode = NotepadMainClass(utilsRef, loggerRef)
        
    
    '''TC: W10DA-1 '''
    def notepad_save_and_open_file_testcase_id_W10DA_1(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            self.logger.info("Killing Notepad process before executing test case")
            #TC Step 1
            self.notepad_maincode.killProcess()
            
            #TC Step 2
            self.logger.info("Starting Notepad.exe process")
            self.notepad_maincode.startNotepadProcess()
            assert self.notepad_maincode.isNotepadProcessExist(),"Notepad.exe is not started"
            self.logger.info("Notepad process started successfully")
            #TC Step 3
            self.logger.info("Entering some text in Untitled-Notepad")
            self.utils.sleepUntil(2)
            self.notepad_maincode.enterTextFromAnotherFile("inputFile.txt")
            self.logger.info("Entering text is done...")
            #TC Step 4
            self.logger.info("Selecting the Menu item File->Exit..")
            self.utils.sleepUntil(2)
            self.notepad_maincode.selectMenu("File->Exit")
            self.logger.info("Selected File->Exit Menu")
            self.utils.sleepUntil(2)
            '''        
            x_coordinate, y_coordinate = self.notepad.getCoordinatesByLocatingGivenImageOnScreen("notepad_pop_up_image.PNG")
            assert (x_coordinate != 0 and y_coordinate !=0) ,"either of the coordinates are empty" 
            self.logger.info("Able to locate the pop-up image after File->Exit")
            
            '''        
            self.logger.info("clicking Save ")
            #TC Step 5
            self.notepad_maincode.clickSave()
            self.utils.sleepUntil(2)
            #TC Step 6
            self.notepad_maincode.enterFileNameAndClickSave(self.__root_dir+self.__notepad_test_dir+"SavedTextFile.txt")
            #TC Step 7
            self.utils.sleepUntil(2)
            self.notepad_maincode.selectMenu("File->Exit")
            self.logger.info("Closed notepad process after saving")
            assert os.path.exists(self.__root_dir+self.__notepad_test_dir+"SavedTextFile.txt"), "File SavedTextFile.txt doesn't exist in the given file path:\""+self.__root_dir+self.__notepad_test_dir+"\""
            
            self.logger.info("Starting notepad process again")
            #TC Step 8
            self.notepad_maincode.startNotepadProcess(),"Notepad is not started"
            #TC Step 9
            assert self.notepad_maincode.isNotepadProcessExist(),"Notepad.exe is not started"
            #TC Step 10, 11
            self.utils.sleepUntil(2)
            self.notepad_maincode.openFile(self.__root_dir+self.__notepad_test_dir+"SavedTextFile.txt")#remove created files-clean up
            #TC Step 12 
            self.utils.sleepUntil(2)
            assert self.notepad_maincode.compareInputAndOutputFiles("inputFile.txt",self.__root_dir+self.__notepad_test_dir+"SavedTextFile.txt"), "The input file inputFile.txt didnt match with output file SavedTextFile.txt"
            self.logger.info("The input file inputFile.txt matched with output file SavedTextFile.txt")
            #TC Step 13
            self.notepad_maincode.selectMenu("File->Exit")
            self.logger.info("Closed the application....")
            return True, "", func.co_name
        except (AssertionError, AttributeError, TypeError, MatchError) as e:
            return False, e, func.co_name
        finally:
            self.notepad_maincode.killProcess()
    
    
    '''TC: W10DA-2 '''    
    def notepad_dont_save_file_testcase_id_W10DA_2(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            self.logger.info("Killing Notepad process before executing test case")
            #TC Step 1
            self.notepad_maincode.killProcess() #check their is no process by name notepad.exe using os.
        
            #TC Step 2
            self.notepad_maincode.startNotepadProcess()
            assert self.notepad_maincode.isNotepadProcessExist(),"Notepad.exe is not started"
            self.logger.info("Notepad process started successfully")
            #TC Step 3
            self.utils.sleepUntil(2)
            self.logger.info("Entering some text in Untitled-Notepad")
            self.notepad_maincode.enterTextFromAnotherFile("inputFile.txt")
            self.logger.info("Entering text is done...")
            #TC Step 4
            self.logger.info("Selecting the Menu item File->Exit..")
            self.utils.sleepUntil(2)
            self.notepad_maincode.selectMenu("File->Exit")
            self.utils.sleepUntil(2)
            #TC Step 5 
            self.notepad_maincode.clickDontSave()
            self.logger.info("Clicked Dont Save")
            self.utils.sleepUntil(3)
            
            #TC Step 6
            assert (self.notepad_maincode.isNotepadProcessExist() == False), "Notepad.exe file didn't close even after not saving file" 
            return True, "", func.co_name
        
        except (AssertionError, AttributeError, TypeError, MatchError) as e:
            return False, e, func.co_name
        finally:
            self.notepad_maincode.killProcess()
        
        
    '''TC:  W10DA-3 '''    
    def notepad_save_delete_open_file_testcase_id_W10DA_3(self): 
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            self.logger.info("Killing Notepad process before executing test case")
            self.utils.killProcess("notepad.exe") #check their is no process by name notepad.exe using os.
            
            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.utils.sleepUntil(2)
            self.logger.info("Launched the WebDriver of notepad app ")
            
            #TC Step 1
#             self.notepad_maincode.startNotepad_via_StartMenu()
            self.notepad_maincode.openNotepad("notepad.exe")
            
            self.utils.sleepUntil(3)
            assert self.notepad_maincode.isNotepadProcessExist(), "Notepad process notepad.exe didn't exist "
            
            ''' Connect the already running notepad process to winium Driver instead of opening new one.'''
            self.notepad_maincode.connectNotepadToWebDriver()
            self.logger.info("connected webdriver to notepad process...")
            
            #TC Step 2
            self.utils.sleepUntil(2)
            self.logger.info("entering text in notepad")
            self.notepad_maincode.enterTextFromAnotherFileUsingWebdriver("inputFile.txt")
            
            
            #TC Step 3
            self.notepad_maincode.clickMenuOptionSave()
            self.utils.sleepUntil(2)
            assert self.notepad_maincode.getAttribute("#32770") == "Save As", "Save As dialog didn't open"
            
            #TC Step 4
            self.notepad_maincode.saveFileAs(self.__root_dir+self.__notepad_test_dir+"sample.txt")
            
            #TC Step 5
            self.notepad_maincode.closeNotepadApp()
            self.utils.sleepUntil(2)
            
            #TC Step 6
            self.notepad_maincode.removeTextFile(self.__root_dir+self.__notepad_test_dir+"sample.txt")
            
            #TC Step 7
#             self.notepad_maincode.startNotepad_via_StartMenu()
            self.notepad_maincode.openNotepad("notepad.exe")
            self.utils.sleepUntil(2)
            assert self.notepad_maincode.isNotepadProcessExist(), "Notepad process notepad.exe didn't exist "
            self.notepad_maincode.clickOpenMenuOptionViaWebdriver()
            self.utils.sleepUntil(2)
            assert self.notepad_maincode.getAttribute("#32770") == "Open", "Open dialog didn't open"
            
            #TC Step 8
            self.notepad_maincode.enterFileNameAndClickOpenViaWebdriver(self.__root_dir+self.__notepad_test_dir+"sample.txt")
            
            #TC Step 9
            self.utils.sleepUntil(2)
            '''
            ********************Need to work on how to recognize the image accross all resolutions. Commenting for now..************************
            x_coordinate, y_coordinate = self.notepad.getCoordinatesByLocatingGivenImageOnScreen("notepad_pop_up_image_file_not_found.PNG")
            assert (x_coordinate != 0 and y_coordinate !=0) ,"either of the coordinates are empty" 
            self.logger.info("clicking OK ")
            '''
            self.notepad_maincode.clickOK()
            self.logger.info("Clicked on Ok button")
            
            self.notepad_maincode.clickCancel()
            self.logger.info("clicked Cancel")
            self.notepad_maincode.closeNotepadApp()
            self.logger.info("Closed Notepad app..")
            
            return True, "", func.co_name
        except (AssertionError, AttributeError, TypeError, MatchError, NoSuchElementException) as e:
            return False, e, func.co_name
               
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of notepad app")
            self.notepad_maincode.killProcess()
    
