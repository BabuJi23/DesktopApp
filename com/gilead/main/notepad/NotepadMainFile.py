'''
Created on Jul 22, 2018

@author: bkotamahanti
'''

import pywinauto
from selenium import webdriver

from _overlapped import NULL
from definitions import NOTEPAD_MAIN_DIR, ROOT_DIR, NOTEPAD_PROCESS

class NotepadMainClass(object):
    minSearchTime = 2
    __process = NOTEPAD_PROCESS
    __notepad_main_dir = NOTEPAD_MAIN_DIR
    __root_dir = ROOT_DIR
        
    def __init__(self, utilsRef, loggerRef):
        self.app = pywinauto.application.Application()
        self.images = self.__root_dir + self.__notepad_main_dir + "images\\"
        self.utils = utilsRef
        self.logger = loggerRef
        self.driver = NULL
        
        
    def startNotepadProcess(self):    
        self.app.start('notepad.exe')
    
    
    def openNotepad(self, filePath):
        self.utils.openApplication(filePath) 
    
       
    def removeTextFile(self, filename):
        self.logger.info("Removing the file..")
        self.utils.removeFile(filename)
    
    def killProcess(self):
        self.utils.killProcess(self.__process) #check their is no process by name notepad.exe using os.  
        
    def isNotepadProcessExist(self):
        return self.utils.isProcessExist("notepad.exe")
    
    
    def closeNotepadApp(self):
        notepad_parent_element = self.driver.find_element_by_class_name("Notepad")
        notepad_application_menubar_element = notepad_parent_element.find_element_by_name("Application")
        notepad_file_menuoption_element = notepad_application_menubar_element.find_element_by_name("File")
        notepad_file_menuoption_element.click()
        notepad_file_menuoption_element.find_element_by_name("Exit").click()
#         self.driver.find_element_by_name("File").click()
#         self.driver.find_element_by_name("Exit").click()
        self.logger.info("Clicked File->Exit menu option ...")
    
    
    def connectNotepadToWebDriver(self):
        self.driver = webdriver.Remote(command_executor='http://localhost:9999',
                                       desired_capabilities={
                                               "debugConnectToRunningApp": "true",
                                               "app" :  r"C:\\Windows\\System32\\notepad.exe",
                                               "args": "-port 345"
                                               }
                                       )
    
    def getCoordinatesByLocatingGivenImageOnScreen(self,imageFile):
        print("==================>"+self.images+imageFile)
        return self.utils.getCoordinatesByLocatingGivenImageOnScreen(self.images+imageFile, self.minSearchTime)
    
    
    def startNotepad_via_StartMenu(self):
        self.utils.startProgramByName_via_StartMenu("notepad")
    
    def enterGivenText(self,line):
        str_list = line.split()
        for i in range(0,len(str_list)):
            self.app.Notepad.edit.type_keys(str_list[i])
            if(i==len(str_list)-1):
                break
            self.app.Notepad.edit.type_keys("{SPACE}")
        self.app.Notepad.edit.type_keys("{ENTER}")
    
    def enterTextFromAnotherFile(self, file):  
        for line in open(self.__root_dir + self.__notepad_main_dir+file,'r'):
            self.enterGivenText(line)
    
    def enterTextFromAnotherFileUsingWebdriver(self, file):
        notepad_parent_element = self.driver.find_element_by_class_name("Notepad")
        notepad_child_text_editor_element = notepad_parent_element.find_element_by_name("Text Editor")
#         fh = open(self.__root_dir + self.__notepad_main_dir+NOTEPAD_MAIN_DIR+file,'r')
#         text = fh.read()
#         fh.close()
        '''not using text as there Webdriver.exe is throwing ConnectionResetError: [WinError 10054] An existing connection was forcibly closed by the remote host
        something wrong with text. As of now entering simple text as below'''
        notepad_child_text_editor_element.click()
        notepad_child_text_editor_element.send_keys("Hello World!")
        self.logger.info("Entering content to a file is done...")
        
    
    def clickMenuOptionSave(self):
        notepad_parent_element = self.driver.find_element_by_class_name("Notepad")
        notepad_application_menubar_element = notepad_parent_element.find_element_by_name("Application")
        notepad_file_menuoption_element = notepad_application_menubar_element.find_element_by_name("File")
        '''clicking menu option File '''
        notepad_file_menuoption_element.click()
        notepad_file_menuoption_element.find_element_by_name("Save").click()
        self.logger.info("clicked the Notepad File->Save option")
    
    def selectMenu(self, menu_option):
        self.app.Notepad.menu_select(menu_option)
    
    def clickSave(self):
        self.app.Notepad.Save.click(double=True)
        self.logger.info("Clicked Save ")
        
    def clickDontSave(self):
        self.app.Notepad.Button2.click()   
    
    def saveFileAs(self, filename):
        notepad_parent_element = self.driver.find_element_by_class_name("Notepad")
        notepad_save_as_dialog_element = notepad_parent_element.find_element_by_name("Save As")
        
        self.driver.find_element_by_class_name("Edit").send_keys(filename)
        notepad_save_as_dialog_element.find_element_by_name("Save").click()
    
    def enterFileNameAndClickSave(self, fileName):
        self.app.SaveAs.Edit.set_edit_text(fileName)
        self.utils.sleepUntil(4)
        self.app.SaveAs.Save.click()
        self.logger.info("Saved the file at location "+fileName)
        
    def enterFileNameAndClickOpen(self, fileName):
        self.app.Open.edit.set_edit_text(fileName)
        self.app.Open.Open.click(double=True)
    
    
    def enterFileNameAndClickOpenViaWebdriver(self, filename):
        open_dialog_element = self.driver.find_element_by_name("Open")
        open_dialog_element.find_element_by_class_name("Edit").send_keys(filename)
        self.logger.info("Entered the file name {}".format(filename))
#         open_dialog_element.find_element_by_class_name("Button").click()
        open_dialog_element.submit()
        self.logger.info("Clicked the Open button in Open dialog")
#         self.driver.find_element_by_class_name("Edit").send_keys(self.__root_dir + self.__notepad_main_dir+filename)
#         self.driver.find_element_by_class_name("Button").click()
#     
    
    def openFile(self, file):
        self.selectMenu("File->Open")
        self.app.Open.edit.set_edit_text(file)
        self.utils.sleepUntil(3)
        self.app.Open.Open.click(double=True)
        
    
    def clickOpenMenuOptionViaWebdriver(self):
        notepad_parent_element = self.driver.find_element_by_class_name("Notepad")
        notepad_application_menubar_element = notepad_parent_element.find_element_by_name("Application")
        notepad_file_menuoption_element = notepad_application_menubar_element.find_element_by_name("File")
        '''clicking menu option File '''
        notepad_file_menuoption_element.click()
        notepad_file_menuoption_element.find_element_by_name("Open...").click()
        
    
    def isPopUpdisplayed(self, property1):
        open_dialog = self.driver.find_element_by_name("Open")
        return self.utils.isElementPresent(open_dialog, {'Name': property1})
    
    
    def clickOK(self):
        self.driver.find_element_by_class_name("CCPushButton").click()
        
    def clickCancel(self):
        self.driver.find_element_by_name("Cancel").click()
        
          
    def getAttribute(self, property1):
        return self.driver.find_element_by_class_name(property1).get_attribute("Name")
    
    def compareInputAndOutputFiles(self, file1, file2):
        flag = True
        with open(self.__root_dir + self.__notepad_main_dir+file1, "r") as f1, open(file2, "r") as f2 :
            for line1, line2 in zip(f1.read(), f2.read()):
                if line1 == line2:
                    continue
                else:
                    flag = False
                    break
        return flag    