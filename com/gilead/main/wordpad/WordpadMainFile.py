'''
Created on Oct 30, 2018
@author: pparamasivan
'''

from definitions import WORDPAD_PROCESS
from selenium import webdriver
import pyautogui as pg


class WordpadMainClass():
    __process = WORDPAD_PROCESS

    
    def __init__(self, utilsRef, loggerRef):
        self.utils = utilsRef
        self.logger = loggerRef
        self.driver = ""
        self.utils.killProcess(self.__process)
    
    def startWordpadApp(self, file):    
        self.utils.openApplication(file)

    def removeWordpadFile(self, filename):
        self.logger.info("Removing the file..")
        self.utils.removeFile(filename)
    
    def killProcess(self):
        self.utils.killProcess(self.__process)
        
    def enterTextInWordpadDocument(self):
        parent_element = self.driver.find_element_by_class_name("WordPadClass")
        text_element = parent_element.find_element_by_name("Rich Text Window")
        text_element.click()
        text_element.send_keys("This is a Wordpad text")
 
    def enterFileNameInSaveAsWindow(self, filename):
        pg.hotkey('alt','n')
        pg.typewrite(filename)
        return filename
    
    def clickSaveButtonAfterFileNameEntered(self):
        parent_element = self.driver.find_element_by_class_name("WordPadClass")
        saveas_element = parent_element.find_element_by_name("Save As")
        saveas_element.find_element_by_name("Save").click()
        
    def openWordpadDocumentFile(self, filename):  
        pg.hotkey('alt','f')
        pg.hotkey('alt','o')
        self.utils.sleepUntil(3)
        pg.hotkey('alt','n')
        self.utils.sleepUntil(3)
        pg.typewrite(filename)
        return filename
    
    def clickWordpadFileOpenButton(self):
        pg.hotkey('alt','o')
        
    def clickWordpadToolbarSaveButton(self):
        pg.hotkey('ctrl','s')    
        
    def clickFileExplorerSaveButton(self): 
        parent_element = self.driver.find_element_by_class_name("WordPadClass")
        saveas_element = parent_element.find_element_by_name("Save As")
        saveas_element.find_element_by_name("Save").click()
            
    def isWordpadProcessExist(self):
        return self.utils.isProcessExist(self.__process)
     
    def clickCloseButtonOfWordpadApp(self):
        parent_element = self.driver.find_element_by_class_name("WordPadClass")
        titlebar_element = parent_element.find_element_by_id("TitleBar")
        titlebar_element.find_element_by_name("Close").click()

    def clickSaveButtonOfWordpadPopUp(self):
        parent_element = self.driver.find_element_by_class_name("WordPadClass")
        wordpad_element = parent_element.find_element_by_class_name("#32770")
        pane_element = wordpad_element.find_element_by_name("WordPad")
        pane_element.find_element_by_name("Save").click()

    def connectWordpadToWebDriver(self):
        self.driver = webdriver.Remote(command_executor="http://localhost:9999",
                                       desired_capabilities={
                                           'app': r'C:\Program Files (x86)\Windows NT\Accessories\wordpad.exe',
                                           'debugConnectToRunningApp': 'true',
                                           "args": "-port 345"
                                        })
        

    def isWordpadElementExist(self, webelement, **property1):
        return self.utils.isElementPresent( webelement,  **property1)

    def clickDontSaveWordpadPopUp(self):
        parent_element = self.driver.find_element_by_class_name("WordPadClass")
        wordpad_element = parent_element.find_element_by_class_name("#32770")
        pane_element = wordpad_element.find_element_by_name("WordPad")
        pane_element.find_element_by_name("Don't Save").click()
        
    def getWordpadDocumentContentText(self):
        doc_element = self.driver.find_element_by_class_name("WordPadClass")
        content_element = doc_element.find_element_by_name("Rich Text Window")
        return content_element.text
        
    def clickCloseWordpadAppButton(self):
        parent_element = self.driver.find_element_by_class_name("WordPadClass")
        titlebar_element = parent_element.find_element_by_id("TitleBar")
        titlebar_element.find_element_by_name("Close").click()

    def clickDontSaveButtonInWordpadPopUp(self):
        parent_element = self.driver.find_element_by_class_name("WordPadClass")
        wordpad_element = parent_element.find_element_by_class_name("#32770")
        pane_element = wordpad_element.find_element_by_name("WordPad")
        pane_element.find_element_by_name("Don't Save").click()     

    def isWordpadElementTextExist(self, webelement, text , **property1):
        return self.utils.isElementTextPresent(webelement, text , **property1)

    def getWordpadPopUpWindowElement(self):
        parent_element = self.driver.find_element_by_class_name("WordPadClass")
        open_dialog_element = parent_element.find_element_by_class_name("#32770")
        return open_dialog_element
    
    def clickOKOnWordpadOpenDialog(self):
        open_dialog_element = self.getWordpadPopUpWindowElement()
        open_dialog_element.find_element_by_name("OK").click()
    
    def clickCancelOnWordpadOpenDialog(self):
        opendialog_element = self.driver.find_element_by_class_name("#32770")
        elements =  opendialog_element.find_elements_by_class_name("Button")
        for element in elements:
            if element.get_attribute("Name") == "Cancel":
                element.click()
                   
      