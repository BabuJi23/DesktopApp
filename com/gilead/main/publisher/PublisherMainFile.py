'''
Created on Oct 29, 2018

@author: bkotamahanti
'''
from definitions import PUBLISHER_PROCESS
from selenium import webdriver
import pyautogui

class PublisherMainClass():
    __process = PUBLISHER_PROCESS
    
    def __init__(self, utilsRef, loggerRef):
        self.utils = utilsRef
        self.logger = loggerRef
        self.driver = ""
        self.utils.killProcess(self.__process)
        
    def killProcess(self):
        self.utils.killProcess(self.__process)
        
    def isPublisherProcessExist(self):
        return self.utils.isProcessExist(self.__process)
    
    def connectPublisherToWebDriver(self):
        self.driver = webdriver.Remote(command_executor="http://localhost:9999",
                         desired_capabilities={
#                              "app" :  r"C:\\Program Files (x86)\\Microsoft Office\\Office16\\MSPUB.EXE",
                                    "app" :  r"MSPUB.EXE",
                                     "debugConnectToRunningApp": 'true',
                                               "args": "-port 345"
                             
                             })
    
    def openApplication(self, fileName):
        self.utils.openApplication(fileName)
        
        
    def getPublisherParentElement(self):
        return self.driver.find_element_by_class_name("MSWinPub")
    
    
    def clickOnBlankTemplate(self):
        parent_element = self.getPublisherParentElement()
        backstage_view_element = parent_element.find_element_by_class_name("NetUIFullpageUIWindow").find_element_by_name("Backstage view")
        new_element = backstage_view_element.find_element_by_class_name("NetUISlabContainer")
        elements = new_element.find_elements_by_class_name("NetUIElement")
        for element in elements:
            if element.get_attribute("Name") == "Featured":
                element.find_element_by_name("Blank 8.5 x 11\"").click()
                break
    
    def saveFile(self, fileName):
        pyautogui.hotkey('ctrl','s')
        pyautogui.hotkey('alt', 'f')
        pyautogui.hotkey('alt', 'a')
        pyautogui.hotkey('alt', 'o')
        self.utils.sleepUntil(3)
        browse_window_element = self.driver.find_element_by_class_name("#32770")
        browse_window_element.find_element_by_class_name("DUIViewWndClassName").find_element_by_class_name("Edit").click()
        browse_window_element.find_element_by_class_name("DUIViewWndClassName").find_element_by_class_name("Edit").clear()
        self.utils.sleepUntil(2)
        ''' keeping it for safe side as this is the prefered way to use it. But Winium driver throwing exception for now which is not the case with other MS applications.'''
#         browse_window_element.find_element_by_class_name("DUIViewWndClassName").find_element_by_class_name("Edit").send_keys(fileName)
        pyautogui.typewrite(fileName) 
        self.utils.sleepUntil(3)
        browse_window_element.find_element_by_name("Save").click()
    
    
    def openPublisherFile(self, fileName):
        pyautogui.hotkey('alt', 'f')
        pyautogui.hotkey('alt', 'o')
        pyautogui.hotkey('alt', 'o')
        self.utils.sleepUntil(3)
        browse_window_element = self.driver.find_element_by_class_name("#32770")
        browse_window_element.find_element_by_class_name("ComboBox").find_element_by_class_name("Edit").click()
        browse_window_element.find_element_by_class_name("ComboBox").find_element_by_class_name("Edit").clear()
        self.utils.sleepUntil(2)
        ''' keeping it for safe side as this is the prefered way to use it. But Winium driver throwing exception for now which is not the case with other MS applications.'''
#         browse_window_element.find_element_by_class_name("DUIViewWndClassName").find_element_by_class_name("Edit").send_keys(fileName)
        pyautogui.typewrite(fileName) 
        self.utils.sleepUntil(3)
        browse_window_element.find_element_by_name("Open").click()
    
    
    def isPublisherElementExist(self, webelement, **property1):
        return self.utils.isElementPresent( webelement,  **property1)
    
    def getPublisherPopUpWindowElement(self):
        parent_element = self.getPublisherParentElement()
        open_dialog_element = parent_element.find_element_by_class_name("#32770")
        return open_dialog_element
    
    
    def clickOKOnPublisherOpenDialog(self):
        open_dialog_element = self.getPublisherPopUpWindowElement()
        open_dialog_element.find_element_by_name("OK").click()
        
    def isPublisherElementTextExist(self, webelement, text , **property1):
        return self.utils.isElementTextPresent(webelement, text , **property1)
    
    
    def clickCancelOnPublisherOpenDialog(self):
        browse_window_element = self.driver.find_element_by_class_name("#32770")
        elements =  browse_window_element.find_elements_by_class_name("Button")
        for element in elements:
            if element.get_attribute("Name") == "Cancel":
                element.click()
           
    def enterTextOnBlankTemplate(self, text):
        parent_element = self.getPublisherParentElement()
        parent_element.find_element_by_class_name("MSPPage").click()
        pyautogui.typewrite(text)
        
    def readTextEntered(self): 
        parent_element = self.getPublisherParentElement()
        parent_element.find_element_by_class_name("MSPPage").click()
        return parent_element.find_element_by_class_name("MSPPage").get_attribute("Value.Value")
    
    def clickCloseButtonOfPublisherApp(self):
        parent_element = self.getPublisherParentElement()
        doctop_element = parent_element.find_element_by_name("MsoDockTop")
        doctop_element.find_element_by_name("Close").click()
    
    
    def clickDontSaveInPopUp(self):
        parent_element = self.getPublisherParentElement()
        pane_element = parent_element.find_element_by_name("Microsoft Publisher")
        pane_element.find_element_by_name("Don't Save").click()
        
