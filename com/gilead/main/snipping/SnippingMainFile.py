'''
Created on Sept 19, 2018

@author: pparamasivan
'''
from selenium import webdriver
import pyautogui
from definitions import  SNIP_TOOL_PROCESS


class SnippingMainClass(object):
    __process = SNIP_TOOL_PROCESS


    def __init__(self, utilsRef, loggerRef):
        self.utils = utilsRef
        self.logger = loggerRef
        self.driver = ""

    def killProcess(self):
        self.utils.killProcess(self.__process)


    ''' Initializing webdriver - by connecting already running Snipping tool application'''

    def connectSnippingToWebDriver(self):
        self.driver = webdriver.Remote(command_executor='http://localhost:9999',
                                       desired_capabilities={
                                           "debugConnectToRunningApp": "true",
                                           "app":  r"C:\\Windows\\System32\\SnippingTool.exe",
                                           "args": "-port 345"
                                       }
                                       )
        self.logger.info(
            "Connected the already running SnippingTool process to winium Driver instead of opening new one.")

    def startSnipping_via_StartMenu(self):
        self.utils.startProgramByName_via_StartMenu("Snipping Tool")

    def isSnippingProcessExist(self):
        return self.utils.isProcessExist(self.__process)

    def clickSnippingToolCloseButton(self):
        parent_element = self.driver.find_element_by_name("Snipping Tool")
        parent_element.find_element_by_name("Close").click()
        

    def isCancelButtonEnabled(self):
        parent_element = self.driver.find_element_by_name("Snipping Tool")
        cancel_element = parent_element.find_element_by_name("Cancel")

        if cancel_element.is_enabled():
            self.logger.info(" Cancel button is enabled")
            return True
        else:
            self.logger.info(" Cancel button is disabled")
            return False 
        
    def clickSnippingToolNewButton(self):
        parent_element = self.driver.find_element_by_name("Snipping Tool")
        toolbar_element = parent_element.find_element_by_class_name('ToolbarWindow32')
        toolbar_element.find_element_by_name("New").click()

    def clickSnippingToolOptionButton(self):
        parent_element = self.driver.find_element_by_name("Snipping Tool")
        toolbar_element = parent_element.find_element_by_class_name('ToolbarWindow32')
        toolbar_element.find_element_by_name("Options").click()
        
    def clickSnippingToolOptionOkButton(self):
        Option_dialog = self.driver.find_element_by_name("Snipping Tool Options")
        Option_dialog.find_element_by_name("OK").click()
