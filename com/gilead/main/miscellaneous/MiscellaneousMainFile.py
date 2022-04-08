'''
Created on Oct 11, 2018

@author: bkotamahanti
'''
import platform
import re
from selenium import webdriver
import pyautogui as pg


class MiscellaneousMainClass():
    
    def __init__(self, utilsRef, loggerRef):
        self.logger = loggerRef
        self.utils = utilsRef
        self.driver =""
    
    def killProcess(self, process):
        self.utils.killProcess(process)
        
    def isGivenProcessExist(self, process):
        return self.utils.isProcessExist(process)
    
    def startProgramByName_via_StartMenu(self, searchText):
        self.utils.startProgramByName_via_StartMenu(searchText)
    
    def connectPerformanceMonitorToWebDriver(self):
        self.driver = webdriver.Remote(command_executor="http://localhost:9999",
                                       desired_capabilities={
                                           'app': r'C:\Windows\System32\mmc.exe',
                                           'debugConnectToRunningApp': 'true',
                                           "args": "-port 345"
                                        })
    
    
    def isOS64Bit(self):
        return platform.machine().endswith('64')
    
    
    def getWindowsPlatform(self):
        return platform.platform()
    
    
    def isSystemHostNameContainsGivenText(self, expText):
        fqdn = self.utils.getSystemHostName()
        self.logger.info("Computer name is:"+fqdn)
        if re.findall(expText, fqdn):
            return True
        return False
    
    
    
    def isElementExist(self, parent_element, **element_attribute_text):
        return self.utils.isElementPresent(parent_element, **element_attribute_text)
    
    def clickPerformanceMonitorOptions(self):
        parent_window_element = self.driver.find_element_by_name("Performance Monitor")
        parent_window_element.find_element_by_name("File").click()
        self.utils.sleepUntil(3)
        self.driver.find_element_by_name("Options...").click()
    
    def getPerformanceMonitorOptionsDialogElement(self):
        parent_window_element = self.driver.find_element_by_name("Performance Monitor")
        options_element = parent_window_element.find_element_by_name("Options")
        return options_element
    
    def closePerformanceMonitor(self):
        parent_window_element = self.driver.find_element_by_name("Performance Monitor")
        options_element = parent_window_element.find_element_by_name("Options")
        options_element.find_element_by_name("OK").click()
        self.utils.sleepUntil(2)
        parent_window_element = self.driver.find_element_by_name("Performance Monitor")
        parent_window_element.find_element_by_name("File").click()
        self.utils.sleepUntil(1)
        self.driver.find_element_by_name("Exit").click()
    
    def isImageVisible(self, scr1): #, scr2, scrollDown=False, state=1):
        a = ""
        b = ""
        try:
            a, b, _, _ = pg.locateOnScreen(scr1, 10)
            self.logger.info("image {} is located at ( {} , {} )".format(scr1, a, b))
            return True
        except:
#             self.close_window()
            self.logger.debug("image {} is  not located".format(scr1))
            return False
        
        