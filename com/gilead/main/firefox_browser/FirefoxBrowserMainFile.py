'''
Created on Oct 2, 2018

@author: bkotamahanti
'''
from com.gilead.main.browser.BrowserMainFile import BrowserParentClass
from definitions import FIREFOX_BROWSER_PROCESS, FIREFOX_DRIVER, ROOT_DIR
from selenium import webdriver

class FirefoxBrowserMainClass(BrowserParentClass):
    __process = FIREFOX_BROWSER_PROCESS
    __firefox_driver = FIREFOX_DRIVER
    __root_dir = ROOT_DIR
    
    def __init__(self, utilsRef, loggerRef):
        self.utils = utilsRef
        self.logger = loggerRef
        self.driver = ""
        
    
    def killProcess(self):
        self.utils.killProcess(self.__process)  
    
    def startFirefoxBrowser(self):
        self.driver = webdriver.Firefox(executable_path=self.__root_dir+self.__firefox_driver)
#         self.driver.maximize_window()
    
    def closeFirefoxBrowser(self):
        self.closeBrowser()
    
    def isFirefoxDriverProcessExist(self):
        return self.utils.isProcessExist(self.__process)
