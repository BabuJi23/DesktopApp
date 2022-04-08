'''
Created on Sep 26, 2018

@author: bkotamahanti
'''
from definitions import CHROME_BROWSER_PROCESS,  ROOT_DIR, CHROME_DRIVER
from selenium import webdriver
from com.gilead.main.browser.BrowserMainFile import BrowserParentClass

class ChromeBrowserMainClass(BrowserParentClass):
    __process = CHROME_BROWSER_PROCESS
    __chrome_driver = CHROME_DRIVER
    __root_dir = ROOT_DIR
    
    def __init__(self, utilsRef, loggerRef):
        self.utils = utilsRef
        self.logger = loggerRef
        self.driver = ""
    
    def killProcess(self):
        self.utils.killProcess(self.__process)
        
    def setChromeOptions(self):
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument("--no-proxy-server")
        chrome_options.add_argument("--disable-extensions")
        return chrome_options
    
    def startChromeBrowser(self):
        chrome_options = self.setChromeOptions()
        self.driver =  webdriver.Chrome(executable_path=self.__root_dir+self.__chrome_driver, chrome_options=chrome_options)
        self.driver.maximize_window()

    def getTextOfAvatierPageTitle(self, id):
        return self.driver.find_element_by_id(id).text
    
    def closeChromeBrowser(self):
        self.closeBrowser()
    
    
    def isChromeDriverProcessExist(self):
        return self.utils.isProcessExist(self.__process)