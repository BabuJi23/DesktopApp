'''
Created on Oct 2, 2018

@author: bkotamahanti
'''

from definitions import EDGE_BROWSER_PROCESS, EDGE_DRIVER_15063, EDGE_DRIVER_16999, ROOT_DIR
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from com.gilead.main.browser.BrowserMainFile import BrowserParentClass

class EdgeBrowserMainClass(BrowserParentClass):
    __process = EDGE_BROWSER_PROCESS
    __edge_driver_16999 = EDGE_DRIVER_16999
    __edge_driver_15063 = EDGE_DRIVER_15063
    
    __root_dir = ROOT_DIR
    
    def __init__(self, utilsRef, loggerRef):
        self.utils = utilsRef
        self.logger = loggerRef
        self.driver = ""
    
    def killProcess(self):
        self.utils.killProcess(self.__process)
    
    def startEdgeBrowser(self):
        if(self.utils.getOSBuild()=='16299'):
            self.driver = webdriver.Edge(executable_path=self.__root_dir+self.__edge_driver_16999)
        elif(self.utils.getOSBuild()=='15063'):
            self.driver = webdriver.Edge(executable_path=self.__root_dir+self.__edge_driver_15063)
        
    def getTextOfAvatierPageTitle(self, id):
        return self.driver.find_element_by_id(id).text
    
    def getDisagreeButtonId(self, id):
        self.driver.find_element_by_id(id).click()
 
    def getAgreeButtonId(self, id):
        self.driver.find_element_by_id(id).click()
       
    def navigateBack(self):
        self.driver.back()

    def getDisagreeValidationMessage(self, id):
        return self.driver.find_element_by_id(id).text

    def getDomainValidationMessage(self, id):
        return self.driver.find_element_by_id(id).text

    def closeEdgeBrowser(self):
        self.closeBrowser()
    
    def isEdgeDriverProcessExist(self):
        return self.utils.isProcessExist(self.__process)
