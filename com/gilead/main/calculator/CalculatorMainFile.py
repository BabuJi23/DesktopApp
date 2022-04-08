'''
Created on Aug 2, 2018

@author: bkotamahanti
'''
from definitions import CALCULATOR_PROCESS_NAME 
from selenium import webdriver
from definitions import cal_dict

class CalculatorMainClass(object):
    minSearchTime = 2
    
    def __init__(self, utilsRef, loggerRef):
        self.utils = utilsRef
        self.logger = loggerRef
        self.driver = ""
    
    ''' Initializing webdriver '''
    def connectCalculatorToWebDriver(self):
        self.driver = webdriver.Remote(command_executor='http://localhost:9999',
                                       desired_capabilities={
                                               "debugConnectToRunningApp": "true",
                                               "app" :  r"C:\\Windows\\System32\\calc.exe" }
                                               )
        self.logger.info("Connected already running calculator process to winium Driver instead of opening new one.")
    
    def startCalculator_via_StartMenu(self):
        return self.utils.startProgramByName_via_StartMenu("Calculator")
    
    def isCalculatorProcessExist(self):
        return self.utils.isProcessExist(CALCULATOR_PROCESS_NAME)   
    
    def openCalculator(self, filePath):
        self.utils.openApplication(filePath) 
    
    def killProcess(self):
        self.utils.killProcess("Calculator.exe")
    
    def addGivenNumbers(self, num1, num2):
        self.driver.find_element_by_id(cal_dict[num1]['AutomationId']).click()
        self.logger.info("clicked given number1: "+str(num1))
        self.driver.find_element_by_id('plusButton').click()
        self.logger.info("clicked + sign ")
        self.driver.find_element_by_id(cal_dict[num2]['AutomationId']).click()
        self.logger.info("clicked given number2: "+str(num2))
        self.driver.find_element_by_id("equalButton").click()
        self.logger.info("---Done---")  
    
    def substractGivenNumbers(self,num1, num2):
        self.driver.find_element_by_id(cal_dict[num1]['AutomationId']).click()
        self.logger.info("clicked given number1: "+str(num1))
        self.driver.find_element_by_id('minusButton').click()
        self.logger.info("clicked - sign ")
        self.driver.find_element_by_id(cal_dict[num2]['AutomationId']).click()
        self.logger.info("clicked given number2: "+str(num2))
        self.driver.find_element_by_id("equalButton").click()
        self.logger.info("---Done---")
    
    def multiplyGivenNumbers(self, num1, num2):      
        self.driver.find_element_by_id(cal_dict[num1]['AutomationId']).click()
        self.logger.info("clicked given number1: "+str(num1))
        self.driver.find_element_by_id('multiplyButton').click()
        self.logger.info("clicked * sign ")
        self.driver.find_element_by_id(cal_dict[num2]['AutomationId']).click()
        self.logger.info("clicked given number2: "+str(num2))
        self.driver.find_element_by_id("equalButton").click()
        self.logger.info("---Done---")
    
    def divideGivenNumbers(self, num1, num2):
        self.driver.find_element_by_id(cal_dict[num1]['AutomationId']).click()
        self.logger.info("clicked given number1: "+str(num1))
        self.driver.find_element_by_id('divideButton').click()
        self.logger.info("clicked / sign ")
        self.driver.find_element_by_id(cal_dict[num2]['AutomationId']).click()
        self.logger.info("clicked given number2: "+str(num2))
        self.driver.find_element_by_id("equalButton").click()
        self.logger.info("---Done---")
    
    def getCalculatorResults(self): 
        self.logger.info("Getting the CalculatorResults")
        return self.driver.find_element_by_id('CalculatorResults').get_attribute("Name")