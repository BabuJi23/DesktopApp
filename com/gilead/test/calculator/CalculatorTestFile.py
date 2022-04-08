'''
Created on Sep 12, 2018

@author: bkotamahanti
'''
from com.gilead.main.calculator.CalculatorMainFile import CalculatorMainClass
from definitions import CALCULATOR_PROCESS_EXECUTABLE
import inspect

class CalculatorTestClass():
    
    __cal_exe = CALCULATOR_PROCESS_EXECUTABLE
    
    def __init__(self, utilsRef, loggerRef):
        self.logger = loggerRef
        self.utils = utilsRef
        self.calculator_maincode = CalculatorMainClass(utilsRef, loggerRef)
    
    '''TC:  W10DA-4 '''
    def calculator_open_perform_arithemetic_operations_close_app_testcase_id_W10DA_4(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the WebDriver of calculator app ")
            
            self.logger.info("Starting execution of function :"+str(func.co_name))
            #TC Step 1
            self.logger.info("Checking and killing the process Calculator.exe before executing the test case")
            self.calculator_maincode.killProcess() 
            #TC Step 2 
            ''' have some weird issues when we start calculator via Start menu. It sometimes opening Control Panel than Calculator.
             hence commenting it out. '''
#             self.calculator_maincode.startCalculator_via_StartMenu()
            self.calculator_maincode.openCalculator(self.__cal_exe)
            self.utils.sleepUntil(4)
            assert self.calculator_maincode.isCalculatorProcessExist(), "Calculator process Calculator.exe didn't exist "
            
            ''' Connect the already running calculator process to winium Driver instead of opening new one.'''
            self.calculator_maincode.connectCalculatorToWebDriver()
            
            self.__num1 = 3
            self.__num2 = 4
            #TC Step 3
            self.calculator_maincode.addGivenNumbers(3,4)
            '''waiting to start next operation'''
            self.utils.sleepUntil(1)
            assert '7' in self.calculator_maincode.getCalculatorResults(), "CalculatorResults:"+self.calculator_maincode.getCalculatorResults()+" didn't match expected result:"+str(self.__num1)+"+"+str(self.__num2)+"="+(str(self.__num1+self.__num2))
            self.logger.info("Addition performed successfully")
            #TC Step 4
            self.__num1 = 7
            self.__num2 = 3
            self.calculator_maincode.substractGivenNumbers(7,3)
            '''waiting to start next operation'''
            self.utils.sleepUntil(1)
            assert '4' in self.calculator_maincode.getCalculatorResults(), "CalculatorResults:"+self.calculator_maincode.getCalculatorResults()+" didn't match expected result:"+str(self.__num1)+"-"+str(self.__num2)+"="+(str(self.__num1-self.__num2))
            self.logger.info("Subtraction performed successfully")
            
            #TC Step 5
            self.__num1 = 5
            self.__num2 = 3
            self.calculator_maincode.multiplyGivenNumbers(5,3)
            '''waiting to start next operation'''
            self.utils.sleepUntil(1)
            assert '15' in self.calculator_maincode.getCalculatorResults(), "CalculatorResults:"+self.calculator_maincode.getCalculatorResults()+" didn't match expected result:"+str(self.__num1)+"*"+str(self.__num2)+"="+(str(self.__num1*self.__num2))
            self.logger.info("Multiplication performed successfully")
            
            #TC Step 6
            self.__num1 = 6
            self.__num2 = 2
            self.calculator_maincode.divideGivenNumbers(6,2)
            assert '3' in self.calculator_maincode.getCalculatorResults(), "CalculatorResults:"+self.calculator_maincode.getCalculatorResults()+" didn't match expected result:"+str(self.__num1)+"/"+str(self.__num2)+"="+(str(self.__num1/self.__num2))
            self.logger.info("Division performed successfully")
            self.calculator_maincode.killProcess()
            assert (self.calculator_maincode.isCalculatorProcessExist()==False), "Calculator process Calculator.exe didn't exist anymore "
            self.logger.info("Calculator process Calculator.exe didn't exist anymore ...")
            return True, "", func.co_name
        
        except (AssertionError, AttributeError, TypeError) as e:
            return False, e, func.co_name
        
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the webDriver process of calculator app")
            self.calculator_maincode.killProcess()