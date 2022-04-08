'''
Created on Sep 19, 2018

@author: bkotamahanti

'''
from definitions import ACROBAT_READER_PROCESS, ACROBAT_READER_MAIN_DIR, ROOT_DIR
from selenium import webdriver
import pyautogui
from selenium.common.exceptions import WebDriverException

class AcrobatReaderMainClass():
    __process = ACROBAT_READER_PROCESS
    __root_dir = ROOT_DIR
    __acrobat_main_dir = ACROBAT_READER_MAIN_DIR
    
    def __init__(self, utilRef, loggerRef):
        self.utils = utilRef
        self.logger = loggerRef
        self.driver = ""
    
    def killProcess(self):
        self.utils.killProcess(self.__process)
    
    def startAcrobatReader_via_StartMenu(self):
        self.utils.startProgramByName_via_StartMenu("Acrobat Reader")
    
    def isAcrobatReaderProcessExist(self):
        return self.utils.isProcessExist(self.__process)
    
    def connectAcrobatReaderToWebDriver(self):
        self.driver = webdriver.Remote(command_executor="http://localhost:9999",
                                       desired_capabilities={ "app" : r'C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe',
                                                             "debugConnectToRunningApp": 'true',
                                                             "args": "-port 345"
                                                             }
                                       )
    
    
    def getAcrobatReaderMainWindowMenuElement(self):
        acrobat_reader_main_window_element = self.driver.find_element_by_class_name("AcrobatSDIWindow") #Name = Adobe Acrobat Reader DC
        menu_bar_element = acrobat_reader_main_window_element.find_element_by_name("Application")
        return menu_bar_element
    
    
    def getAcrobatReaderMainWindowElement(self):
        acrobat_reader_main_window_element = self.driver.find_element_by_class_name("AcrobatSDIWindow")
        return acrobat_reader_main_window_element
    
    def isAcrobatReaderElementExist(self, webelement, **property1):
        return self.utils.isElementPresent( webelement,  **property1)
    
    def clickPrintwindowCancelButton(self):
        acrobat_reader_main_window_element = self.driver.find_element_by_class_name("AcrobatSDIWindow")
        print_window_element  = acrobat_reader_main_window_element.find_element_by_name("Print")
        print_window_element.find_element_by_name("Cancel").click()
    
    def clickAcrobatReaderOpenMenuOption(self):
        pyautogui.hotkey('ctrl','o')
    
    def clickAcrobatReaderClose(self):
        pyautogui.hotkey('ctrl','q')

    def enterFileNameAndClickToOpen(self, filePath): 
        ''' in this method we use combination of winium and pyautogui tools '''
        open_window_element = self.driver.find_element_by_class_name("#32770")
        file_edit_text_element = open_window_element.find_element_by_class_name("Edit")
        file_edit_text_element.click()
        ''' using pyauto gui to enter fileName as winium is stopping abruptly while trying to enter filename '''
        #pyautogui.typewrite(self.__root_dir+self.__acrobat_main_dir+fileName)
        pyautogui.typewrite(filePath)
        #self.logger.info("Entered Filename as :"+self.__root_dir+self.__acrobat_main_dir+fileName)
        self.logger.info("Entered Filename as :" +filePath)

        '''finding Open button and clicking it '''
        button_elements = self.driver.find_element_by_class_name("#32770").find_elements_by_class_name("Button")
        for button_element in button_elements:
            if button_element.get_attribute("Name") == "Open" :
                button_element.click()
                self.logger.info("Clicked Open Window Open button")
                break
    
    def clickShareFileAttachToEmailMenuOption(self):
        main_window = self.driver.find_element_by_name("GIAT_test.pdf - Adobe Acrobat Reader DC")
        menu_bar = main_window.find_element_by_name("Application")
        menu_bar.find_element_by_name("File").click()
        
        self.utils.sleepUntil(3)
        try:
            main_window.find_element_by_name("File")
            main_window.find_element_by_name("Share File").click()
        except WebDriverException as e:
            self.logger.info("Not able to click the Share File as it is disabled {}".format(str(e)))
            self.logger.info("It happens if the process AcroRd32.exe is not closed properly")
            pass
        
        self.utils.sleepUntil(3)
        ''' Keeping the below code if there is different version of acrobat reader present then this below code is used '''  
#         open_menu = main_window.find_element_by_class_name("#32768")
#         open_menu.find_element_by_name("Send File").click()
#         open_menu.find_element_by_name("Send File").find_element_by_name("Attach to Email...").click()
#     
    def clickSendEmailDialogContinueButton(self, image1, image2):
        a,b,_,_ = pyautogui.locateOnScreen(image1)
#         print(a,b)# 1072, 425
        pyautogui.click(x=a+40,y=b+8, clicks=1)
        self.utils.sleepUntil(3)
        a,b,_,_ = pyautogui.locateOnScreen(image2)
#         print(a,b)# 1072, 425
        pyautogui.click(x=a+40,y=b+8, clicks=1)
        
        ''' Keeping the below code if there is different version of acrobat reader present then this below code is used ''' 
#         send_dialog = self.driver.find_element_by_name("Send Email")
#         grp_box_element = send_dialog.find_element_by_class_name("GroupBox")
#         grp_box_element.find_element_by_name("Continue").click()

    def clickAcrobatReaderPrintMenuOption(self):
        pyautogui.hotkey('ctrl','p')                