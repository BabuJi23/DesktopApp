'''
Created on April 25, 2019

@author: ngarimella
'''
from definitions import DEVICE_MANAGER_PROCESS, SYSTEM_IMAGES_DIR, ROOT_DIR
import pyautogui as pg
from pyscreeze import ImageNotFoundException
from selenium import webdriver

class SystemsettingsMainClass():

    __images_dir = ROOT_DIR+SYSTEM_IMAGES_DIR
    
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
    
    def isElementExist(self, parent_element, **element_attribute_text):
        return self.utils.isElementPresent(parent_element, **element_attribute_text)
    

    def isImageVisible(self, scr1): #, scr2, scrollDown=False, state=1):
        a = ""
        b = ""
        try:
            a, b, _, _ = pg.locateOnScreen(scr1, 10)
            return True
            self.logger.info("image {} is located at ( {} , {} )".format(scr1, a, b))
        except:
#             self.close_window()
            self.logger.debug("image {} is  not located".format(scr1))
            return False
    
    
    def doubleClickOnImage(self, src1, offset_x, offset_y,scrollDown=False):
        if scrollDown:
            pg.press('pagedown')
            pg.press('pagedown')
            
        a = ""
        b = ""
        a, b, _, _ = pg.locateOnScreen(src1, 10)
        self.utils.sleepUntil(5)
#         pg.moveTo(a+40, b+8, duration=2)
        pg.click(x=a+offset_x, y=b+offset_y, clicks=1, duration=2)
        self.logger.info("Clicked on image {} at location ( {}, {} )".format(src1, a+offset_x, b+offset_y))

    def isWorkingProperlyImageVisisble(self, state):  
        try:
            if state == 1:
                a, b, _, _ = pg.locateOnScreen(self.__images_dir + '//windows_checking.PNG', 10)
                self.logger.info("Able to locate image {}/{} at location ( {}, {} )".format(self.__images_dir, '//windows_checking.PNG', a, b) )
                self.utils.sleepUntil(2)
                return True
            elif state == 2:
                _, _, _, _ = pg.locateOnScreen(self.__images_dir + 'allow_remoteconnections_computer.PNG', 10)
                self.utils.sleepUntil(2)
                return True
            elif state == 3:
                _, _, _, _ = pg.locateOnScreen(self.__images_dir + 'allow_connections_enabled_check.PNG', 10)
                self.utils.sleepUntil(2)
                return True
            elif state == 4:
                _, _, _, _ = pg.locateOnScreen(self.__images_dir + 'remote_users_existing_check.PNG', 10)
                self.utils.sleepUntil(2)
                return True
            elif state == 5:
                _, _, _, _ = pg.locateOnScreen(self.__images_dir + 'sytemlog_check.PNG', 10)
                self.utils.sleepUntil(2)
                return True
            elif state == 6:
                _, _, _, _ = pg.locateOnScreen(self.__images_dir + 'automatically_restart_check.PNG', 10)
                self.utils.sleepUntil(2)
                return True
            elif state == 7:
                _, _, _, _ = pg.locateOnScreen(self.__images_dir + 'drive_protection_check.PNG', 10)
                self.utils.sleepUntil(2)
                return True
            elif state == 8:
                _, _, _, _ = pg.locateOnScreen(self.__images_dir + 'screen_power.PNG', 10)
                self.utils.sleepUntil(2)
                return True
            elif state == 9:
                _, _, _, _ = pg.locateOnScreen(self.__images_dir + 'screen_plugin.PNG', 10)
                self.utils.sleepUntil(2)
                return True
            elif state == 10:
                _, _, _, _ = pg.locateOnScreen(self.__images_dir + 'sleep_power.PNG', 10)
                self.utils.sleepUntil(2)
                return True
            elif state == 11:
                _, _, _, _ = pg.locateOnScreen(self.__images_dir + 'sleep_plugin.PNG', 10)
                self.utils.sleepUntil(2)
                return True
            elif state == 12:
                _, _, _, _ = pg.locateOnScreen(self.__images_dir + 'timezone_check.PNG', 10)
                self.utils.sleepUntil(2)
                return True
        except ImageNotFoundException as e:
            e.print_stack_trace()
            return False  
        return False


    def close_window(self):
        pg.hotkey('alt', 'f4')
        self.logger.info("Closed the window with alt+f4 key")
     
        
        