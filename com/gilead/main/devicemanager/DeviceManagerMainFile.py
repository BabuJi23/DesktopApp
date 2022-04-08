'''
Created on Nov 6, 2018

@author: bkotamahanti
'''

from definitions import DEVICE_MANAGER_PROCESS, DEVICE_MANAGER_IMAGES_DIR, ROOT_DIR
import pyautogui as pg
from pyscreeze import ImageNotFoundException

class DeviceManagerMainClass():
    __process = DEVICE_MANAGER_PROCESS
    __images_dir = ROOT_DIR+DEVICE_MANAGER_IMAGES_DIR
    
    def __init__(self, utilsRef, loggerRef):
        self.utils = utilsRef
        self.logger = loggerRef
        self.driver = ""
        self.utils.killProcess(self.__process)
    
    
    def killProcess(self):
        self.utils.killProcess(self.__process)
    
    
    def maximize_window(self):
        pg.hotkey('win', 'up')
    
    
    def close_window(self):
        pg.hotkey('alt', 'f4')
        self.logger.info("Closed the window with alt+f4 key")
    
    
    def openApplication(self, fileName):
        self.utils.openApplication(fileName)
        
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
        self.utils.sleepUntil(3)
#         pg.moveTo(a+40, b+8, duration=2)
        pg.click(x=a+offset_x, y=b+offset_y, clicks=2, duration=2)
        self.logger.info("Clicked on image {} at location ( {}, {} )".format(src1, a+offset_x, b+offset_y))
     
    def isWorkingProperlyImageVisisble(self, state=1):  
        try:
            if state == 1:
                a, b, _, _ = pg.locateOnScreen(self.__images_dir + '//generic_working_properly.PNG', 10)
                self.logger.info("Able to locate image {}/{} at location ( {}, {} )".format(self.__images_dir, '//generic_working_properly.PNG', a, b) )
                self.utils.sleepUntil(2)
                return True
            elif state == 2:
                _, _, _, _ = pg.locateOnScreen(self.__images_dir + 'no_device_driver.PNG', 10)
                self.utils.sleepUntil(2)
                return True
            if state == 3:
                _, _, _, _ = pg.locateOnScreen(self.__images_dir + 'disabled.PNG', 10)
                self.utils.sleepUntil(2)
                return True
        except ImageNotFoundException as e:
            e.print_stack_trace()
            return False
            
        return False
        
#         if scrollDown:
#             pg.press('pagedown')
#             pg.press('pagedown')
#         a, b, _, _ = pg.locateOnScreen(scr2, 20)
#         pg.moveTo(a, b)
#         pg.click(clicks=2)
#         time.sleep(1)
#        
#     
#     
