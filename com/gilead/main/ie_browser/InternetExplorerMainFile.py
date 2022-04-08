'''
Created on Oct 4, 2018

@author: pparamasivan
'''

from definitions import IE_BROWSER_PROCESS, IE_DRIVER, ROOT_DIR, IE_IMAGES_DIR
from selenium import webdriver
#from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from com.gilead.main.browser.BrowserMainFile import BrowserParentClass
from selenium.webdriver.ie.options import Options
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import pyautogui as pg


class InternetExplorerMainClass(BrowserParentClass):
    __process = IE_BROWSER_PROCESS
    __ie_driver = IE_DRIVER
    __root_dir = ROOT_DIR
    __images_dir = ROOT_DIR+IE_IMAGES_DIR
    
    def __init__(self, utilsRef, loggerRef):
        self.utils = utilsRef
        self.logger = loggerRef
        self.driver = ""
    
    def killProcess(self):
        self.utils.killProcess(self.__process)

    def openApplication(self):
        self.utils.openApplication(self.__process)

    def startInternetExplorerBrowser(self):
        ''' opts = Options()
        opts.ignore_protected_mode_settings = True
        opts.ignore_zoom_level = True
        opts.require_window_focus = True '''
#         opts.set_capability(InternetExplorerDriver.INITIAL_BROWSER_URL, valu)
        cap = DesiredCapabilities().INTERNETEXPLORER
        cap['platform'] = "windows"
        cap['version'] = "11"
        cap['browserName'] = "internet explorer"
        cap['ignoreProtectedModeSettings'] = True
        cap['IntroduceInstabilityByIgnoringProtectedModeSettings'] = True
        cap['nativeEvents'] = True
        cap['ignoreZoomSetting'] = True
        cap['requireWindowFocus'] = True
        cap['INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS'] = True
#         cap['INITIAL_BROWSER_URL'] = ""
        self.driver = webdriver.Ie(capabilities=cap, executable_path=self.__root_dir+self.__ie_driver)
        

    def getTextOfAvatierPageTitle(self, label):
        return self.driver.find_element_by_id(label).text
    
    def closeInternetExplorerBrowser(self):
        self.closeBrowser()
    
    def isInternetExplorerDriverProcessExist(self):
        return self.utils.isProcessExist(self.__process)

    def isImageVisible(self, scr1): #, scr2, scrollDown=False, state=1):
        a = ""
        b = ""
        try:
            a, b, _, _ = pg.locateOnScreen(scr1, 10)
            return True
            self.logger.info("image {} is located at ( {} , {} )".format(scr1, a, b))
        except:
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
        pg.click(x=a+offset_x, y=b+offset_y, clicks=1, duration=2)
        self.logger.info("Clicked on image {} at location ( {}, {} )".format(src1, a+offset_x, b+offset_y))

    def clickNextbutton(self):
        self.driver.find_element_by_name("About Internet Explorer").click

    def check_if_ImageExist(self,state):  
        try:
            if state == 1:
                a, b, _, _ = pg.locateOnScreen(self.__images_dir + '//version_check.PNG', 10)
                self.logger.info("Able to locate image {}/{} at location ( {}, {} )".format(self.__images_dir, '//version_check.PNG', a, b) )
                self.utils.sleepUntil(2)
                return True
            elif state == 2:
                a, b, _, _ = pg.locateOnScreen(self.__images_dir + '//IEdefault_check.PNG', 10)
                self.logger.info("Able to locate image {}/{} at location ( {}, {} )".format(self.__images_dir, '//IEdefault_check.PNG', a, b) )
                self.utils.sleepUntil(2)
                return True
            elif state == 3:
                a, b, _, _ = pg.locateOnScreen(self.__images_dir + '//mediumhigh_check.PNG', 10)
                self.logger.info("Able to locate image {}/{} at location ( {}, {} )".format(self.__images_dir, '//mediumhigh_check.PNG', a, b) )
                self.utils.sleepUntil(2)
                return True
            elif state == 4:
                a, b, _, _ = pg.locateOnScreen(self.__images_dir + '//trustedsites_check.PNG', 10)
                self.logger.info("Able to locate image {}/{} at location ( {}, {} )".format(self.__images_dir, '//trustedsites_check.PNG', a, b) )
                self.utils.sleepUntil(2)
                return True 
            elif state == 5:
                a, b, _, _ = pg.locateOnScreen(self.__images_dir + '//restrictedsites_check.PNG', 10)
                self.logger.info("Able to locate image {}/{} at location ( {}, {} )".format(self.__images_dir, '//restrictedsites_check.PNG', a, b) )
                self.utils.sleepUntil(2)
                return True 
            elif state == 6:
                a, b, _, _ = pg.locateOnScreen(self.__images_dir + '//connections_check.PNG', 10)
                self.logger.info("Able to locate image {}/{} at location ( {}, {} )".format(self.__images_dir, '//connections_check.PNG', a, b) )
                self.utils.sleepUntil(2)
                return True
            elif state == 7:
                a, b, _, _ = pg.locateOnScreen(self.__images_dir + '//automatically_check.PNG', 10)
                self.logger.info("Able to locate image {}/{} at location ( {}, {} )".format(self.__images_dir, '//automatically_check.PNG', a, b) )
                self.utils.sleepUntil(2)
                return True
            elif state == 8:
                a, b, _, _ = pg.locateOnScreen(self.__images_dir + '//customsecurity_check.PNG', 10)
                self.logger.info("Able to locate image {}/{} at location ( {}, {} )".format(self.__images_dir, '//customsecurity_check.PNG', a, b) )
                self.utils.sleepUntil(2)
                return True
            elif state == 9:
                a, b, _, _ = pg.locateOnScreen(self.__images_dir + '//localintranet_check.PNG', 10)
                self.logger.info("Able to locate image {}/{} at location ( {}, {} )".format(self.__images_dir, '//localintranet_check.PNG', a, b) )
                self.utils.sleepUntil(2)
                return True     
        except ImageNotFoundException as e:
            e.print_stack_trace()
            return False

    # def clickFileTabInOutlookMainWindow(self):
    #     self.driver.find_element_by_name("File Tab").click()

    # def clickFileMenu(self):
    #     self.clickFileTabInOutlookMainWindow()
    #     file_menu_element = self.driver.find_element_by_name("Backstage view").find_element_by_name("Help")
    #     self.utils.sleepUntil(3)

    

        
