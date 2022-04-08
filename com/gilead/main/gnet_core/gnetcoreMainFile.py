'''
Created on June 28, 2019

@author: ngarimella
'''

from definitions import IE_BROWSER_PROCESS, IE_DRIVER, ROOT_DIR, GNET_IMAGES_DIR
from definitions import CHROME_BROWSER_PROCESS,  ROOT_DIR, CHROME_DRIVER
from selenium import webdriver
#from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from com.gilead.main.browser.BrowserMainFile import BrowserParentClass
from selenium.webdriver.ie.options import Options
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import pyautogui as pg
import selenium 
from selenium.webdriver.common.by import By


class gnetcoreMainClass(BrowserParentClass):
    __process = CHROME_BROWSER_PROCESS
    __chrome_driver = CHROME_DRIVER
    __ie_driver = IE_DRIVER
    __root_dir = ROOT_DIR
    __images_dir = ROOT_DIR+GNET_IMAGES_DIR
    
    def __init__(self, utilsRef, loggerRef):
        self.utils = utilsRef
        self.logger = loggerRef
        self.driver = ""
    
    def killProcess(self):
        self.utils.killProcess(self.__process)

    def startProgramByName_via_StartMenu(self, searchText):
        self.utils.startProgramByName_via_StartMenu(searchText) 

    def setChromeOptions(self):
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument("--no-proxy-server")
        chrome_options.add_argument("--disable-extensions")
        return chrome_options

    def startChromeBrowser(self):
        chrome_options = self.setChromeOptions()
        self.driver =  webdriver.Chrome(executable_path=self.__root_dir+self.__chrome_driver, chrome_options=chrome_options)
        self.driver.maximize_window()
        self.driver.get("https://gnethome.gilead.com")

    def OpenIDMwindow(self):
        chrome_options = self.setChromeOptions()
        self.driver =  webdriver.Chrome(executable_path=self.__root_dir+self.__chrome_driver, chrome_options=chrome_options)
        self.driver.maximize_window()
        self.driver.get("https://gnethome.gilead.com/employeeresources/IT/Enterprise_Applications/IDM/Pages/Home.aspx")

    def OpenGVaultwindow(self):
        chrome_options = self.setChromeOptions()
        self.driver =  webdriver.Chrome(executable_path=self.__root_dir+self.__chrome_driver, chrome_options=chrome_options)
        self.driver.maximize_window()
        self.driver.get("https://gvault.veevavault.com/ui/")
        self.utils.sleepUntil(10)

    def openApplication(self):
        self.utils.openApplication(self.__process)

    def getTextOfAvatierPageTitle(self, title):
        return self.driver.find_element_by_id(title).text
    
    def closeInternetExplorerBrowser(self):
        self.closeBrowser()
    
    def isInternetExplorerDriverProcessExist(self):
        return self.utils.isProcessExist(self.__process)   

    def press_esc(self):
        pg.hotkey('Esc')
        self.logger.info("Closed the window with Esc key")

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

    def clickCalendarButton(self):
        pg.press('pagedown')
        parent=self.driver.window_handles[0]
        self.driver.find_element_by_xpath('//*[@title="Calendar"]').click()
        self.utils.sleepUntil(15)
        child1 = self.driver.window_handles[1]
        self.driver.switch_to_window(child1)
        self.driver.find_element_by_link_text('Home').click()
        self.utils.sleepUntil(4)
        self.logger.info("Clicked Advisory Board in Category section")
        self.driver.find_element_by_xpath('//div[@id="divChkCategory"]/div[1]/label').click()
        self.utils.sleepUntil(2)
        self.logger.info("Clicked AML&D in Category section")
        self.driver.find_element_by_xpath('//div[@id="divChkCategory"]/div[7]/label').click()
        self.utils.sleepUntil(2)
        self.logger.info("Clicked Operating Group in Category section")
        self.driver.find_element_by_xpath('//div[@id="divChkCategory"]/div[8]/label').click()
        self.utils.sleepUntil(2)
        self.logger.info("Clicked Asia in Region section")
        self.driver.find_element_by_xpath('//div[@id="divChkRegion"]/div[2]/label').click()
        self.utils.sleepUntil(5)
        self.driver.close()

    def clickGxPLearnButton(self):
        pg.press('pagedown')
        parent=self.driver.window_handles[0]
        self.driver.find_element_by_xpath('//*[@title="GxPLearn"]').click()
        self.utils.sleepUntil(4)
        child1 = self.driver.window_handles[1]
        self.driver.switch_to_window(child1)
        self.driver.find_element_by_id('history').click()
        self.utils.sleepUntil(2)
        self.driver.find_element_by_id('catalog').click()
        self.utils.sleepUntil(2)
        self.driver.find_element_by_id('reports').click()
        self.utils.sleepUntil(2)
        self.driver.find_element_by_id('todo').click()
        self.utils.sleepUntil(2)
        self.driver.find_element_by_id('curriculum').click()
        self.driver.close()

    def clickWorkdayTab(self):
        pg.press('pagedown')
        parent=self.driver.window_handles[0]
        self.driver.find_element_by_xpath('//*[@title="Workday Portal"]').click()
        self.utils.sleepUntil(4)
        child1 = self.driver.window_handles[1]
        self.driver.switch_to_window(child1)
        self.utils.sleepUntil(8)
        self.driver.find_element_by_xpath('//*[@class="workdayHome-x"]/li[4]').click()
        self.utils.sleepUntil(4)
        self.driver.close()
  

    def clickSparcTab(self):
        pg.press('pagedown')
        parent=self.driver.window_handles[0]
        self.driver.find_element_by_xpath('//*[@title="SPARC Application"]').click()
        self.utils.sleepUntil(4)
        child1 = self.driver.window_handles[1]
        self.driver.switch_to_window(child1)
        self.utils.sleepUntil(8)
        self.driver.find_element_by_xpath('//*[@id="sparc_nav_main_top"]/a[1]').click()
        self.utils.sleepUntil(4)
        self.driver.find_element_by_xpath('//div[@class="content_list_wrapper"]/div/div/div/div/div[1]').click()
        self.utils.sleepUntil(4)
        self.driver.find_element_by_xpath('//*[@id="sparc_nav_main_top"]/a[2]').click()
        self.utils.sleepUntil(4)
        self.driver.find_element_by_xpath('//div[@class="content_list_wrapper"]/div/div/div/div/div[1]').click()
        self.utils.sleepUntil(4)
        self.driver.find_element_by_xpath('//*[@id="sparc_nav_main_top"]/a[3]').click()
        self.utils.sleepUntil(4)
        self.driver.close()

    def clickGLearnTab(self):
        pg.press('pagedown')
        parent=self.driver.window_handles[0]
        self.driver.find_element_by_xpath('//*[@title="GLEARN"]').click()
        self.utils.sleepUntil(8)
        child1 = self.driver.window_handles[1]
        self.driver.switch_to_window(child1)
        self.utils.sleepUntil(8)
 
    
    def clickSearchTab(self):
        clear_text = self.driver.find_element_by_xpath("//input[@value='Search People & All GNet']").clear()
        search_person = self.driver.find_element_by_xpath("//input[@value='Search People & All GNet']")
        search_person.send_keys("Kevin White")
        self.utils.sleepUntil(4)
        self.driver.find_element_by_xpath('//*[@class="ms-srch-sb-searchLink"]').click()
        self.utils.sleepUntil(4)
        self.driver.find_element_by_id('NameFieldLink').click()
        self.utils.sleepUntil(8)
        self.driver.close()

    def clickEthicsTab(self):
        pg.press('pagedown')
        parent=self.driver.window_handles[0]
        self.driver.find_element_by_xpath('//*[@title="Ethics & Compliance"]').click()
        self.utils.sleepUntil(4)
        child1 = self.driver.window_handles[1]
        self.driver.switch_to_window(child1)
        self.utils.sleepUntil(10)
        self.driver.find_element_by_xpath('//div[@class="ms-rtestate-field"]/p[1]/a[2]').click()
        self.utils.sleepUntil(4)
        self.driver.close()

    def close_window(self):
        pg.hotkey('alt', 'f4')
        self.logger.info("Closed the window with alt+f4 key")

    def clickIDMTab(self):
        parent=self.driver.window_handles[0]
        self.utils.sleepUntil(4)
        self.driver.find_element_by_xpath('//*[@id="zz10_RootAspMenu"]/li[2]/a/span/span').click()
        self.utils.sleepUntil(4)
        self.driver.find_element_by_xpath('//*[@id="zz10_RootAspMenu"]/li[4]/a/span/span').click()
        self.utils.sleepUntil(4)
        self.driver.find_element_by_xpath('//*[@class="ms-srch-group-content"]/div[1]/table/tbody/tr/td/a').click()
        self.utils.sleepUntil(4)

    def clickGVaultTab(self):
        # import pdb;pdb.set_trace()
        parent=self.driver.window_handles[0]
        self.utils.sleepUntil(15)
        self.driver.find_element_by_xpath('//*[@class="mainTabs"]/a[2]/span').click()
        self.utils.sleepUntil(8)
        self.driver.find_element_by_xpath('//*[@class="mainTabs"]/a[3]/span').click()
        self.utils.sleepUntil(4)
        self.driver.find_element_by_xpath('//*[@class="mainTabs"]/a[4]/span').click()
        self.utils.sleepUntil(4)
        
        
     

       
       
      


    

    

       

       
        
    
    

       

      
        

    

        
