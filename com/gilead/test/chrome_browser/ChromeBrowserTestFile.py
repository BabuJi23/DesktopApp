'''
Created on Sep 26, 2018

@author: bkotamahanti
'''
from com.gilead.main.chrome_browser.ChromeBrowserMainFile import ChromeBrowserMainClass
from selenium.common.exceptions import NoSuchElementException,\
    WebDriverException
from definitions import GILEAD_URL, GOOGLE_URL, TRUVADA_URL,  AVATIER_URL

import inspect


class ChromeBrowserTestClass():
    __gilead_url = GILEAD_URL
    __google_url = GOOGLE_URL
    __truvada_url = TRUVADA_URL
    __avatier_url = AVATIER_URL
        
    def __init__(self, utilsRef, loggerRef):
        self.utils = utilsRef
        self.logger = loggerRef
        self.chrome_maincode = ChromeBrowserMainClass(utilsRef, loggerRef)
        
    
    def chrome_open_url_close_testcase_id_W10DA_16(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            #TC Step 1
            self.logger.info(
                "Checking and killing the process chromedriver.exe if any before executing the test case")
            self.chrome_maincode.killProcess()
            
            
            #TC Step 2
            self.chrome_maincode.startChromeBrowser()
            self.logger.info("Started Chrome browser and maximized the window")
            
            #TC Step 3
            self.chrome_maincode.navigateToURL(self.__google_url)
            self.logger.info("Navigated to Google.com")
            assert self.chrome_maincode.getPageTitle()=="Google", "Google Page didn't load properly"
            assert self.chrome_maincode.isElementPresentByName("q"), "Not able to find google element on Loaded page"
            self.logger.info("verified Google.com elements on Loaded page")
            
#             self.utils.sleepUntil(2)
#             self.chrome_maincode.navigateToURL(self.__gmail_url)
#             self.logger.info("Navigated to Gmail.com")
#             
#             self.utils.sleepUntil(20)
#             
#             assert "Gmail" in self.chrome_maincode.getPageTitle(), "Gmail Page didn't load properly"
#             assert self.chrome_maincode.isElementPresentByClassName("hero_home__link__desktop"), "Not able to find gmail element on Loaded page"
#             self.logger.info("verified Gmail.com elements on Loaded page")
#             
            '''Not able to navigate to Gilead site bcaz of Gilead proxy settings.  '''
            
#             self.utils.sleepUntil(2)
#             self.chrome_maincode.navigateToURL(self.__gilead_url)
#             self.logger.info("Navigated to Gilead.com")
#             self.utils.sleepUntil(5)
#             assert self.chrome_maincode.getPageTitle()=="gilead.com", "Page didn't load properly"
#             self.logger.info("verified Gilead.com elements on Loaded page")
            
            #TC Step 4
            self.chrome_maincode.closeChromeBrowser()
            self.logger.info("closed Chrome browser")
            
            #TC Step 5
            assert self.chrome_maincode.isChromeDriverProcessExist()== False, "ChromeDriver process chromedriver.exe didn't stop yet"
            self.logger.info("ChromeDriver process chromedriver.exe not there anymore in task manager")
            
            return True, "", func.co_name  
        
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        
        finally:
            '''kill the web driver '''
            self.chrome_maincode.closeChromeBrowser()
            self.logger.info("Killed the ChromeDriver process of Chrome browser app")
            self.chrome_maincode.killProcess()


    ''' TC ID: W10DA-32'''
        
    def avatier_credential_password_reset_using_chrome_testcase_id_W10DA_32(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1 Check if Chrome process "chrome.exe" is running and if so, kill it    
            self.logger.info(
                "Checking and killing the process chrome.exe if any before executing the test case")
            self.chrome_maincode.killProcess()
         
            #TC Step 2 Open Chrome browser
            self.chrome_maincode.startChromeBrowser()
            self.logger.info("Started Chrome browser and maximized the window")
            
            #TC Step 3 Navigate to URL http://reset.gilead.com
            self.chrome_maincode.navigateToURL(self.__avatier_url)
            self.logger.info("Navigated to reset.gilead.com")
            pwd_reset =  self.chrome_maincode.getTextOfAvatierPageTitle('LabelSelfServicePWResetAndSync')
            assert "SELF-SERVICE" in pwd_reset, "GILEAD Page title not found or not loaded"

            self.logger.info("verified reset.gilead.com elements on Loaded page")
   
   
            #TC Step 4 Click 'X' button to close the Chrome browser window
            self.chrome_maincode.closeChromeBrowser()
            self.logger.info("closed Chrome browser")
            
            #TC Step 5 Check in Task Manager if "chrome.exe" is running
            assert self.chrome_maincode.isChromeDriverProcessExist()== False, "ChromeDriver process chromedriver.exe still running"
            self.logger.info("ChromeDriver process chromedriver.exe not found in task manager")

            return True, "", func.co_name  
        
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        
        finally:
            '''kill the web driver '''
            self.chrome_maincode.closeChromeBrowser()
            self.logger.info("Killed the ChromeDriver process of Chrome browser app")
            self.chrome_maincode.killProcess()

