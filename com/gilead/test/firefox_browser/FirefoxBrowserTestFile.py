'''
Created on Oct 2, 2018

@author: bkotamahanti
'''
from com.gilead.main.firefox_browser.FirefoxBrowserMainFile import FirefoxBrowserMainClass
from definitions import GILEAD_URL, GOOGLE_URL
import inspect
from selenium.common.exceptions import NoSuchElementException,\
    WebDriverException

class FirefoxBrowserTestClass():
    __gilead_url = GILEAD_URL
    __google_url = GOOGLE_URL
#     __gmail_url = GMAIL_URL
    
    
    def __init__(self, utilsRef, loggerRef):
        self.utils = utilsRef
        self.logger = loggerRef
        self.firefox_maincode = FirefoxBrowserMainClass(utilsRef, loggerRef)
        
    def mozilla_firefox_open_url_close_testcase_id_W10DA_43(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1
            self.logger.info(
                "Checking and killing the process firefox.exe if any before executing the test case")
            self.firefox_maincode.killProcess()
            
            
            #TC Step 2
            self.firefox_maincode.startFirefoxBrowser()
            self.logger.info("Started Firefox browser and maximized the window")
            
            #TC Step 3
            self.firefox_maincode.navigateToURL(self.__google_url)
            self.logger.info("Navigated to Google.com")
            assert self.firefox_maincode.getPageTitle()=="Google", "Google Page didn't load properly"
            assert self.firefox_maincode.isElementPresentByName("q"), "Not able to find google element on Loaded page"
            self.logger.info("verified Google.com elements on Loaded page")
            
#             self.utils.sleepUntil(3)
#             self.firefox_maincode.navigateToURL(self.__gmail_url)
#             self.logger.info("Navigated to Gmail.com")
#             assert "Gmail" in self.firefox_maincode.getPageTitle(), "Gmail Page didn't load properly"
#             assert self.firefox_maincode.isElementPresentByClassName("hero_home__link__desktop"), "Not able to find gmail element on Loaded page"
#             self.logger.info("verified Gmail.com elements on Loaded page")
#             
#             '''Not able to navigate to Gilead site bcaz of Gilead proxy settings.  '''
#             self.utils.sleepUntil(3)
#             self.edge_maincode.navigateToURL(self.__gilead_url)
#             self.logger.info("Navigated to Gmail.com")
#             self.utils.sleepUntil(5)
#             assert self.edge_maincode.getPageTitle()=="gilead.com", "Page didn't load properly"
#             self.logger.info("verified Gilead.com elements on Loaded page")
            
            #TC Step 4
            self.firefox_maincode.closeFirefoxBrowser()
            self.logger.info("closed Firefox browser")
            
            #TC Step 5
            assert self.firefox_maincode.isFirefoxDriverProcessExist()== False, "FirefoxDriver process firefox.exe didn't stop yet"
            self.logger.info("FirefoxDriver process firefox.exe not there anymore in task manager")
            
            return True, "", func.co_name  
        
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name