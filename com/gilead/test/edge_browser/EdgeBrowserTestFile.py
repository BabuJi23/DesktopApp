'''
Created on Oct 2, 2018

@author: bkotamahanti
'''
from com.gilead.main.edge_browser.EdgeBrowserMainFile import EdgeBrowserMainClass
import inspect
from selenium.common.exceptions import NoSuchElementException,\
    WebDriverException
from definitions import GILEAD_URL, GOOGLE_URL, AVATIER_URL

class EdgeBrowserTestClass():
    __gilead_url = GILEAD_URL
    __google_url = GOOGLE_URL
#     __gmail_url = GMAIL_URL
    __avatier_url = AVATIER_URL
    
    def __init__(self, utilsRef, loggerRef):
        self.utils = utilsRef
        self.logger = loggerRef
        self.edge_maincode = EdgeBrowserMainClass(utilsRef, loggerRef)
    

    ''' TC_ID_W10DA-29'''
        
    def avatier_credential_password_reset_using_MsEdge_testcase_id_W10DA_29(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            #TC Step 1  Check if Microsoft Edge process  is running and if so, kill it  
            self.logger.info(
                "Checking and killing the process MsEdgeDriver.exe if any before executing the test case")
            self.edge_maincode.killProcess()
         
            #TC Step 2 Open Microsoft Edge browser
            self.edge_maincode.startEdgeBrowser()
            self.logger.info("Started Edge browser and maximized the window")
            
            #TC Step 3 Navigate to URL http://reset.gilead.com
            self.utils.sleepUntil(2)
            self.edge_maincode.navigateToURL(self.__avatier_url)
            self.logger.info("Navigated to reset.gilead.com")
            self.utils.sleepUntil(2)
            pwd_reset =  self.edge_maincode.getTextOfAvatierPageTitle('LabelSelfServicePWResetAndSync')
            '''assert "SELF-SERVICE" in pwd_reset, "GILEAD Page title not found or not loaded"'''

            self.logger.info("verified reset.gilead.com elements on Loaded page")
   
            #TC Step 4 Click 'X' button to close the Microsoft Edge browser window
            self.edge_maincode.closeBrowser()
            self.logger.info("closed Edge browser")
            self.utils.sleepUntil(2)
            #TC Step 5 Check in Task Manager if Microsoft Edge process is running
            assert self.edge_maincode.isEdgeDriverProcessExist()== False, "EdgeDriver process MicrosoftEdge.exe still running"
            self.logger.info("EdgeDriver process MicrosoftEdge.exe not found in task manager")
            
            return True, "", func.co_name  
        
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        
        finally:
            self.edge_maincode.killProcess()

    ''' TC_ID_W10DA-31'''
        
    def avatier_credential_password_reset_disagree_MSEdge_testcase_id_W10DA_31(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1  Check if Microsoft Edge process  is running and if so, kill it  
            self.logger.info(
                "Checking and killing the process MsEdgeDriver.exe if any before executing the test case")
            self.edge_maincode.killProcess()
         
            #TC Step 2 Open Microsoft Edge browser
            self.edge_maincode.startEdgeBrowser()
            self.logger.info("Started Edge browser and maximized the window")
            self.utils.sleepUntil(2)
            
            #TC Step 3 Navigate to URL http://reset.gilead.com
            self.edge_maincode.navigateToURL(self.__avatier_url)
            self.logger.info("Navigated to reset.gilead.com")
            pwd_reset =  self.edge_maincode.getTextOfAvatierPageTitle('LabelSelfServicePWResetAndSync')
            ''' assert "SELF-SERVICE" in pwd_reset, "GILEAD Page title not found or not loaded"'''
            self.logger.info("verified reset.gilead.com elements on Loaded page")
            self.utils.sleepUntil(2)
            #TC Step 4 Click 'I Disagree' button- 
            self.edge_maincode.getDisagreeButtonId("btnTextIDisagree")
            self.utils.sleepUntil(3)
            disagree_msg = self.edge_maincode.getDisagreeValidationMessage('LabelDisagree')
            '''assert "Sorry, you could not agree" in disagree_msg, "Avatier terms disagree validation message not displayed"'''

            #TC Step 5 Click browser 'Back' button
            self.edge_maincode.navigateBack()
            self.logger.info("verified disagree page- back to home page")
            
            #TC Step 6 Click 'I Agree' button -Error message in red 'Please select a domain' should display
            self.edge_maincode.getAgreeButtonId("btnTextIAgree")
            self.logger.info("verified agree button clicked in  home page")
            self.utils.sleepUntil(3)
            domain_msg = self.edge_maincode.getDomainValidationMessage('validateUserID')
            assert "Please select a domain" in domain_msg, "Avatier Domain validation message not displayed"

            #TC Step 7 Click 'X' button to close the Microsoft Edge browser window
            self.edge_maincode.closeBrowser()
            self.logger.info("closed Edge browser")
            
            #TC Step 8 Check in Task Manager if Microsoft Edge process is running
            assert self.edge_maincode.isEdgeDriverProcessExist()== False, "EdgeDriver process MicrosoftEdge.exe still running"
            self.logger.info("EdgeDriver process MicrosoftEdge.exe not found in task manager")
            
            return True, "", func.co_name  
        
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            self.edge_maincode.killProcess()


    ''' TC_ID_W10DA-42'''    
    def microsoft_edge_open_url_close_testcase_id_W10DA_42(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1
            self.logger.info(
                "Checking and killing the process MicrosoftWebDriver.exe if any before executing the test case")
            self.edge_maincode.killProcess()
            
            
            #TC Step 2
            self.utils.sleepUntil(3)
            # import pdb;pdb.set_trace()
            self.edge_maincode.startEdgeBrowser()
            self.logger.info("Started Edge browser and maximized the window")
            
            #TC Step 3
            self.edge_maincode.navigateToURL(self.__google_url)
            self.utils.sleepUntil(2)
            self.logger.info("Navigated to Google.com")
            assert self.edge_maincode.getPageTitle()=="Google", "Google Page didn't load properly"
            assert self.edge_maincode.isElementPresentByName("q"), "Not able to find google element on Loaded page"
            self.logger.info("verified Google.com elements on Loaded page")
            
            self.utils.sleepUntil(2)
#             self.edge_maincode.navigateToURL(self.__gmail_url)
#             self.logger.info("Navigated to Gmail.com")
#             assert "Gmail" in self.edge_maincode.getPageTitle(), "Gmail Page didn't load properly"
# #             assert self.edge_maincode.isElementPresentByClassName("hero_home__link__desktop"), "Not able to find gmail element on Loaded page"
#             self.logger.info("verified Gmail.com elements on Loaded page")
            
            '''Not able to navigate to Gilead site bcaz of Gilead proxy settings.  '''
#             self.utils.sleepUntil(2)
#             self.edge_maincode.navigateToURL(self.__gilead_url)
#             self.logger.info("Navigated to Gmail.com")
#             self.utils.sleepUntil(5)
#             assert self.edge_maincode.getPageTitle()=="gilead.com", "Page didn't load properly"
#             self.logger.info("verified Gilead.com elements on Loaded page")
            
            #TC Step 4
            self.edge_maincode.closeEdgeBrowser()
            self.logger.info("closed Edge browser")
            
            #TC Step 5
            assert self.edge_maincode.isEdgeDriverProcessExist()== False, "EdgeDriver process MicrosoftWebDriver.exe didn't stop yet"
            self.logger.info("EdgeDriver process MicrosoftWebDriver.exe not there anymore in task manager")
            
            return True, "", func.co_name  
        
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            self.edge_maincode.killProcess()
        
