'''
Created on Oct 4, 2018

@author: pparamasivan
'''
from com.gilead.main.ie_browser.InternetExplorerMainFile import InternetExplorerMainClass
import inspect
from selenium.common.exceptions import NoSuchElementException,\
    WebDriverException
from definitions import  AVATIER_URL,ROOT_DIR, IE_IMAGES_DIR

class InternetExplorerTestClass():

    __avatier_url = AVATIER_URL
    __images_dir = ROOT_DIR+IE_IMAGES_DIR
    
    def __init__(self, utilsRef, loggerRef):
        self.utils = utilsRef
        self.logger = loggerRef
        self.ie_maincode = InternetExplorerMainClass(utilsRef, loggerRef)

    ''' TC_ID_W10DA-30'''
    def avatier_credential_password_reset_using_IE_Browser_testcase_id_W10DA_30(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1  Check if Internet Explorer process  is running and if so, kill it 
            self.logger.info(
                "Checking and killing the process iexplore.exe if any before executing the test case")
            self.ie_maincode.killProcess()
         
            #TC Step 2 Open Internet Explorer browser
            self.ie_maincode.startInternetExplorerBrowser()
            self.logger.info("Started InternetExplorer browser and maximized the window")
            self.utils.sleepUntil(5)
            #TC Step 3 Navigate to URL http://reset.gilead.com
            self.ie_maincode.navigateToURL(self.__avatier_url)
            self.utils.sleepUntil(2)
            self.logger.info("Navigated to reset.gilead.com")

            pwd_reset =  self.ie_maincode.getTextOfAvatierPageTitle("LabelSelfServicePWResetAndSync")
            assert "SELF-SERVICE" in pwd_reset, "GILEAD Page title not found or not loaded"
            self.logger.info("verified reset.gilead.com element on Loaded Edge browser page")

            #TC Step 4 Click 'X' button to close the Internet Explorer browser window
            self.ie_maincode.closeInternetExplorerBrowser()
            self.logger.info("closed Internet Explorer browser")
            
            self.utils.sleepUntil(5)
            
            #TC Step 6 Check in Task Manager if  Internet Explorer processes are running
            if self.ie_maincode.isInternetExplorerDriverProcessExist():
                self.ie_maincode.killProcess()
    
            assert self.ie_maincode.isInternetExplorerDriverProcessExist()== False, "Internet Explorer process iexplore.exe still running"
            self.logger.info("Internet Explorer process iexplore.exe not found in task manager")
            
            return True, "", func.co_name  
        
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name

    ''' TC_ID_W10DA-117'''

    def tier0_application_IEversion_testcase_id_W10DA_117(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1  Check if Internet Explorer process  is running and if so, kill it 
            self.logger.info(
                "Checking and killing the process iexplore.exe if any before executing the test case")
            self.ie_maincode.killProcess()
         
            #TC Step 2 Open Internet Explorer browser
            self.ie_maincode.openApplication()
            self.logger.info("Started InternetExplorer browser and maximized the window")
            self.utils.sleepUntil(5)

            #TC Step 3
            assert self.ie_maincode.isImageVisible(self.__images_dir + 'button_help.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'button_help.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.ie_maincode.doubleClickOnImage(self.__images_dir + 'button_help.PNG', 40, 8 )
            self.utils.sleepUntil(5)

            #TC Step 4
            assert self.ie_maincode.isImageVisible(self.__images_dir + 'button_about.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'button_about.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.ie_maincode.doubleClickOnImage(self.__images_dir + 'button_about.PNG', 40, 8 )
            self.utils.sleepUntil(5)

            #TC Step 5
            assert self.ie_maincode.check_if_ImageExist(1), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'windows_checking.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(5)

            return True, "", func.co_name          
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            self.ie_maincode.killProcess()
    

    ''' TC_ID_W10DA-118'''

    def tier0_application_internetoptions_testcase_id_W10DA_118(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1  Check if Internet Explorer process  is running and if so, kill it 
            self.logger.info(
                "Checking and killing the process iexplore.exe if any before executing the test case")
            self.ie_maincode.killProcess()
         
            #TC Step 2 Open Internet Explorer browser
            self.ie_maincode.openApplication()
            self.logger.info("Started InternetExplorer browser and maximized the window")
            self.utils.sleepUntil(5)
         
            #TC Step 3
            assert self.ie_maincode.isImageVisible(self.__images_dir + 'button_tools.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'button_tools.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.ie_maincode.doubleClickOnImage(self.__images_dir + 'button_tools.PNG', 40, 8 )
            self.utils.sleepUntil(5)

            #TC Step 4
            assert self.ie_maincode.isImageVisible(self.__images_dir + 'button_IEoption.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'button_IEoption.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.ie_maincode.doubleClickOnImage(self.__images_dir + 'button_IEoption.PNG', 40, 8 )
            self.utils.sleepUntil(5)

            #TC Step 5
            assert self.ie_maincode.check_if_ImageExist(2), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'windows_checking.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno ) 
            self.utils.sleepUntil(5)      
            
            return True, "", func.co_name  
        
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            #TC Step 6 
            self.ie_maincode.killProcess()
        
        
    ''' TC_ID_W10DA-119'''

    def tier0_application_internet_testcase_id_W10DA_119(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1  Check if Internet Explorer process  is running and if so, kill it 
            self.logger.info(
                "Checking and killing the process iexplore.exe if any before executing the test case")
            self.ie_maincode.killProcess()
         
            #TC Step 2 Open Internet Explorer browser
            self.ie_maincode.openApplication()
            self.logger.info("Started InternetExplorer browser and maximized the window")
            self.utils.sleepUntil(5)
         
            #TC Step 3
            assert self.ie_maincode.isImageVisible(self.__images_dir + 'button_tools.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'button_tools.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.ie_maincode.doubleClickOnImage(self.__images_dir + 'button_tools.PNG', 40, 8 )
            self.utils.sleepUntil(5)

            #TC Step 4
            assert self.ie_maincode.isImageVisible(self.__images_dir + 'button_IEoption.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'button_IEoption.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.ie_maincode.doubleClickOnImage(self.__images_dir + 'button_IEoption.PNG', 40, 8 )
            self.utils.sleepUntil(5)

            #TC Step 5
            assert self.ie_maincode.isImageVisible(self.__images_dir + 'button_security.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'button_security.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.ie_maincode.doubleClickOnImage(self.__images_dir + 'button_security.PNG', 40, 8 )
            self.utils.sleepUntil(5)

            #TC Step 6
            assert self.ie_maincode.isImageVisible(self.__images_dir + 'button_internet.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'button_internet.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.ie_maincode.doubleClickOnImage(self.__images_dir + 'button_internet.PNG', 40, 8 )
            self.utils.sleepUntil(5)

            #TC Step 7
            assert self.ie_maincode.check_if_ImageExist(3), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'windows_checking.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno ) 
            self.utils.sleepUntil(5)  
    
            return True, "", func.co_name  
        
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            #TC Step 8 
            self.ie_maincode.killProcess()
           
        
    ''' TC_ID_W10DA-121 '''

    def tier0_application_internet_trustedsite_testcase_id_W10DA_121(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1  Check if Internet Explorer process  is running and if so, kill it 
            self.logger.info(
                "Checking and killing the process iexplore.exe if any before executing the test case")
            self.ie_maincode.killProcess()
         
            #TC Step 2 Open Internet Explorer browser
            self.ie_maincode.openApplication()
            self.logger.info("Started InternetExplorer browser and maximized the window")
            self.utils.sleepUntil(5)
         
            #TC Step 3
            assert self.ie_maincode.isImageVisible(self.__images_dir + 'button_tools.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'button_tools.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.ie_maincode.doubleClickOnImage(self.__images_dir + 'button_tools.PNG', 40, 8 )
            self.utils.sleepUntil(5)

            #TC Step 4
            assert self.ie_maincode.isImageVisible(self.__images_dir + 'button_IEoption.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'button_IEoption.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.ie_maincode.doubleClickOnImage(self.__images_dir + 'button_IEoption.PNG', 40, 8 )
            self.utils.sleepUntil(5)

            #TC Step 5
            assert self.ie_maincode.isImageVisible(self.__images_dir + 'button_security.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'button_security.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.ie_maincode.doubleClickOnImage(self.__images_dir + 'button_security.PNG', 40, 8 )
            self.utils.sleepUntil(5)

            #TC Step 6
            assert self.ie_maincode.isImageVisible(self.__images_dir + 'button_trustedsites.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'button_trustedsites.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.ie_maincode.doubleClickOnImage(self.__images_dir + 'button_trustedsites.PNG', 40, 8 )
            self.utils.sleepUntil(5)

            #TC Step 5
            assert self.ie_maincode.check_if_ImageExist(4), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'windows_checking.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno ) 
            self.utils.sleepUntil(5)  
        
            return True, "", func.co_name  
        
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            self.ie_maincode.killProcess()
        
        
    ''' TC_ID_W10DA-120 '''

    def tier0_application_internet_localintranet_testcase_id_W10DA_120(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1  Check if Internet Explorer process  is running and if so, kill it 
            self.logger.info(
                "Checking and killing the process iexplore.exe if any before executing the test case")
            self.ie_maincode.killProcess()
         
            #TC Step 2 Open Internet Explorer browser
            self.ie_maincode.openApplication()
            self.logger.info("Started InternetExplorer browser and maximized the window")
            self.utils.sleepUntil(5)
         
            #TC Step 3
            assert self.ie_maincode.isImageVisible(self.__images_dir + 'button_tools.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'button_tools.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.ie_maincode.doubleClickOnImage(self.__images_dir + 'button_tools.PNG', 40, 8 )
            self.utils.sleepUntil(5)

            #TC Step 4
            assert self.ie_maincode.isImageVisible(self.__images_dir + 'button_IEoption.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'button_IEoption.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.ie_maincode.doubleClickOnImage(self.__images_dir + 'button_IEoption.PNG', 40, 8 )
            self.utils.sleepUntil(5)

            #TC Step 5
            assert self.ie_maincode.isImageVisible(self.__images_dir + 'button_security.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'button_security.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.ie_maincode.doubleClickOnImage(self.__images_dir + 'button_security.PNG', 40, 8 )
            self.utils.sleepUntil(5)

            #TC Step 6
            assert self.ie_maincode.check_if_ImageExist(8), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'windows_checking.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno ) 
            self.utils.sleepUntil(5)

            #TC Step 7
            assert self.ie_maincode.isImageVisible(self.__images_dir + 'button_defaultlevel.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'button_defaultlevel.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.ie_maincode.doubleClickOnImage(self.__images_dir + 'button_defaultlevel.PNG', 40, 8 )
            self.utils.sleepUntil(5)  

             #TC Step 8
            assert self.ie_maincode.check_if_ImageExist(9), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'windows_checking.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno ) 
            self.utils.sleepUntil(5)
          
            return True, "", func.co_name  
        
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            self.ie_maincode.killProcess()


    ''' TC_ID_W10DA-122 '''

    def tier0_application_internet_restrictedsites_testcase_id_W10DA_122(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1  Check if Internet Explorer process  is running and if so, kill it 
            self.logger.info(
                "Checking and killing the process iexplore.exe if any before executing the test case")
            self.ie_maincode.killProcess()
         
            #TC Step 2 Open Internet Explorer browser
            self.ie_maincode.openApplication()
            self.logger.info("Started InternetExplorer browser and maximized the window")
            self.utils.sleepUntil(5)
         
            #TC Step 3
            assert self.ie_maincode.isImageVisible(self.__images_dir + 'button_tools.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'button_tools.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.ie_maincode.doubleClickOnImage(self.__images_dir + 'button_tools.PNG', 40, 8 )
            self.utils.sleepUntil(5)

            #TC Step 4
            assert self.ie_maincode.isImageVisible(self.__images_dir + 'button_IEoption.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'button_IEoption.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.ie_maincode.doubleClickOnImage(self.__images_dir + 'button_IEoption.PNG', 40, 8 )
            self.utils.sleepUntil(5)

            #TC Step 5
            assert self.ie_maincode.isImageVisible(self.__images_dir + 'button_security.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'button_security.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.ie_maincode.doubleClickOnImage(self.__images_dir + 'button_security.PNG', 40, 8 )
            self.utils.sleepUntil(5)

            #TC Step 6
            assert self.ie_maincode.isImageVisible(self.__images_dir + 'button_restrictedsites.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'button_restrictedsites.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.ie_maincode.doubleClickOnImage(self.__images_dir + 'button_restrictedsites.PNG', 40, 8 )
            self.utils.sleepUntil(5)

            #TC Step 5
            assert self.ie_maincode.check_if_ImageExist(5), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'windows_checking.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno ) 
            self.utils.sleepUntil(5)  
           
            return True, "", func.co_name
         
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            self.ie_maincode.killProcess()
       
        

    ''' TC_ID_W10DA-123 '''

    def tier0_application_internet_connections_testcase_id_W10DA_123(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1  Check if Internet Explorer process  is running and if so, kill it 
            self.logger.info(
                "Checking and killing the process iexplore.exe if any before executing the test case")
            self.ie_maincode.killProcess()
         
            #TC Step 2 Open Internet Explorer browser
            self.ie_maincode.openApplication()
            self.logger.info("Started InternetExplorer browser and maximized the window")
            self.utils.sleepUntil(5)
         
            #TC Step 3
            assert self.ie_maincode.isImageVisible(self.__images_dir + 'button_tools.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'button_tools.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.ie_maincode.doubleClickOnImage(self.__images_dir + 'button_tools.PNG', 40, 8 )
            self.utils.sleepUntil(5)

            #TC Step 4
            assert self.ie_maincode.isImageVisible(self.__images_dir + 'button_IEoption.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'button_IEoption.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.ie_maincode.doubleClickOnImage(self.__images_dir + 'button_IEoption.PNG', 40, 8 )
            self.utils.sleepUntil(5)

            #TC Step 5
            assert self.ie_maincode.isImageVisible(self.__images_dir + 'button_connections.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'button_connections.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.ie_maincode.doubleClickOnImage(self.__images_dir + 'button_connections.PNG', 40, 8 )
            self.utils.sleepUntil(5)

            #TC Step 6
            assert self.ie_maincode.isImageVisible(self.__images_dir + 'button_LANsettings.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'button_LANsettings.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.ie_maincode.doubleClickOnImage(self.__images_dir + 'button_LANsettings.PNG', 40, 8 )
            self.utils.sleepUntil(5)

            #TC Step 7
            assert self.ie_maincode.check_if_ImageExist(6), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'windows_checking.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno ) 
            self.utils.sleepUntil(5)  
        
            return True, "", func.co_name
         
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            self.ie_maincode.killProcess()
         

    ''' TC_ID_W10DA-124 '''

    def tier0_application_general_browserhistory_testcase_id_W10DA_124(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1  Check if Internet Explorer process  is running and if so, kill it 
            self.logger.info(
                "Checking and killing the process iexplore.exe if any before executing the test case")
            self.ie_maincode.killProcess()
         
            #TC Step 2 Open Internet Explorer browser
            self.ie_maincode.openApplication()
            self.logger.info("Started InternetExplorer browser and maximized the window")
            self.utils.sleepUntil(5)
         
            #TC Step 3
            assert self.ie_maincode.isImageVisible(self.__images_dir + 'button_tools.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'button_tools.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.ie_maincode.doubleClickOnImage(self.__images_dir + 'button_tools.PNG', 40, 8 )
            self.utils.sleepUntil(5)

            #TC Step 4
            assert self.ie_maincode.isImageVisible(self.__images_dir + 'button_IEoption.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'button_IEoption.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.ie_maincode.doubleClickOnImage(self.__images_dir + 'button_IEoption.PNG', 40, 8 )
            self.utils.sleepUntil(5)

            #TC Step 5
            assert self.ie_maincode.isImageVisible(self.__images_dir + 'button_settings.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'button_connections.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.ie_maincode.doubleClickOnImage(self.__images_dir + 'button_settings.PNG', 40, 8 )
            self.utils.sleepUntil(5)

            #TC Step 6
            assert self.ie_maincode.check_if_ImageExist(7), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'windows_checking.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno ) 
            self.utils.sleepUntil(5)  
    
            return True, "", func.co_name
         
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            self.ie_maincode.killProcess()
        
        
        
        
        