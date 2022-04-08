'''
Created on Oct 2, 2018
@author: pparamasivan
'''
import inspect

from definitions import GOOGLE_LEGACY_BROWSER_FILE, PERFORMANCE_MONITOR_PROCESS, ROOT_DIR, MISC_IMAGES_DIR
from com.gilead.main.miscellaneous.MiscellaneousMainFile import MiscellaneousMainClass
from selenium.common.exceptions import NoSuchElementException,\
    WebDriverException
from inspect import currentframe, getframeinfo

class MiscellaneousTestClass():
    __filename = GOOGLE_LEGACY_BROWSER_FILE
    __performance_monitor_process = PERFORMANCE_MONITOR_PROCESS
    __images_dir=ROOT_DIR+MISC_IMAGES_DIR
    
    def __init__(self, utilsRef, loggerRef):
        self.logger = loggerRef
        self.utils = utilsRef
        self.miscellaneous_maincode = MiscellaneousMainClass(utilsRef, loggerRef)

    
    '''TC: W10DA-33 '''  
    def check_google_legacy_browser_support_testcase_id_W10DA_33(self):
         
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            self.logger.info("Navigating the file folder using Windows explorer")
            
            #TC Step 1 & 2 Using Windows File explorer navigate to C:\Program File\Google\Legacy Browser Support folder
            assert self.utils.isFileExists(self.__filename), "Google legacy support folder file didn't Exists"

            self.logger.info("Google legacy support folder file Verification Completed")
            return True, "", func.co_name 
            
        except (AssertionError, AttributeError, TypeError) as e:
            return False, e, func.co_name
    
    '''TC: W10DA-15 '''
    def performance_monitor_id_W10DA_15(self): 
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the winium driver of Performance monitor app... ")
            
            # TC Step 1
            self.logger.info(
                "Checking and killing the process mmc.exe if any before executing the test case")
            self.miscellaneous_maincode.killProcess(self.__performance_monitor_process)
            
            # TC Step 2
            self.miscellaneous_maincode.startProgramByName_via_StartMenu("Performance Monitor")
            self.logger.info("Started program Performance monitor via start menu")
            self.utils.sleepUntil(10)
            
            assert self.miscellaneous_maincode.isGivenProcessExist("mmc.exe") , "Performance monitor process didn't start"
            
            '''connecting winium driver to Performance monitor app opened '''
            self.miscellaneous_maincode.connectPerformanceMonitorToWebDriver()
            self.logger.info("connected Performance monitor app to Winium driver ")

            assert self.miscellaneous_maincode.isElementExist(self.miscellaneous_maincode.driver, **{'Name':'Performance Monitor'}), "Performance monitor app didn't open properly"
            self.logger.info("Performance monitor GUI opened successfully")
            
            # TC Step 3
            self.miscellaneous_maincode.clickPerformanceMonitorOptions()
            self.logger.info("Clicked Performance monitor File -> Options menu")
            
            assert self.miscellaneous_maincode.isElementExist(self.miscellaneous_maincode.getPerformanceMonitorOptionsDialogElement(), **{'Name':'Disk Cleanup'}), "Options dialog didnt open up properly"
            self.logger.info("Options dialog opened up normally")
            
            self.utils.sleepUntil(3)
            self.miscellaneous_maincode.closePerformanceMonitor()
            self.miscellaneous_maincode.killProcess(self.__performance_monitor_process)
            self.logger.info("Closed Performance monitor app successfully" )
            
            self.utils.sleepUntil(10)
            assert self.miscellaneous_maincode.isGivenProcessExist("mmc.exe")==False , "Performance monitor process still open start"
            
            return True, "", func.co_name 
            
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the winium driver process of Performance monitor app")
            self.miscellaneous_maincode.killProcess(self.__performance_monitor_process)
    
    
    '''TC: W10DA-099 '''  
    def gccm_windows_activation_testcase_id_W10DA_099(self):
         
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            self.utils.executeSystemCommands("control /name microsoft.system")            
            self.utils.pageDown()
            
            self.utils.sleepUntil(8)
            # import pdb;pdb.set_trace()
            assert self.miscellaneous_maincode.isImageVisible(self.__images_dir +"windows_activation_image.PNG"), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'windows_activation_image.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.close_window()
            self.logger.info("GCCM Windows activation verification Completed")
            return True, "", func.co_name 
            
        except (AssertionError, AttributeError, TypeError) as e:
            return False, e, func.co_name

    
    '''TC: W10DA-100 '''    
    def gccm_windows_computername_testcase_id_W10DA_100(self):
       
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            assert self.miscellaneous_maincode.isSystemHostNameContainsGivenText("gilead.com"), "FQDN:"+self.utils.getSystemHostName()+ " doesn't contain gilead.com "          
            self.logger.info("FQDN contains gilead.com")
            return True, "", func.co_name 
            
        except (AssertionError, AttributeError, TypeError) as e:
            return False, e, func.co_name
        
        
    '''TC: W10DA-101 '''    
    def gccm_windows_System_type_testcase_id_W10DA_101(self):
       
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            assert self.miscellaneous_maincode.isOS64Bit(), "OS is not 64 bit "          
            self.logger.info("OS is 64 bit")
            return True, "", func.co_name 
            
        except (AssertionError, AttributeError, TypeError) as e:
            return False, e, func.co_name
    
    
    '''TC: W10DA-102 ''' 
    def gccm_windows_OS_type_testcase_id_W10DA_102(self):
       
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            assert "Windows-10-10.0" in self.miscellaneous_maincode.getWindowsPlatform(), "Its not Windows 10 machine "          
            self.logger.info("Its Windows 10 machine")
            return True, "", func.co_name 
            
        except (AssertionError, AttributeError, TypeError) as e:
            return False, e, func.co_name