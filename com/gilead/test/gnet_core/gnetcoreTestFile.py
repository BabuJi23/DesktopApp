'''
Created on June 28, 2019

@author: ngarimella
'''
# from com.gilead.main.ie_browser.InternetExplorerMainFile import InternetExplorerMainClass
from com.gilead.main.gnet_core.gnetcoreMainFile import gnetcoreMainClass
import inspect
from selenium.common.exceptions import NoSuchElementException,\
    WebDriverException
from definitions import  AVATIER_URL,ROOT_DIR, GNET_IMAGES_DIR

class gnetcoreTestClass():

    __avatier_url = AVATIER_URL
    __images_dir = ROOT_DIR+GNET_IMAGES_DIR
    
    def __init__(self, utilsRef, loggerRef):
        self.utils = utilsRef
        self.logger = loggerRef
        self.gnet_maincode = gnetcoreMainClass(utilsRef, loggerRef)

    ''' TC_ID_W10DA-143 '''

    def gnet_core_applications_calendar_W10DA_143(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1  Check if chrome process  is running and if so, kill it 
            self.logger.info(
                "Checking and killing the process iexplore.exe if any before executing the test case")
            self.gnet_maincode.killProcess()
         
            #TC Step 2 Open Chrome Browser
            self.gnet_maincode.startChromeBrowser()

            self.logger.info("Started google Chrome browser and maximized the window")
            self.utils.sleepUntil(5)

            self.gnet_maincode.press_esc()

            #TC Step 3 Click Calendar Button
            self.gnet_maincode.clickCalendarButton()
          
            return True, "", func.co_name
         
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            self.gnet_maincode.killProcess()

    ''' TC_ID_W10DA-144 '''

    def gnet_core_applications_gxplearn_W10DA_144(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1  Check if chrome process  is running and if so, kill it 
            self.logger.info(
                "Checking and killing the process iexplore.exe if any before executing the test case")
            self.gnet_maincode.killProcess()
         
            #TC Step 2 Open Chrome Browser
            self.gnet_maincode.startChromeBrowser()

            self.logger.info("Started google Chrome browser and maximized the window")
            self.utils.sleepUntil(5)

            self.gnet_maincode.press_esc()

            #TC Step 3 Click GxPLearn Button
            self.gnet_maincode.clickGxPLearnButton()
          
            return True, "", func.co_name
         
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            self.gnet_maincode.killProcess()

    ''' TC_ID_W10DA-145 '''

    def gnet_core_applications_workday_W10DA_145(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1  Check if chrome process  is running and if so, kill it 
            self.logger.info(
                "Checking and killing the process iexplore.exe if any before executing the test case")
            self.gnet_maincode.killProcess()
         
            #TC Step 2 Open Chrome Browser
            self.gnet_maincode.startChromeBrowser()

            self.logger.info("Started google Chrome browser and maximized the window")
            self.utils.sleepUntil(5)

            self.gnet_maincode.press_esc()

            #TC Step 3 Click Workday button
            self.gnet_maincode.clickWorkdayTab()
          
            return True, "", func.co_name
         
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            self.gnet_maincode.killProcess()
        
    ''' TC_ID_W10DA-146 '''

    def gnet_core_applications_sparc_W10DA_146(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1  Check if chrome process  is running and if so, kill it 
            self.logger.info(
                "Checking and killing the process iexplore.exe if any before executing the test case")
            self.gnet_maincode.killProcess()
         
            #TC Step 2 Open Chrome Browser
            self.gnet_maincode.startChromeBrowser()

            self.logger.info("Started google Chrome browser and maximized the window")
            self.utils.sleepUntil(5)

            self.gnet_maincode.press_esc()

            #TC Step 3 Click Sparc button
            self.gnet_maincode.clickSparcTab()
          
            return True, "", func.co_name
         
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            self.gnet_maincode.killProcess()

        
    ''' TC_ID_W10DA-147 '''

    def gnet_core_applications_glearn_W10DA_147(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1  Check if chrome process  is running and if so, kill it 
            self.logger.info(
                "Checking and killing the process iexplore.exe if any before executing the test case")
            self.gnet_maincode.killProcess()
         
            #TC Step 2 Open Chrome Browser
            self.gnet_maincode.startChromeBrowser()

            self.logger.info("Started google Chrome browser and maximized the window")
            self.utils.sleepUntil(5)

            self.gnet_maincode.press_esc()

            #TC Step 3 Click GLearn button
            self.gnet_maincode.clickGLearnTab()

            #TC Step 4 Click View All to see history
            assert self.gnet_maincode.isImageVisible( self.__images_dir + 'history_viewall.PNG'), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'history_viewall.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2)
            self.gnet_maincode.doubleClickOnImage(self.__images_dir + 'history_viewall.PNG', 40, 8)
            self.utils.sleepUntil(6)
            return True, "", func.co_name
         
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            self.gnet_maincode.killProcess()

    ''' TC_ID_W10DA-148 '''

    def gnet_core_applications_search_W10DA_148(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1  Check if chrome process  is running and if so, kill it 
            self.logger.info(
                "Checking and killing the process iexplore.exe if any before executing the test case")
            self.gnet_maincode.killProcess()
         
            #TC Step 2 Open Chrome Browser
            self.gnet_maincode.startChromeBrowser()

            self.logger.info("Started google Chrome browser and maximized the window")
            self.utils.sleepUntil(5)

            self.gnet_maincode.press_esc()

            #TC Step 3 Click Search button
            self.gnet_maincode.clickSearchTab()
          
            return True, "", func.co_name
         
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            self.gnet_maincode.killProcess()
        
    ''' TC_ID_W10DA-149 '''

    def gnet_core_applications_ethics_compliance_W10DA_149(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1  Check if chrome process  is running and if so, kill it 
            self.logger.info(
                "Checking and killing the process iexplore.exe if any before executing the test case")
            self.gnet_maincode.killProcess()
         
            #TC Step 2 Open Chrome Browser
            self.gnet_maincode.startChromeBrowser()

            self.logger.info("Started google Chrome browser and maximized the window")
            self.utils.sleepUntil(5)

            self.gnet_maincode.press_esc()

            #TC Step 3 Click Ethics button
            self.gnet_maincode.clickEthicsTab()
          
            return True, "", func.co_name
         
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            self.gnet_maincode.killProcess()

    ''' TC_ID_W10DA-150 '''

    def gnet_core_applications_gvdi_W10DA_150(self):
        try:
            func = inspect.currentframe().f_back.f_code
            #TC Step 1 start
            self.utils.startProgramByName_via_StartMenu("VMWare Horizon Client")
            self.utils.sleepUntil(5)
            
            #TC Step 2 check if image exist
            assert self.gnet_maincode.isImageVisible( self.__images_dir + 'gvdi_image.PNG'), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'gvdi_image.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(5)
            
            #TC Step 3 Click Ethics button
            self.gnet_maincode.close_window()
          
            return True, "", func.co_name
         
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name


    ''' TC_ID_W10DA-151 '''

    def gnet_core_applications_IDM_W10DA_151(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1  Check if chrome process  is running and if so, kill it 
            self.logger.info(
                "Checking and killing the process iexplore.exe if any before executing the test case")
            self.gnet_maincode.killProcess()
         
            #TC Step 2 Open Chrome Browser
            self.gnet_maincode.OpenIDMwindow()
            self.logger.info("Started google Chrome browser and maximized the window")
            self.utils.sleepUntil(5)

            #TC Step 3 Click hyperlinks
            self.gnet_maincode.clickIDMTab()
          
            return True, "", func.co_name
         
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            self.gnet_maincode.killProcess()

    ''' TC_ID_W10DA-152 '''

    def gnet_core_applications_printer_W10DA_152(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1  Start the program
            self.utils.startProgramByName_via_StartMenu("\\\\fcprint")
            self.utils.sleepUntil(5)

            #TC Step 2 Check if the priter exist
            assert self.gnet_maincode.isImageVisible(self.__images_dir + 'printer_image.PNG'), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'printer_image.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(5)
            self.logger.info("Remote Foster City printers are available")
            self.gnet_maincode.close_window()

            #TC Step 3 Stockley Park printers
            self.utils.startProgramByName_via_StartMenu("\\\\lnprint02\\")
            self.utils.sleepUntil(5)

            #TC Step 4 Check if the Stockley Park printer exist
            assert self.gnet_maincode.isImageVisible(self.__images_dir + 'printer_image.PNG'), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'printer_image.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(5)
            self.logger.info("Remote Stocklry Park printers are available")
            self.gnet_maincode.close_window()

            #TC Step 5 Munich(Germany) printers
            self.utils.startProgramByName_via_StartMenu("\\\\muprint02\\")
            self.utils.sleepUntil(5)

            #TC Step 6 Check if the Munich(Germany) printer exist
            assert self.gnet_maincode.isImageVisible(self.__images_dir + 'printer_image.PNG'), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'printer_image.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(5)
            self.logger.info("Remote Munich(Germany) printers are available")
            
            return True, "", func.co_name
         
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            self.gnet_maincode.close_window()

    ''' TC_ID_W10DA-153 '''

    def gnet_core_applications_GVault_W10DA_153(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1  Check if chrome process  is running and if so, kill it 
            self.logger.info(
                "Checking and killing the process iexplore.exe if any before executing the test case")
         
            #TC Step 2 Open Chrome Browser
            self.gnet_maincode.OpenGVaultwindow()

            self.logger.info("Started google Chrome browser and Opened GVault window")
            self.utils.sleepUntil(5)

            #TC Step 3 Click hyperlinks
            self.gnet_maincode.clickGVaultTab()
          
            return True, "", func.co_name
         
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        # finally:
        #     self.gnet_maincode.killProcess()





      
         
         
        
        
        
        
        