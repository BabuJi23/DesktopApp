'''
Created on Nov 6, 2018

@author: bkotamahanti
'''
from com.gilead.main.devicemanager.DeviceManagerMainFile import DeviceManagerMainClass
from definitions import DEVICE_MANAGER_BINARY_PATH, DEVICE_MANAGER_IMAGES_DIR, ROOT_DIR
import inspect
from selenium.common.exceptions import NoSuchElementException,\
    WebDriverException
from inspect import currentframe, getframeinfo
import pyautogui
    



class DeviceManagerTestClass():
    __devicemanager_binary_path = DEVICE_MANAGER_BINARY_PATH
    __images_dir = ROOT_DIR+DEVICE_MANAGER_IMAGES_DIR
#     exceptions = [] 
    
    def __init__(self, utilsRef, loggerRef):
        self.utils = utilsRef
        self.logger = loggerRef
        self.dm_maincode = DeviceManagerMainClass(utilsRef, loggerRef)
    
    '''TC ID: W10DA_079 '''
    def devicemanager_computer_drive_testcase_id_W10DA_079(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1
            self.logger.info("Checking and killing the process mmc.exe if any before executing the test case")
            self.dm_maincode.killProcess()
            
            #TC Step 2
            self.dm_maincode.openApplication(self.__devicemanager_binary_path)
            self.utils.sleepUntil(2)
            
            #TC Step 3
            assert self.dm_maincode.isImageVisible( self.__images_dir + 'computer_icon_image.PNG'), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'computer_icon_image.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2)
            self.dm_maincode.doubleClickOnImage(self.__images_dir + 'computer_icon_image.PNG', 40, 8)
    
            #TC Step 4
            assert self.dm_maincode.isImageVisible(self.__images_dir + 'computer_ACPI_driver_icon.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'computer_ACPI_driver_icon.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.dm_maincode.doubleClickOnImage(self.__images_dir + 'computer_ACPI_driver_icon.PNG', 40, 8 )
            self.utils.sleepUntil(10)
            
            #TC Step 5
#             assert self.dm_maincode.isWorkingProperlyImageVisisble(1), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'generic_working_properly.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
           

            
            #TC Step 6
            self.utils.sleepUntil(2)
            self.dm_maincode.close_window()
            
#             assert ( self.exceptions == [] ) , "Exceptions recorded.....===================>{}".format(self.iterateExceptions(self.exceptions))
#             assert ( self.exceptions == [] ) , "Exceptions recorded.....===================>{}".format(list(map(self.iterateExceptions, self.exceptions)))
#             self.logger.info("No Exceptions recorded")
            return True, "", func.co_name  

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        
        finally:
            self.dm_maincode.killProcess()
            
    '''TC ID: W10DA_080 '''
    def devicemanager_disk_drive_testcase_id_W10DA_080(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1
            self.logger.info("Checking and killing the process mmc.exe if any before executing the test case")
            self.dm_maincode.killProcess()
            
            #TC Step 2
            self.dm_maincode.openApplication(self.__devicemanager_binary_path)
            self.utils.sleepUntil(2)
            
            #TC Step 3
            assert self.dm_maincode.isImageVisible( self.__images_dir + 'disk_drivers_icon_image.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'disk_drivers_icon_image.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2)
            self.dm_maincode.doubleClickOnImage(self.__images_dir + 'disk_drivers_icon_image.PNG', 40, 8)
    
            #TC Step 4
            assert self.dm_maincode.isImageVisible(self.__images_dir + 'disk_driver_GB_image.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'disk_driver_micron_device.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.dm_maincode.doubleClickOnImage(self.__images_dir + 'disk_driver_GB_image.PNG', -40, 8 )
            self.utils.sleepUntil(10)
            
#             assert self.dm_maincode.isWorkingProperlyImageVisisble(), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'generic_working_properly.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
        
            #TC Step 6
            self.utils.sleepUntil(2)
            self.dm_maincode.close_window()
            
            return True, "", func.co_name  

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        
        finally:
            self.dm_maincode.killProcess()
    
    
    '''TC ID: W10DA_081 '''
    def devicemanager_dsiplay_adaptor_drive_testcase_id_W10DA_081(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1
            self.logger.info("Checking and killing the process mmc.exe if any before executing the test case")
            self.dm_maincode.killProcess()
            
            #TC Step 2
            self.dm_maincode.openApplication(self.__devicemanager_binary_path)
            self.utils.sleepUntil(2)
            
            #TC Step 3
            assert self.dm_maincode.isImageVisible( self.__images_dir + 'display_adaptor_driver_image.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'display_adaptor_driver_image.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2)
            self.dm_maincode.doubleClickOnImage(self.__images_dir + 'display_adaptor_driver_image.PNG', 40, 8)
    
            #TC Step 4
            assert self.dm_maincode.isImageVisible(self.__images_dir + 'intel_adaptor_driver_image.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'intel_adaptor_driver_image.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.dm_maincode.doubleClickOnImage(self.__images_dir + 'intel_adaptor_driver_image.PNG', 40, 8 )
            self.utils.sleepUntil(10)
            
#             assert self.dm_maincode.isWorkingProperlyImageVisisble(), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'generic_working_properly.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
        
            #TC Step 6
            self.utils.sleepUntil(2)
            self.dm_maincode.close_window()
            
            return True, "", func.co_name  

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        
        finally:
            self.dm_maincode.killProcess()
    
    '''TC ID: W10DA_082 '''
    def devicemanager_imagingdevices_driver_testcase_id_W10DA_082(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1
            self.logger.info("Checking and killing the process mmc.exe if any before executing the test case")
            self.dm_maincode.killProcess()
            
            #TC Step 2
            self.dm_maincode.openApplication(self.__devicemanager_binary_path)
            self.utils.sleepUntil(2)
            
            #TC Step 3
            assert self.dm_maincode.isImageVisible( self.__images_dir + 'monitor_driver_image.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'monitor_driver_image.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2)
            self.dm_maincode.doubleClickOnImage(self.__images_dir + 'monitor_driver_image.PNG', 40, 8)
    
            #TC Step 4
            assert self.dm_maincode.isImageVisible(self.__images_dir + 'generic_monitor_driver_image.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'generic_monitor_driver_image.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.dm_maincode.doubleClickOnImage(self.__images_dir + 'generic_monitor_driver_image.PNG', 40, 8 )
            self.utils.sleepUntil(10)
            
#             assert self.dm_maincode.isWorkingProperlyImageVisisble(), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'generic_working_properly.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
        
            #TC Step 6
            self.utils.sleepUntil(2)
            self.dm_maincode.close_window()
            
            return True, "", func.co_name  

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        
        finally:
            self.dm_maincode.killProcess()

    
    
    
    '''TC ID: W10DA_083 '''
    def devicemanager_keyboard_driver_testcase_id_W10DA_083(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1
            self.logger.info("Checking and killing the process mmc.exe if any before executing the test case")
            self.dm_maincode.killProcess()
            
            #TC Step 2
            self.dm_maincode.openApplication(self.__devicemanager_binary_path)
            self.utils.sleepUntil(2)
            
            #TC Step 3
            assert self.dm_maincode.isImageVisible( self.__images_dir + 'Keyboards_driver_mage.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'Keyboards_driver_mage.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2)
            self.dm_maincode.doubleClickOnImage(self.__images_dir + 'Keyboards_driver_mage.PNG', 40, 8)
    
            #TC Step 4
            assert self.dm_maincode.isImageVisible(self.__images_dir + 'keyboard_icon_image.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'keyboard_icon_image.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.dm_maincode.doubleClickOnImage(self.__images_dir + 'keyboard_icon_image.PNG', -40, 8 )
            self.utils.sleepUntil(10)
            
#             assert self.dm_maincode.isWorkingProperlyImageVisisble(), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'generic_working_properly.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
        
            #TC Step 6
            self.utils.sleepUntil(2)
            self.dm_maincode.close_window()
            
            return True, "", func.co_name  

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        
        finally:
            self.dm_maincode.killProcess()

    
    '''TC ID: W10DA_084 '''
    def devicemanager_mouse_driver_testcase_id_W10DA_084(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1
            self.logger.info("Checking and killing the process mmc.exe if any before executing the test case")
            self.dm_maincode.killProcess()
            
            #TC Step 2
            self.dm_maincode.openApplication(self.__devicemanager_binary_path)
            self.utils.sleepUntil(2)
            
            
            #TC Step 3
            assert self.dm_maincode.isImageVisible( self.__images_dir + 'mice_driver_image.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'mice_driver_image.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            (x, y)= self.utils.getCoordinatesByLocatingGivenImageOnScreen(self.__images_dir + 'mice_driver_image.PNG', 10)
            self.utils.sleepUntil(2)
            
            self.dm_maincode.doubleClickOnImage(self.__images_dir + 'mice_driver_image.PNG', 40, 8)
            self.utils.sleepUntil(3)
            
            
            ''' since the mouse drivers are different from laptop to laptop using relative coordinates approach to click on mouse device driver '''
            x = x+40
            y= y+8
            self.utils.moveToLocationAndClick(x, y+15)
            
            self.utils.sleepUntil(10)
            
#             assert self.dm_maincode.isWorkingProperlyImageVisisble(), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'generic_working_properly.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
        
            #TC Step 6
            self.utils.sleepUntil(2)
            self.dm_maincode.close_window()
            
            return True, "", func.co_name  

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        
        finally:
            self.dm_maincode.killProcess()


    '''TC ID: W10DA_085 '''
    def devicemanager_monitor_driver_testcase_id_W10DA_085(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1
            self.logger.info("Checking and killing the process mmc.exe if any before executing the test case")
            self.dm_maincode.killProcess()
            
            #TC Step 2
            self.dm_maincode.openApplication(self.__devicemanager_binary_path)
            self.utils.sleepUntil(2)
            
            #TC Step 3
            assert self.dm_maincode.isImageVisible( self.__images_dir + 'monitor_driver_image.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'monitor_driver_image.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2)
            self.dm_maincode.doubleClickOnImage(self.__images_dir + 'monitor_driver_image.PNG', 40, 8)
    
            #TC Step 4
            assert self.dm_maincode.isImageVisible(self.__images_dir + 'generic_monitor_driver_image.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'generic_monitor_driver_image.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.dm_maincode.doubleClickOnImage(self.__images_dir + 'generic_monitor_driver_image.PNG', 40, 8 )
            self.utils.sleepUntil(10)
            
#             assert self.dm_maincode.isWorkingProperlyImageVisisble(), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'generic_working_properly.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
        
            #TC Step 6
            self.utils.sleepUntil(2)
            self.dm_maincode.close_window()
            
            return True, "", func.co_name  

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        
        finally:
            self.dm_maincode.killProcess()
    

    '''TC ID: W10DA_086 '''
    def devicemanager_network_adaptor_driver_testcase_id_W10DA_086(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1
            self.logger.info("Checking and killing the process mmc.exe if any before executing the test case")
            self.dm_maincode.killProcess()
            
            #TC Step 2
            self.dm_maincode.openApplication(self.__devicemanager_binary_path)
            self.utils.sleepUntil(2)
            
            #TC Step 3
            assert self.dm_maincode.isImageVisible( self.__images_dir + 'network_adaptors_image_icon.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'network_adaptors_image_icon.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2)
            self.dm_maincode.doubleClickOnImage(self.__images_dir + 'network_adaptors_image_icon.PNG', 40, 8)
            self.utils.sleepUntil(10)
            
            assert self.dm_maincode.isImageVisible( self.__images_dir + 'bluetooth_device_image.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'bluetooth_device_image.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2)
            self.dm_maincode.doubleClickOnImage(self.__images_dir + 'bluetooth_device_image.PNG', 40, 8)
            self.utils.sleepUntil(10)
            #TC Step 4
            
#             assert self.dm_maincode.isWorkingProperlyImageVisisble(), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'generic_working_properly.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
    
            #TC Step 6
            self.utils.sleepUntil(2)
            self.dm_maincode.close_window()
            
            return True, "", func.co_name  

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        
        finally:
            self.dm_maincode.killProcess()




    
    '''TC ID: W10DA_087 '''
    def devicemanager_processor_driver_testcase_id_W10DA_087(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            #TC Step 1
            self.logger.info("Checking and killing the process mmc.exe if any before executing the test case")
            self.dm_maincode.killProcess()
            
            #TC Step 2
            self.dm_maincode.openApplication(self.__devicemanager_binary_path)
            self.utils.sleepUntil(2)
#             (x, y)= self.utils.getCoordinatesByLocatingGivenImageOnScreen(self.__images_dir + 'mice_driver_image.PNG', 10)
#             pyautogui.click(x, y)
#             pyautogui.press('down', 10, 1)
            
            #TC Step 3
            ''' To have the visibility of Processor drivers we need to minimize network drivers'''
            assert self.dm_maincode.isImageVisible( self.__images_dir + 'network_adaptors_image_icon.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'mprocessors_driver_image.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2)
            self.dm_maincode.doubleClickOnImage(self.__images_dir + 'network_adaptors_image_icon.PNG', 40, 8)
            self.utils.sleepUntil(2)
            
            self.dm_maincode.maximize_window()
    
            assert self.dm_maincode.isImageVisible( self.__images_dir + 'processors_driver_image.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'mprocessors_driver_image.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2)
            self.dm_maincode.doubleClickOnImage(self.__images_dir + 'processors_driver_image.PNG', 40, 8)
    
            #TC Step 4
            assert self.dm_maincode.isImageVisible(self.__images_dir + 'Intelcore_processor_image_icon.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'Intelcore_processor_image_icon.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.dm_maincode.doubleClickOnImage(self.__images_dir + 'Intelcore_processor_image_icon.PNG', 40, 8 )
            self.utils.sleepUntil(10)
            
#             assert self.dm_maincode.isWorkingProperlyImageVisisble(), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'generic_working_properly.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
    
            #TC Step 6
            self.utils.sleepUntil(2)
            self.dm_maincode.close_window()
            
            return True, "", func.co_name  

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        
        finally:
            self.dm_maincode.killProcess()

    
    
    
    
    
    def deviceManager_DeviceDriverTest(self, image1, image2):
        try: 
            #TC Step 3
            assert self.dm_maincode.isImageVisible( self.__images_dir + image1 ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, image1, getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2)
            self.dm_maincode.doubleClickOnImage(self.__images_dir + image1, 40, 8)
    
            #TC Step 4
            assert self.dm_maincode.isImageVisible(self.__images_dir + image2 ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, image2, getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.dm_maincode.doubleClickOnImage(self.__images_dir + image2, 40, 8 )
            self.utils.sleepUntil(2)
            
            #TC Step 5
            try:
                assert self.dm_maincode.isWorkingProperlyImageVisisble(), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'generic_working_properly.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
                self.exceptions.append( "Exception:{} /fileName: {} / function {} /at LineNumber: {}".format(str(e), getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno ))
                self.logger.debug( "+++{} / fileName: {} / Function: {} /at LineNumber: {}".format( str(e), getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno  ))
            finally:
                self.dm_maincode.close_window()
                
                
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
#             global exceptions
            self.exceptions.append(e)
            self.logger.debug("==========>{}".format(self.exceptions))

        
#         finally:
#             #TC Step 6
#             self.logger.info("Reached the finally block to close the window")
#             self.dm_maincode.close_window()
#     
#     def decorator_func(self, func):
#         def wrapper(self, *args, **kwargs):
#             try:
#                 return func(*args, **kwargs)
#                 #TC Step 5
#                 assert self.dm_maincode.isWorkingProperlyImageVisisble(), "Not able to locate the image at location {}/{}".format(self.__images_dir,'generic_working_properly.PNG')
#                 #TC Step 6
#                 self.dm_maincode.close_window()
#             except Exception as e:
#                 global exceptions
#                 exceptions.append(e)
#                 
#         return wrapper
        
        
'''saving code if we want to test all device manager tests in one single test case and below code helps to track all exceptions together"
#     def iterateExceptions(self, exception):
#         self.logger.debug("Eception is '{}'".format(exception)) 
#         return exception       
    
        
#             try:
#                 assert self.dm_maincode.isWorkingProperlyImageVisisble(), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'generic_working_properly.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
#             except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
#                 self.exceptions.append( "Exception:{} /fileName: {} / function {} /at LineNumber: {}".format(str(e), getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno ))
#                 self.logger.debug( "+++{} / fileName: {} / Function: {} /at LineNumber: {}".format( str(e), getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno  ))
#             finally:
#                 self.dm_maincode.close_window()
                
'''                
    

    
        
        
    