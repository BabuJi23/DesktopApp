'''
Created on April 25, 2019
@author: ngarimella
'''
import inspect

from definitions import GOOGLE_LEGACY_BROWSER_FILE, PERFORMANCE_MONITOR_PROCESS, SYSTEM_IMAGES_DIR, ROOT_DIR
from com.gilead.main.system.SystemsettingsMainFile import SystemsettingsMainClass
from selenium.common.exceptions import NoSuchElementException,\
    WebDriverException
from inspect import currentframe, getframeinfo


class SystemsettingsTestClass():
    __images_dir = ROOT_DIR+SYSTEM_IMAGES_DIR
    
    def __init__(self, utilsRef, loggerRef):
        self.logger = loggerRef
        self.utils = utilsRef
        self.systemsettings_maincode = SystemsettingsMainClass(utilsRef, loggerRef)

    
    '''TC: W10DA-103 '''
    def system_settings_performanceframe_testcase_id_W10DA_103(self): 
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            # TC Step 1
            # self.logger.info(
            #     "Checking and killing the process mmc.exe if any before executing the test case")
            # self.miscellaneous_maincode.killProcess(self.__performance_monitor_process)
            
            # TC Step 2
            self.systemsettings_maincode.startProgramByName_via_StartMenu("System Settings")
            self.logger.info("Started system settings via start menu")
            self.utils.sleepUntil(5)
            
            # TC Step 3
            assert self.systemsettings_maincode.isImageVisible(self.__images_dir + 'performance_settings.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'performance_settings.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.systemsettings_maincode.doubleClickOnImage(self.__images_dir + 'performance_settings.PNG', 40, 8 )
            self.utils.sleepUntil(5)
            
            # TC Step 4
            try:
                assert self.systemsettings_maincode.isWorkingProperlyImageVisisble(1), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'windows_checking.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
                self.exceptions.append( "Exception:{} /fileName: {} / function {} /at LineNumber: {}".format(str(e), getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno ))
                self.logger.debug( "+++{} / fileName: {} / Function: {} /at LineNumber: {}".format( str(e), getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno  ))
            finally:
                self.systemsettings_maincode.close_window()
                self.utils.sleepUntil(2)
                self.systemsettings_maincode.close_window()

            return True, "", func.co_name 
            
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name


    '''TC: W10DA-105 '''
    def system_settings_remoteusers_testcase_id_W10DA_105(self): 
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
                    
            # TC Step 1
            # self.logger.info(
            #     "Checking and killing the process mmc.exe if any before executing the test case")
            # self.miscellaneous_maincode.killProcess(self.__performance_monitor_process)
            
            # TC Step 2
            self.systemsettings_maincode.startProgramByName_via_StartMenu("System Settings")
            self.logger.info("Started system settings via start menu")
            self.utils.sleepUntil(5)
          
            # TC Step 3
            assert self.systemsettings_maincode.isImageVisible(self.__images_dir + 'remote_setings.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'remote_setings.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.systemsettings_maincode.doubleClickOnImage(self.__images_dir + 'remote_setings.PNG', 40, 8 )
            self.utils.sleepUntil(5)

            # TC Step 4
            try:
                assert self.systemsettings_maincode.isWorkingProperlyImageVisisble(2), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'allow_remoteconnections_computer.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
                self.exceptions.append( "Exception:{} /fileName: {} / function {} /at LineNumber: {}".format(str(e), getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno ))
                self.logger.debug( "+++{} / fileName: {} / Function: {} /at LineNumber: {}".format( str(e), getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno  ))
            self.logger.info("Allow remote connections to this computer check is enabled")

            # TC Step 5
            try:
                assert self.systemsettings_maincode.isWorkingProperlyImageVisisble(3), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'allow_connections_enabled_check.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
                self.exceptions.append( "Exception:{} /fileName: {} / function {} /at LineNumber: {}".format(str(e), getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno ))
                self.logger.debug( "+++{} / fileName: {} / Function: {} /at LineNumber: {}".format( str(e), getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno  ))
            self.logger.info("Allow connections only from computers running Desktop with Network Level Authenticationis enabled")
            
            # TC Step 6
            assert self.systemsettings_maincode.isImageVisible(self.__images_dir + 'remote_users.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'remote_setings.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.systemsettings_maincode.doubleClickOnImage(self.__images_dir + 'remote_users.PNG', 40, 8 )
            self.utils.sleepUntil(5)
           
            # TC Step 7
            try:
                assert self.systemsettings_maincode.isWorkingProperlyImageVisisble(4), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'remote_users_existing_check.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
                self.exceptions.append( "Exception:{} /fileName: {} / function {} /at LineNumber: {}".format(str(e), getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno ))
                self.logger.debug( "+++{} / fileName: {} / Function: {} /at LineNumber: {}".format( str(e), getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno  ))
            finally:
                self.systemsettings_maincode.close_window()
                self.utils.sleepUntil(2) 
                self.systemsettings_maincode.close_window()
            return True, "", func.co_name 
            
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name


    '''TC: W10DA-104 '''
    def system_settings_system_failure_testcase_id_W10DA_104(self): 
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
     
            # TC Step 1
            # self.logger.info(
            #     "Checking and killing the process mmc.exe if any before executing the test case")
            # self.miscellaneous_maincode.killProcess(self.__performance_monitor_process)
            
            # TC Step 2
            self.systemsettings_maincode.startProgramByName_via_StartMenu("System Settings")
            self.logger.info("Started system settings via start menu")
            self.utils.sleepUntil(5)
            
            # TC Step 3
            assert self.systemsettings_maincode.isImageVisible(self.__images_dir + 'startup_recovery_info.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'startup_recovery_info.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.systemsettings_maincode.doubleClickOnImage(self.__images_dir + 'startup_recovery_settings.PNG', 40, 8 )
            self.utils.sleepUntil(5)
            
            # TC Step 4
            try:
                assert self.systemsettings_maincode.isWorkingProperlyImageVisisble(5), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'sytemlog_check.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
                self.exceptions.append( "Exception:{} /fileName: {} / function {} /at LineNumber: {}".format(str(e), getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno ))
                self.logger.debug( "+++{} / fileName: {} / Function: {} /at LineNumber: {}".format( str(e), getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno  ))
            self.logger.info("Write an event to the system log is enabled")
            self.utils.sleepUntil(2)
            
            # TC Step 5
            try:
                assert self.systemsettings_maincode.isWorkingProperlyImageVisisble(6), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'automatically_restart_check.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
                self.exceptions.append( "Exception:{} /fileName: {} / function {} /at LineNumber: {}".format(str(e), getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno ))
                self.logger.debug( "+++{} / fileName: {} / Function: {} /at LineNumber: {}".format( str(e), getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno  ))
            finally:
                self.systemsettings_maincode.close_window()
                self.utils.sleepUntil(2)
                self.systemsettings_maincode.close_window()
            self.utils.sleepUntil(2)

            return True, "", func.co_name 
            
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name


    '''TC: W10DA-106 '''
    def system_settings_system_protection_testcase_id_W10DA_106(self): 
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            # TC Step 1
            # self.logger.info(
            #     "Checking and killing the process mmc.exe if any before executing the test case")
            # self.miscellaneous_maincode.killProcess(self.__performance_monitor_process)
            
            # TC Step 2
            self.systemsettings_maincode.startProgramByName_via_StartMenu("System Settings")
            self.logger.info("Started system settings via start menu")
            self.utils.sleepUntil(5)
            
            # TC Step 3
            assert self.systemsettings_maincode.isImageVisible(self.__images_dir + 'system_protection_settings.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'system_protection_settings.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.systemsettings_maincode.doubleClickOnImage(self.__images_dir + 'system_protection_settings.PNG', 40, 8 )
            self.utils.sleepUntil(5)
            
            # TC Step 4
            try:
                assert self.systemsettings_maincode.isWorkingProperlyImageVisisble(7), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'drive_protection_check.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
                self.exceptions.append( "Exception:{} /fileName: {} / function {} /at LineNumber: {}".format(str(e), getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno ))
                self.logger.debug( "+++{} / fileName: {} / Function: {} /at LineNumber: {}".format( str(e), getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno  ))
            finally:
                self.systemsettings_maincode.close_window()
                self.utils.sleepUntil(2)

            return True, "", func.co_name 
            
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name


    '''TC: W10DA-107 '''
    def windows_settings_power_sleep_testcase_id_W10DA_107(self): 
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
         
            # TC Step 1
            # self.logger.info(
            #     "Checking and killing the process mmc.exe if any before executing the test case")
            # self.miscellaneous_maincode.killProcess(self.__performance_monitor_process)
            
            # TC Step 2
            self.systemsettings_maincode.startProgramByName_via_StartMenu("Settings")
            self.logger.info("Started windows settings via start menu")
            self.utils.sleepUntil(5)
            
            # TC Step 3
            assert self.systemsettings_maincode.isImageVisible(self.__images_dir + 'windows_settings_system.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'windows_settings_system.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.systemsettings_maincode.doubleClickOnImage(self.__images_dir + 'windows_settings_system.PNG', 40, 8 )
            self.utils.sleepUntil(5)

            # TC Step 4
            assert self.systemsettings_maincode.isImageVisible(self.__images_dir + 'power_sleep.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'power_sleep.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.systemsettings_maincode.doubleClickOnImage(self.__images_dir + 'power_sleep.PNG', 40, 8 )
            self.utils.sleepUntil(5)
            
            # TC Step 5
            try:
                assert self.systemsettings_maincode.isWorkingProperlyImageVisisble(8), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'screen_power.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
                assert self.systemsettings_maincode.isWorkingProperlyImageVisisble(9), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'screen_plugin.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
                assert self.systemsettings_maincode.isWorkingProperlyImageVisisble(10), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'sleep_power.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
                assert self.systemsettings_maincode.isWorkingProperlyImageVisisble(11), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'sleep_plugin.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
                self.exceptions.append( "Exception:{} /fileName: {} / function {} /at LineNumber: {}".format(str(e), getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno ))
                self.logger.debug( "+++{} / fileName: {} / Function: {} /at LineNumber: {}".format( str(e), getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno  ))
            finally:
                self.systemsettings_maincode.close_window()
                self.utils.sleepUntil(2)

            return True, "", func.co_name 
            
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name

    '''TC: W10DA-108 '''
    def windows_settings_date_time_testcase_id_W10DA_108(self): 
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
          
            # TC Step 1
            # self.logger.info(
            #     "Checking and killing the process mmc.exe if any before executing the test case")
            # self.miscellaneous_maincode.killProcess(self.__performance_monitor_process)
            
            # TC Step 2
            self.systemsettings_maincode.startProgramByName_via_StartMenu("Settings")
            self.logger.info("Started windows settings via start menu")
            self.utils.sleepUntil(5)

            #Added by Joseph 3/31/2020
            #self.systemsettings_maincode.doubleClickOnImage(self.__images_dir + 'settings_home.PNG', 10,10)
            #self.utils.sleepUntil(2)

            # TC Step 3
            assert self.systemsettings_maincode.isImageVisible(self.__images_dir + 'time_language.PNG' ), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir, 'time_language.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
            self.utils.sleepUntil(2) 
            self.systemsettings_maincode.doubleClickOnImage(self.__images_dir + 'time_language.PNG', 40, 8 )
            self.utils.sleepUntil(5)
            
            # TC Step 4
            try:
                assert self.systemsettings_maincode.isWorkingProperlyImageVisisble(12), "Not able to locate the image at location {}{} / fileName: {} / function {} /at LineNumber: {}".format(self.__images_dir,'timezone_check.PNG', getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno )
                self.utils.sleepUntil(5)
            except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
                self.exceptions.append( "Exception:{} /fileName: {} / function {} /at LineNumber: {}".format(str(e), getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno ))
                self.logger.debug( "+++{} / fileName: {} / Function: {} /at LineNumber: {}".format( str(e), getframeinfo(currentframe()).filename, getframeinfo(currentframe()).function, getframeinfo(currentframe()).lineno  ))
            finally:
                self.systemsettings_maincode.close_window()
                self.utils.sleepUntil(2)

            return True, "", func.co_name 
            
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name




