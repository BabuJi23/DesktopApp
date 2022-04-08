'''
Created on Nov 06, 2018
@author: pparamasivan
'''
import inspect

from definitions import  ADMIN_TOOL_MMC_PROCESS, ADMIN_TOOL_ODBC_PROCESS
from com.gilead.main.admintool.AdminToolMainFile import AdminToolMainClass
from selenium.common.exceptions import NoSuchElementException,\
    WebDriverException


class AdminToolTestClass():
    __mmc_process = ADMIN_TOOL_MMC_PROCESS
    __odbc_process = ADMIN_TOOL_ODBC_PROCESS
    __explorer_process = "explorer.exe"
    __diagnostic_process = "MdSched.exe"
    __resource_monitor_process = "perfmon.exe"
    __msconfig_process ="msconfig.exe"
    __msinfo32_process = "msinfo32.exe"
    __powershell_process = "powershell_ise.exe"
    __defragment_process = "dfrgui.exe"
    __cleanup_process = "cleanmgr.exe"
    __iscsi_process = "iscsicpl.exe"
    

    def __init__(self, utilsRef, loggerRef):
        self.logger = loggerRef
        self.utils = utilsRef
        self.admintool_maincode = AdminToolMainClass(utilsRef, loggerRef)

    def admin_tool_testcase_id_W10DA_60_61_64_66_70_73_76(self): 
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the winium driver of Component Services app... ")
            
            # TC Step 1
            self.logger.info(
                "Checking and killing the process mmc.exe if any before executing the test case")
            self.admintool_maincode.killProcess(self.__mmc_process)

            '''TC: W10DA-60 Admin Tool - Component Services - Open > Close'''
            # TC Step 2
            self.admintool_maincode.startAdminToolApp("comexp.msc")
            self.logger.info("Started program Component Services via start menu")
            self.utils.sleepUntil(8)
              
            assert self.admintool_maincode.isGivenProcessExist("mmc.exe") , "Component Services process didn't start"
              
            '''connecting winium driver to Component Services app opened '''
            self.admintool_maincode.connectAdminToolToWebDriver()
            self.logger.info("connected Admin Tool app to Winium driver ")
              
            assert self.admintool_maincode.isElementExist(self.admintool_maincode.driver, **{'Name':'Component Services'}), "Component Services app didn't open properly"
            self.logger.info("Component Services GUI opened successfully")
              
            # TC Step 3
            self.admintool_maincode.clickAdminToolFileExitLink()
            self.logger.info("Clicked Component Services File -> Exit menu")
            self.logger.info("<==================== Component Services: Admin Tool test completed successfully. ================>")
                 
  
            '''TC: W10DA-61  Admin Tool - Computer Management - Open > Close'''      
             
            # TC Step 2
            self.admintool_maincode.startAdminToolApp("compmgmt.msc")
            self.logger.info("Started program Computer Management via start menu")
            self.utils.sleepUntil(8)
              
            assert self.admintool_maincode.isGivenProcessExist("mmc.exe") , "Computer Management process didn't start"
              
            '''connecting winium driver to Computer Management app opened '''
            self.admintool_maincode.connectAdminToolToWebDriver()
            self.logger.info("connected Admin Tool app to Winium driver ")
  
            assert self.admintool_maincode.isElementExist(self.admintool_maincode.driver, **{'Name':'Computer Management'}), "Computer Management app didn't open properly"
            self.logger.info("Computer Management GUI opened successfully")
              
            # TC Step 3
            self.admintool_maincode.clickAdminToolFileExitLink()
            self.logger.info("Clicked Computer Management File -> Exit menu")
            self.logger.info("<==================== Computer Management: Admin Tool test completed successfully. ================>")
              
            '''TC: W10DA-64 Admin Tool - Event Viewer - Open > Close'''  
            # TC Step 2
            self.admintool_maincode.startAdminToolApp("eventvwr.msc")
            self.logger.info("Started program Event Viewer via start menu")
            self.utils.sleepUntil(8)
              
            assert self.admintool_maincode.isGivenProcessExist("mmc.exe") , "Event Viewer process didn't start"
              
            '''connecting winium driver to Event Viewer app opened '''
            self.admintool_maincode.connectAdminToolToWebDriver()
            self.logger.info("connected Event Viewer app to Winium driver ")
  
            assert self.admintool_maincode.isElementExist(self.admintool_maincode.driver, **{'Name':'Event Viewer'}), "Event Viewer app didn't open properly"
            self.logger.info("Event Viewer GUI opened successfully")
              
            # TC Step 3
            self.admintool_maincode.clickAdminToolFileExitLink()
            self.logger.info("Clicked Event Viewer File -> Exit menu")           
            self.logger.info("<==================== Event Viewer: Admin Tool test completed successfully. ================>")
            
            '''TC: W10DA-66 Admin Tool - Local Security Policy - Open > Close'''
            # TC Step 2
            self.utils.sleepUntil(3)
            #self.admintool_maincode.startAdminToolApp("secpol.msc")
            self.admintool_maincode.startAdminTool_via_StartMenu("Local Security Policy")
            self.logger.info("Started program Local Security Policy via start menu")
            self.utils.sleepUntil(16)
            
            assert self.admintool_maincode.isGivenProcessExist("mmc.exe") , "Local Security Policy process didn't start"
            
            '''connecting winium driver to Local Security Policy app opened '''
            self.admintool_maincode.connectAdminToolToWebDriver()
            self.logger.info("connected Admin Tool app to Winium driver ")
            
            assert self.admintool_maincode.isElementExist(self.admintool_maincode.driver, **{'Name':'Local Security Policy'}), "Local Security Policy app didn't open properly"
            self.logger.info("Local Security Policy GUI opened successfully")
            
            # TC Step 3
            self.admintool_maincode.clickAdminToolFileExitLink()
            self.logger.info("Clicked Local Security Policy File -> Exit menu")
            self.logger.info("<==================== Local Security Policy: Admin Tool test completed successfully. ================>")
               

            '''TC: W10DA-70  Admin Tool - Print Management - Open > Close'''      
           
            # TC Step 2
            self.admintool_maincode.startAdminTool_via_StartMenu("printmanagement.msc")
            
            self.logger.info("Started program Print Management via start menu")
            self.utils.sleepUntil(16)
            
            assert self.admintool_maincode.isGivenProcessExist("mmc.exe") , "Print Management process didn't start"
            
            '''connecting winium driver to Print Management app opened '''
            self.admintool_maincode.connectAdminToolToWebDriver()
            self.logger.info("connected Admin Tool app to Winium driver ")

            assert self.admintool_maincode.isElementExist(self.admintool_maincode.driver, **{'Name':'Print Management'}), "Print Management app didn't open properly"
            self.logger.info("Print Management GUI opened successfully")
            
            # TC Step 3
            self.admintool_maincode.clickAdminToolFileExitLink()
            self.logger.info("Clicked Print Management File -> Exit menu")
            self.logger.info("<==================== Print Management: Admin Tool test completed successfully. ================>")
            
            '''TC: W10DA-73  Admin Tool - Services - Open > Close'''  
            # TC Step 2
            self.admintool_maincode.startAdminToolApp("services.msc")
            self.logger.info("Started Services via start menu")
            self.utils.sleepUntil(12)
            
            assert self.admintool_maincode.isGivenProcessExist("mmc.exe") , "Services process didn't start"
            
            '''connecting winium driver to Services app opened '''
            self.admintool_maincode.connectAdminToolToWebDriver()
            self.logger.info("connected Services app to Winium driver ")

            assert self.admintool_maincode.isElementExist(self.admintool_maincode.driver, **{'Name':'Services'}), "Services app didn't open properly"
            self.logger.info("Services GUI opened successfully")
            
            # TC Step 3
            self.admintool_maincode.clickAdminToolFileExitLink()
            self.logger.info("Clicked Services File -> Exit menu")           
            self.logger.info("<==================== Services: Admin Tool test completed successfully. ================>")
   
            '''TC: W10DA-76 Admin Tool - Task Scheduler - Open > Close''' 
            # TC Step 2
            self.admintool_maincode.startAdminToolApp("taskschd.msc")
            self.logger.info("Started Task Scheduler via start menu")
            self.utils.sleepUntil(8)
            
            assert self.admintool_maincode.isGivenProcessExist("mmc.exe") , "Task Scheduler process didn't start"
            
            '''connecting winium driver to Task Scheduler app opened '''
            self.admintool_maincode.connectAdminToolToWebDriver()
            self.logger.info("connected Task Scheduler app to Winium driver ")

            assert self.admintool_maincode.isElementExist(self.admintool_maincode.driver, **{'Name':'Task Scheduler'}), "Task Scheduler app didn't open properly"
            self.logger.info("Task Scheduler GUI opened successfully")
            
            # TC Step 3
            self.admintool_maincode.clickAdminToolFileExitLink()
            self.logger.info("Clicked Task Scheduler File -> Exit menu")           
            self.logger.info("<==================== Task Scheduler: Admin Tool test completed successfully. ================>")   
            
            # TC Step 4
            self.utils.sleepUntil(3)
#             assert self.admintool_maincode.isGivenProcessExist("mmc.exe")==False , "Admin Tool process still open start"
            self.logger.info("Closed Admin Tool application successfully" )
            
            return True, "", func.co_name 

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the winium driver process of Admin tool application.")
  
  
            '''TC: W10DA-62 Admin Tool - Defragment and Optimize Drivers - Open > Close'''       
    def admin_tool_defragment_optimize_testcase_id_W10DA_62(self): 
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the winium driver of System Information app... ")
            
            # TC Step 1
            self.logger.info("Checking and killing the process dfrgui.exe if any before executing the test case")
            self.admintool_maincode.killProcess(self.__defragment_process)

            # TC Step 2
            self.admintool_maincode.startAdminTool_via_StartMenu("dfrgui.exe")
            self.logger.info("Started program Defragment and Optimize Drivers via start menu")
            self.utils.sleepUntil(5)
              
            assert self.admintool_maincode.isGivenProcessExist("dfrgui.exe") , "Defragment and Optimize Driver process didn't start"
              
            '''connecting winium driver to Defragment and Optimize Drivers app opened '''
            self.admintool_maincode.connectAdminToolDefragmentOptimizeWebDriver()
            self.logger.info("connected Admin Tool app to Winium driver ")
              
            assert self.admintool_maincode.isElementExist(self.admintool_maincode.driver, **{'Name':'Optimize Drives'}), "Defragment and Optimize Drivers app didn't open properly"
            self.logger.info("Defragment and Optimize Drivers GUI opened successfully")
              
            # TC Step 3
            self.admintool_maincode.clickAdminToolCloseButton()
            self.logger.info("Clicked Defragment and Optimize Drivers  close Button")
            self.logger.info("<====================  Admin Tool test Defragment and Optimize Drivers completed successfully. ================>")
  
            # TC Step 4
            self.utils.sleepUntil(4)
#             assert self.admintool_maincode.isGivenProcessExist("dfrgui.exe")==False , "Admin Tool Defragment and Optimize Drivers process still open"
            self.logger.info("Closed Admin Tools application successfully" )
            
            return True, "", func.co_name 

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the winium driver process of Admin tool application.")
            self.admintool_maincode.killProcess(self.__defragment_process)
                             
  
            '''TC: W10DA-63 Admin Tool - Disk Cleanup - Open > Close'''       
    def admin_tool_disk_cleanup_testcase_id_W10DA_63(self): 
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the winium driver of System Information app... ")
            
            # TC Step 1
            self.logger.info("Checking and killing the process cleanmgr.exe if any before executing the test case")
            self.admintool_maincode.killProcess(self.__cleanup_process)

            # TC Step 2
            self.admintool_maincode.startAdminTool_via_StartMenu("cleanmgr.exe")
            self.logger.info("Started program Disk Cleanup via start menu")
            self.utils.sleepUntil(15)
              
            assert self.admintool_maincode.isGivenProcessExist("cleanmgr.exe") , "Disk Cleanup process didn't start"
              
            '''connecting winium driver to Disk Cleanup app opened '''
            self.admintool_maincode.connectAdminToolDiskCleanupWebDriver()
            self.logger.info("connected Admin Tool app to Winium driver ")
              
#             assert self.admintool_maincode.isElementExist(self.admintool_maincode.driver, **{'Name':'Disk Cleanup for OS (C:)'}), "Disk Cleanup app didn't open properly"
            self.utils.sleepUntil(10)
            self.logger.info("Disk Cleanup GUI opened successfully")
              
            # TC Step 3
            self.admintool_maincode.clickAdminToolDiskCleanupCancelButton()
            self.logger.info("Clicked Disk Cleanup cancel Button")
            self.logger.info("<====================  Admin Tool test Disk Cleanup completed successfully. ================>")
  
            # TC Step 4
#             assert self.admintool_maincode.isGivenProcessExist("cleanmgr.exe")==False , "Admin Tool Disk Cleanup process still open"
            self.logger.info("Closed Admin Tools application successfully" )
            
            return True, "", func.co_name 

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the winium driver process of Admin tool application.")
            self.admintool_maincode.killProcess(self.__cleanup_process)
  
            '''TC: W10DA-65 Admin Tool -  iSCSi  - Open > Close'''       
    def admin_tool_iSCSi_testcase_id_W10DA_65(self): 
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the winium driver of System Information app... ")
            
            # TC Step 1
            self.logger.info("Checking and killing the process iscsicpl.exe if any before executing the test case")
            self.admintool_maincode.killProcess(self.__iscsi_process)

            # TC Step 2
            self.admintool_maincode.startAdminTool_via_StartMenu("iscsicpl.exe")
            self.logger.info("Started program iSCSi via start menu")
            self.utils.sleepUntil(5)
              
            assert self.admintool_maincode.isGivenProcessExist("iscsicpl.exe") , "iSCSi process didn't start"
              
            '''connecting winium driver to iSCSi app opened '''
            self.admintool_maincode.connectAdminTooliSCSiWebDriver()
            self.logger.info("connected Admin Tool app to Winium driver ")
              
            assert self.admintool_maincode.isElementExist(self.admintool_maincode.driver, **{'Name':'iSCSI Initiator Properties'}), "iSCSi app didn't open properly"
            self.utils.sleepUntil(10)
            self.logger.info("iSCSi GUI opened successfully")
              
            # TC Step 3
            self.admintool_maincode.clickAdminToolOkButton()
            self.logger.info("Clicked iSCSi OK Button")
            self.logger.info("<====================  Admin Tool test iSCSi completed successfully. ================>")
  
            # TC Step 4
#             assert self.admintool_maincode.isGivenProcessExist("iscsicpl.exe")==False , "Admin Tool iSCSi process still open"
            self.logger.info("Closed Admin Tools application successfully" )
            
            return True, "", func.co_name 

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the winium driver process of Admin tool application.")
            self.admintool_maincode.killProcess(self.__iscsi_process)

    def admin_tool_odbcad32_testcase_id_W10DA_67_68(self): 
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the winium driver of ODBC data Services app... ")
            
            # TC Step 1
            self.logger.info(
                "Checking and killing the process mmc.exe if any before executing the test case")
            self.admintool_maincode.killProcess(self.__odbc_process)

            '''TC: W10DA-67 Admin Tool - ODBC (32-bit) - Open > Close'''
            # TC Step 2
            self.admintool_maincode.startAdminToolApp("odbcad32.exe")
            self.logger.info("Started program ODBC Data Sources (32-bit) via start menu")
            self.utils.sleepUntil(10)
              
            assert self.admintool_maincode.isGivenProcessExist("odbcad32.exe") , "ODBC Data Sources (32-bit) process didn't start"
              
            '''connecting winium driver to ODBC Data Sources (32-bit) app opened '''
            self.admintool_maincode.connectAdminToolOdbcToWebDriver()
            self.logger.info("connected Admin Tool app to Winium driver ")
              
            assert self.admintool_maincode.isElementExist(self.admintool_maincode.driver, **{'Name':'ODBC Data Source Administrator (32-bit)'}), "ODBC Data Sources (32-bit) app didn't open properly"
            self.logger.info("ODBC Data Sources (32-bit) GUI opened successfully")
              
            # TC Step 3
            #self.admintool_maincode.clickAdminToolOkButton()
            self.admintool_maincode.killProcess(self.__odbc_process)
            self.logger.info("Clicked ODBC Data Sources (32-bit)  close Button")
            self.logger.info("<==================== ODBC Data Sources (32-bit): Admin Tool test completed successfully. ================>")
 
            '''TC: W10DA-68 Admin Tool - ODBC (64-bit) - Open > Close'''
            # TC Step 2
            self.admintool_maincode.startAdminTool_via_StartMenu("ODBC Data Sources (64-bit)")
            self.logger.info("Started program ODBC Data Sources (64-bit) via start menu")
            self.utils.sleepUntil(12)
              
            assert self.admintool_maincode.isGivenProcessExist("odbcad32.exe") , "ODBC Data Sources (64-bit) process didn't start"
              
            '''connecting winium driver to ODBC Data Sources (64-bit) app opened '''
            self.admintool_maincode.connectAdminToolOdbcToWebDriver()
            self.logger.info("connected Admin Tool app to Winium driver ")
              
            assert self.admintool_maincode.isElementExist(self.admintool_maincode.driver, **{'Name':'ODBC Data Source Administrator (64-bit)'}), "ODBC Data Sources (64-bit) app didn't open properly"
            self.logger.info("ODBC Data Sources (64-bit) GUI opened successfully")
              
            # TC Step 3
            #self.admintool_maincode.clickAdminToolOkButton()
            self.admintool_maincode.killProcess(self.__odbc_process)
            self.logger.info("Clicked ODBC Data Sources (64-bit)  Close Button")
            self.logger.info("<==================== ODBC Data Sources (64-bit): Admin Tool test completed successfully. ================>")
                 
            # TC Step 4
            self.utils.sleepUntil(10)
#             assert self.admintool_maincode.isGivenProcessExist("odbcad32.exe")==False , "Admin Tool process still open start"
            self.logger.info("Closed Admin Tool application successfully" )
            
            return True, "", func.co_name 

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the winium driver process of Admin tool application.")
            self.admintool_maincode.killProcess(self.__odbc_process)
            
        '''TC: W10DA-69 Admin Tool - Memory Diagnostic  - Open > Close'''                 
    def admin_tool_memory_diagnostic_testcase_id_W10DA_69(self): 
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the winium driver of Memory Diagnostic app... ")
            
            # TC Step 1
            self.logger.info(
                "Checking and killing the process mmc.exe if any before executing the test case")
            self.admintool_maincode.killProcess(self.__diagnostic_process)


            # TC Step 2
            self.admintool_maincode.startAdminTool_via_StartMenu("Memory Diagnostic")
            self.logger.info("Started program Memory Diagnostic via start menu")
            self.utils.sleepUntil(8)
              
            assert self.admintool_maincode.isGivenProcessExist("MdSched.exe") , "Memory Diagnostic process didn't start"
              
            '''connecting winium driver to Memory Diagnostic app opened '''
            self.admintool_maincode.connectAdminToolMemoryDiagnosticWebDriver()
            self.logger.info("connected Admin Tool app to Winium driver ")
              
            assert self.admintool_maincode.isElementExist(self.admintool_maincode.driver, **{'Name':'Windows Memory Diagnostic'}), "Memory Diagnostic app didn't open properly"
            self.logger.info("Memory Diagnostic GUI opened successfully")
              
            # TC Step 3
            self.admintool_maincode.clickAdminToolMemoryDiagnosticCancelButton()
            self.logger.info("Clicked Memory Diagnostic  close Button")
            self.logger.info("<==================== Memory Diagnostic: Admin Tool test Memory Diagnostic completed successfully. ================>")

            # TC Step 4
            self.utils.sleepUntil(10)
#             assert self.admintool_maincode.isGivenProcessExist("MdSched.exe")==False , "Admin Tool Memory Diagnostic process still open"
            self.logger.info("Closed Admin Tool Memory Diagnostic application successfully" )
            
            return True, "", func.co_name 

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the winium driver process of Admin tool application.")    
            self.admintool_maincode.killProcess(self.__diagnostic_process) 
            
            '''TC: W10DA-72 Admin Tool -Resource Monitor - Open > Close'''
    def admin_tool_resource_monitor_testcase_id_W10DA_72(self): 
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the winium driver of Resource monitor app... ")
            
            # TC Step 1
            self.logger.info("Checking and killing the process perfmon.exe if any before executing the test case")
            self.admintool_maincode.killProcess(self.__resource_monitor_process)
            
            # TC Step 2
            self.admintool_maincode.startAdminTool_via_StartMenu("Resource Monitor")
            self.logger.info("Started program Resource monitor via start menu")
            self.utils.sleepUntil(8)
        
            assert self.admintool_maincode.isGivenProcessExist("perfmon.exe") , "Resource monitor process didn't start"
            
            '''connecting winium driver to Resource monitor app opened '''
            self.admintool_maincode.connectAdminToolResourceMonitorWebDriver()
            self.logger.info("connected Resource monitor app to Winium driver ")

            # TC Step 3
            assert self.admintool_maincode.isElementExist(self.admintool_maincode.driver, **{'Name':'Resource Monitor'}), "Performance monitor app didn't open properly"
            self.logger.info("Resource monitor GUI opened successfully")
            
            self.admintool_maincode.closeAdminToolResourceMonitor()
            self.logger.info("Closed Resource monitor app successfully" )
            
            # TC Step 4
            self.utils.sleepUntil(10)
#             assert self.admintool_maincode.isGivenProcessExist("perfmon.exe")==False , "Resource monitor process still open start"
            
            return True, "", func.co_name 
            
        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the winium driver process of Resource monitor app")   
            self.admintool_maincode.killProcess(self.__resource_monitor_process)
            
  
    def admin_tool_reliability_monitor_Firewall_testcase_id_W10DA_71_77(self): 
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the winium driver of Reliability monitor app... ")
            
            # TC Step 1
            ''' Explorer.exe  is part of operating system always running for windows explorer. This validation is not helpful'''
#             self.logger.info("Checking and killing the process explorer.exe if any before executing the test case")
#             self.admintool_maincode.killProcess(self.__explorer_process)

            '''TC: W10DA-71 Admin Tool - Reliability Monitor - Open > Close'''
            # TC Step 2
            self.admintool_maincode.startAdminTool_via_StartMenu("reliability history")
            self.logger.info("Started program Reliability monitor via start menu")
            self.utils.sleepUntil(8)
              
            assert self.admintool_maincode.isGivenProcessExist("explorer.exe") , "Reliability Monitor explorer process didn't start"
              
            '''connecting winium driver to Reliability monitor app opened '''
            self.admintool_maincode.connectAdminToolExplorerWebDriver()
            self.logger.info("connected Admin Tool app to Winium driver ")
              
            assert self.admintool_maincode.isElementExist(self.admintool_maincode.driver, **{'Name':'Reliability Monitor'}), "Reliability monitor app didn't open properly"
            self.logger.info("Reliability monitor GUI opened successfully")
              
            # TC Step 3
            #self.utils.sleepUntil(3)
            self.admintool_maincode.clickAdminToolReliabilityMonitorCloseButton()
            self.logger.info("Clicked Reliability monitor  close Button")
            self.logger.info("<==================== Reliability monitor: Admin Tool test Reliability monitor completed successfully. ================>")
  
            '''TC: W10DA-77 Admin Tool - Windows Defender Firewall - Open > Close'''
            # TC Step 2
            self.utils.sleepUntil(10)
            self.admintool_maincode.startAdminTool_via_StartMenu("Windows Defender Firewall")
            self.logger.info("Started program Windows Defender Firewall via start menu")
            self.utils.sleepUntil(10)
               
            assert self.admintool_maincode.isGivenProcessExist("explorer.exe") , "Windows Defender Firewall explorer process didn't start"
               
            '''connecting winium driver to Windows Defender Firewall app opened '''
            self.admintool_maincode.connectAdminToolExplorerWebDriver()
            self.logger.info("connected Admin Tool app to Winium driver ")
               
            assert self.admintool_maincode.isElementExist(self.admintool_maincode.driver, **{'Name':'Windows Defender Firewall'}), "Windows Defender Firewall app didn't open properly"
            self.logger.info("Windows Defender Firewall GUI opened successfully")
               
            # TC Step 3
            #self.utils.sleepUntil(3)
            self.admintool_maincode.clickAdminToolWndowsFirewallCloseButton()
            self.logger.info("Clicked Windows Defender Firewall  close Button")
            self.logger.info("<==================== Reliability monitor: Admin Tool test Windows Defender Firewall completed successfully. ================>")

            # TC Step 4
            ''' Explorer.exe is part of operating system always running for windows explorer. After killing explorer.exe process, it automatically restarts. This validation is not helpful'''
#             if self.admintool_maincode.isGivenProcessExist("explorer.exe"):
#                 self.admintool_maincode.killProcess("explorer.exe")   

#             assert self.admintool_maincode.isGivenProcessExist("explorer.exe")==False , "Admin Tool Reliability monitor process still open"
            self.logger.info("Closed Admin Tools application successfully" )
            
            return True, "", func.co_name 

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the winium driver process of Admin tool application.")
            self.admintool_maincode.killProcess("explorer.exe")
            
            
        '''TC: W10DA-74 Admin Tool - System Configuration - Open > Close'''       
    def admin_tool_system_configuration_testcase_id_W10DA_74(self): 
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the winium driver of System Configuration app... ")
            
            # TC Step 1
            self.logger.info("Checking and killing the process msconfig.exe if any before executing the test case")
            self.admintool_maincode.killProcess(self.__msconfig_process)

            # TC Step 2
            self.admintool_maincode.startAdminTool_via_StartMenu("System Configuration")
            self.logger.info("Started program System Configuration via start menu")
            self.utils.sleepUntil(8)
              
            assert self.admintool_maincode.isGivenProcessExist("msconfig.exe") , "System Configuration process didn't start"
              
            '''connecting winium driver to System Configuration app opened '''
            self.admintool_maincode.connectAdminToolSystemConfigurationWebDriver()
            self.logger.info("connected Admin Tool app to Winium driver ")
              
            assert self.admintool_maincode.isElementExist(self.admintool_maincode.driver, **{'Name':'System Configuration'}), "System Configuration app didn't open properly"
            self.logger.info("System Configuration GUI opened successfully")
              
            # TC Step 3
            #self.utils.sleepUntil(3)
            self.admintool_maincode.clickAdminToolOkButton()
            self.logger.info("Clicked System Configuration  close Button")
            self.logger.info("<==================== System Configuration: Admin Tool test System Configuration completed successfully. ================>")
  
            # TC Step 4
            assert self.admintool_maincode.isGivenProcessExist("msconfig.exe")==False , "Admin Tool System Configuration process still open"
            self.logger.info("Closed Admin Tools application successfully" )
            
            return True, "", func.co_name 

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the winium driver process of Admin tool application.")
            self.admintool_maincode.killProcess(self.__msconfig_process)
                     
                     
        '''TC: W10DA-75 Admin Tool - System Information - Open > Close'''       
    def admin_tool_system_information_testcase_id_W10DA_75(self): 
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the winium driver of System Information app... ")
            
            # TC Step 1
            self.logger.info("Checking and killing the process msinfo32.exe if any before executing the test case")
            self.admintool_maincode.killProcess(self.__msinfo32_process)

            # TC Step 2
            self.admintool_maincode.startAdminTool_via_StartMenu("System Information")
            self.logger.info("Started program System Information via start menu")
            self.utils.sleepUntil(8)
              
            assert self.admintool_maincode.isGivenProcessExist("msinfo32.exe") , "System Information process didn't start"
              
            '''connecting winium driver to System Information app opened '''
            self.admintool_maincode.connectAdminToolWindowsPowershellWebDriver()
            self.logger.info("connected Admin Tool app to Winium driver ")
              
            assert self.admintool_maincode.isElementExist(self.admintool_maincode.driver, **{'Name':'System Information'}), "System Information app didn't open properly"
            self.logger.info("System Information GUI opened successfully")
              
            # TC Step 3
            #self.utils.sleepUntil(3)
            self.admintool_maincode.clickAdminToolSystemInfoFileExitLink()
            self.logger.info("Clicked System Information  close Button")
            self.logger.info("<==================== System Information: Admin Tool test System Information completed successfully. ================>")
  
            # TC Step 4
#             assert self.admintool_maincode.isGivenProcessExist("msinfo32.exe")==False , "Admin Tool System Information process still open"
            self.logger.info("Closed Admin Tools application successfully" )
            
            return True, "", func.co_name 

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the winium driver process of Admin tool application.")
            self.admintool_maincode.killProcess(self.__msinfo32_process)
            
                             
            '''TC: W10DA-78 Admin Tool - Windows PowerShell ISE - Open > Close'''       
    def admin_tool_windows_powerShell_testcase_id_W10DA_78(self): 
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            
            '''start webdriver '''
            pid = self.utils.startWiniumDesktopDriver()
            self.logger.info("Launched the winium driver of System Information app... ")
            
            # TC Step 1
            self.logger.info("Checking and killing the process powershell_ise.exe if any before executing the test case")
            self.admintool_maincode.killProcess(self.__powershell_process)

            # TC Step 2
            self.admintool_maincode.startAdminToolApp("powershell_ise.exe")
            self.logger.info("Started program Windows PowerShell via start menu")
            self.utils.sleepUntil(8)
              
            assert self.admintool_maincode.isGivenProcessExist("powershell_ise.exe") , "Windows PowerShell process didn't start"
              
            '''connecting winium driver to Windows PowerShell app opened '''
            self.admintool_maincode.connectAdminToolSystemConfigurationWebDriver()
            self.logger.info("connected Admin Tool app to Winium driver ")
              
            assert self.admintool_maincode.isElementExist(self.admintool_maincode.driver, **{'Name':'Administrator: Windows PowerShell ISE (x86)'}), "Windows PowerShell app didn't open properly"
            self.logger.info("Windows PowerShell GUI opened successfully")
              
            # TC Step 3
            #self.utils.sleepUntil(3)
            self.admintool_maincode.clickAdminToolPowerShellCloseButton()
            self.logger.info("Clicked Windows PowerShell  close Button")
            self.logger.info("<==================== Windows PowerShell: Admin Tool test Windows PowerShell completed successfully. ================>")
  
            # TC Step 4
            if self.admintool_maincode.isGivenProcessExist("powershell_ise.exe"):
                self.admintool_maincode.killProcess("powershell_ise.exe")   

#             assert self.admintool_maincode.isGivenProcessExist("powershell_ise.exe")==False , "Admin Tool Windows PowerShell process still open"
            self.logger.info("Closed Admin Tools application successfully" )
            
            return True, "", func.co_name 

        except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
            return False, e, func.co_name
        finally:
            '''kill the web driver '''
            pid.terminate()
            self.logger.info("Killed the winium driver process of Admin tool application.")
                             
                                    
            
