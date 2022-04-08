'''
Created on Nov 05, 2018

@author: pparamasivan
'''
from selenium import webdriver

class AdminToolMainClass():
    
    def __init__(self, utilsRef, loggerRef):
        self.logger = loggerRef
        self.utils = utilsRef
        self.driver =""
    
    def killProcess(self, process):
        self.utils.killProcess(process)
        
    def isGivenProcessExist(self, process):
        return self.utils.isProcessExist(process)
    
    def startAdminToolApp(self, appName):
        self.utils.openApplication(appName) 
        
    def startAdminTool_via_StartMenu(self, SearchText):
        self.utils.startProgramByName_via_StartMenu(SearchText)
        
    def connectAdminToolToWebDriver(self):
        self.driver = webdriver.Remote(command_executor="http://localhost:9999",
                                       desired_capabilities={
                                           'app': r'C:\Windows\System32\mmc.exe',
                                           'debugConnectToRunningApp': 'true',
                                           "args": "-port 345"
                                        })
    def isElementExist(self, parent_element, **element_attribute_text):
        return self.utils.isElementPresent(parent_element, **element_attribute_text)
    
    def clickAdminToolFileExitLink(self):
        parent_window_element = self.driver.find_element_by_class_name("MMCMainFrame")
        parent_window_element.find_element_by_name("File").click()
        self.utils.sleepUntil(3)
        self.driver.find_element_by_name("Exit").click()
    
    def connectAdminToolOdbcToWebDriver(self):
        self.driver = webdriver.Remote(command_executor="http://localhost:9999",
                                       desired_capabilities={
                                           'app': r'C:\Windows\System32\odbcad32.exe',
                                           'debugConnectToRunningApp': 'true',
                                           "args": "-port 345"
                                        })
    
    def clickAdminToolOkButton(self):
        parent__element = self.driver.find_element_by_class_name("#32770")
        parent__element.find_element_by_name("OK").click()
  
            
    def connectAdminToolExplorerWebDriver(self):
        self.driver = webdriver.Remote(command_executor="http://localhost:9999",
                                       desired_capabilities={
                                           'app': r'C:\Windows\System32\explorer.exe',
                                           'debugConnectToRunningApp': 'true',
                                           "args": "-port 345"
                                        })
    
    def clickAdminToolReliabilityMonitorCloseButton(self):
        parent__element = self.driver.find_element_by_name("Reliability Monitor")
        parent__element.find_element_by_name("Close").click()        
        
    def clickAdminToolWndowsFirewallCloseButton(self):
        parent__element = self.driver.find_element_by_name("Windows Defender Firewall")
        parent__element.find_element_by_name("Close").click()     
        
        
    def connectAdminToolMemoryDiagnosticWebDriver(self):
        self.driver = webdriver.Remote(command_executor="http://localhost:9999",
                                       desired_capabilities={
                                           'app': r'C:\Windows\System32\MdSched.exe',
                                           'debugConnectToRunningApp': 'true',
                                           "args": "-port 345"
                                        })    
        
         
    def clickAdminToolMemoryDiagnosticCancelButton(self):
        parent_element = self.driver.find_element_by_class_name("#32770")
        child_element = parent_element.find_element_by_name("Windows Memory Diagnostic")
        child_element.find_element_by_name("Cancel").click()  
        
    def connectAdminToolResourceMonitorWebDriver(self):
        self.driver = webdriver.Remote(command_executor="http://localhost:9999",
                                       desired_capabilities={
                                           'app': r'C:\Windows\System32\perfmon.exe',
                                           'debugConnectToRunningApp': 'true',
                                           "args": "-port 345"
                                        })
        

    def closeAdminToolResourceMonitor(self):
        parent_element = self.driver.find_element_by_name("Resource Monitor")
        child_element = parent_element.find_element_by_name("Application")
        child_element.find_element_by_name("File").click()
        self.utils.sleepUntil(1)
        self.driver.find_element_by_name("Exit").click()
        
    def connectAdminToolSystemConfigurationWebDriver(self):
        self.driver = webdriver.Remote(command_executor="http://localhost:9999",
                                       desired_capabilities={
                                           'app': r'C:\Windows\System32\msconfig.exe',
                                           'debugConnectToRunningApp': 'true',
                                           "args": "-port 345"
                                        })
             
    def clickAdminToolCloseButton(self):
        parent__element = self.driver.find_element_by_class_name("#32770")
        parent__element.find_element_by_name("Close").click() 
               

    def connectAdminToolSystemInformationWebDriver(self):
        self.driver = webdriver.Remote(command_executor="http://localhost:9999",
                                       desired_capabilities={
                                           'app': r'C:\Windows\System32\msinfo32.exe',
                                           'debugConnectToRunningApp': 'true',
                                           "args": "-port 345"
                                        })
                      
    def clickAdminToolSystemInfoFileExitLink(self):
        parent_window_element = self.driver.find_element_by_class_name("#32770")
        parent_window_element.find_element_by_name("File").click()
        self.utils.sleepUntil(3)
        self.driver.find_element_by_name("Exit").click()                 
                      
                      
    def connectAdminToolWindowsPowershellWebDriver(self):
        self.driver = webdriver.Remote(command_executor="http://localhost:9999",
                                       desired_capabilities={
                                           'app': r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell_ise.exe',
                                           'debugConnectToRunningApp': 'true',
                                           "args": "-port 345"
                                        })

    def clickAdminToolPowerShellCloseButton(self):
        parent__element = self.driver.find_element_by_class_name("ConsoleWindowClass")
        parent__element.find_element_by_name("Close").click()   

    def connectAdminToolDefragmentOptimizeWebDriver(self):
        self.driver = webdriver.Remote(command_executor="http://localhost:9999",
                                       desired_capabilities={
                                           'app': r'C:\Windows\System32\dfrgui.exe',
                                           'debugConnectToRunningApp': 'true',
                                           "args": "-port 345"
                                        })
        
    def connectAdminToolDiskCleanupWebDriver(self):
        self.driver = webdriver.Remote(command_executor="http://localhost:9999",
                                       desired_capabilities={
                                           'app': r'C:\Windows\System32\cleanmgr.exe',
                                           'debugConnectToRunningApp': 'true',
                                           "args": "-port 345"
                                        })    
        
    def clickAdminToolDiskCleanupCancelButton(self):
        parent__element = self.driver.find_element_by_class_name("#32770")
        parent__element.find_element_by_name("Cancel").click() 
            
    def connectAdminTooliSCSiWebDriver(self):
        self.driver = webdriver.Remote(command_executor="http://localhost:9999",
                                       desired_capabilities={
                                           'app': r'C:\Windows\System32\iscsicpl.exe',
                                           'debugConnectToRunningApp': 'true',
                                           "args": "-port 345"
                                        })    
        


