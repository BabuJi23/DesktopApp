'''
Created on Oct 24, 2018
@author: pparamasivan
'''

from definitions import ACCESS_PROCESS
from selenium import webdriver
#import pypyodbc as pyodbc
import pyautogui as pg
from win32com.client import Dispatch


class AccessMainClass():
    __process = ACCESS_PROCESS

    def __init__(self, utilsRef, loggerRef):
        self.utils = utilsRef
        self.logger = loggerRef
        self.driver = ""
        self.utils.killProcess(self.__process)
    
    def killProcess(self):
        self.utils.killProcess(self.__process)
       
    def startMicroSoftAccessApp(self, file):    
        self.utils.openApplication(file)
    

    def createAccessDatabase(self, dbname):
        try:
            accApp = Dispatch("Access.Application")
            dbEngine = accApp.DBEngine
            workspace = dbEngine.Workspaces(0)

            dbLangGeneral = ';LANGID=0x0409;CP=1252;COUNTRY=0'
            newdb = workspace.CreateDatabase(dbname, dbLangGeneral, 64)

            newdb.Execute("""CREATE TABLE Table1 (
                     ID autoincrement,
                     Name varchar(10),
                     Dept varchar(10),
                     Details varchar(30));""")
            self.logger.info("Created Access Database at {}".format(dbname))
        except Exception as e:
            print(e)

        finally:
            accApp.DoCmd.CloseDatabase
            accApp.Quit
            newdb = None
            workspace = None
            dbEngine = None
            accApp = None


    def isAccessProcessExist(self):
        return self.utils.isProcessExist(self.__process)

    def getAccessDatabaseElement(self):
        parent_element = self.driver.find_element_by_class_name("OMain")
        navi_host_element = parent_element.find_element_by_name("Navigation Pane Host")
        return navi_host_element

    def clickCloseButtonOfAccessApp(self):
#         parent_element = self.driver.find_element_by_class_name("OMain")
#         doc_top_element = parent_element.find_element_by_name("MsoDockTop")
#         ribbon_pane_element = doc_top_element.find_element_by_class_name("MsoCommandBar")
#         pane_element = ribbon_pane_element.find_element_by_class_name("NetUIHWNDElement")
#         ribbon_element = pane_element.find_element_by_name("Ribbon")
#         ribbon_element.find_element_by_name("Close").click()
        pg.hotkey('alt','f4')
    

    def connectAccessToWebDriver(self):
        self.driver = webdriver.Remote(command_executor="http://localhost:9999",
                                       desired_capabilities={
                                           'app': r'C:\Program Files (x86)\Microsoft Office\root\Office16\MSACCESS.EXE',
                                           'debugConnectToRunningApp': 'true',
                                           "args": "-port 345"
                                        })
        

    def isAccessElementExist(self, webelement, **property1):
        return self.utils.isElementPresent( webelement,  **property1)
    
    def isAccessElementTextExist(self, webelement, text , **property1):
        return self.utils.isElementTextPresent(webelement, text , **property1)

    def clickAccessBlankDatabaseLink(self):
        parent_element = self.driver.find_element_by_name("Access")
        pane_element = parent_element.find_element_by_class_name("FullpageUIHost")
        child_element = pane_element.find_element_by_class_name("NetUIFullpageUIWindow")
        backstage_element = child_element.find_element_by_class_name("NetUIScrollViewer")
        group_element = backstage_element.find_element_by_class_name("NetUIElement")
        group1_element = group_element.find_element_by_class_name("NetUIElement")
        featured_element = group1_element.find_element_by_name("Featured")
        featured_element.find_element_by_name("Blank database").submit()

    def clickAccessDatabaseCreateButton(self):
        database_element = self.driver.find_element_by_name("Template Preview")
        database_element.find_element_by_name("Create Blank database").click()


    def enterDataInAccessDatabase(self):
        parent_element = self.driver.find_element_by_class_name("OMain")
        workspace_element = parent_element.find_element_by_class_name("MDIClient")
        table1_element = workspace_element.find_element_by_name("Table1")   
        grid_element = table1_element.find_element_by_name("Grid")
        newrow_element = grid_element.find_element_by_name("New Row")
        row1col2_element = newrow_element.find_element_by_name("Row 1, Column Click to Add, ")
        row1col2_element.find_element_by_name("Edit Box").click()
        pg.typewrite("Name")
        pg.press('right')
        pg.typewrite("DeptName")
        pg.press('right')
        pg.typewrite("Details")
        pg.press('right')
    
    def getAccessDatabaseTableElement(self):
        parent_element = self.driver.find_element_by_class_name("OMain")
        host_element = parent_element.find_element_by_name("Navigation Pane Host")
        pane_element = host_element.find_element_by_class_name("NetUIHWNDElement")
        navpane_element = pane_element.find_element_by_class_name("NetUINetUI")
        table_element = navpane_element.find_element_by_name("Tables")
        return table_element 

        
    def clickAccessDBCloseButton(self):
        parent_element = self.driver.find_element_by_class_name("OMain")
        MSoDocTop_element = parent_element.find_element_by_name("MsoDockTop")
        toolbar_element = MSoDocTop_element.find_element_by_class_name("MsoCommandBar")
        pane_element = toolbar_element.find_element_by_class_name("MsoWorkPane")
        panechild_element = pane_element.find_element_by_class_name("NUIPane")
        panechild1_element = panechild_element.find_element_by_class_name("NetUIHWNDElement")
        ribbon_element = panechild1_element.find_element_by_class_name("NetUInetpane")
        ribbon_element.find_element_by_name("Close").click()

    def clickAccessDialogNoButton(self):   
        parent_element = self.driver.find_element_by_class_name("OMain")
        dialog_element = parent_element.find_element_by_class_name("#32770")
        dialog_element.find_element_by_name("No").click()

    def clickOpenOtherFilesFolderLink(self):
        parent_element = self.driver.find_element_by_name("Access")
        backstage_element = parent_element.find_element_by_name("Backstage view")
        backstage_element.find_element_by_name("Open Other Files").click()
    
    def clickOpenWindowBrowseButton(self):
        parent_element = self.driver.find_element_by_name("Access")
        backstage_element = parent_element.find_element_by_name("Backstage view")
        backstage_element.find_element_by_name("Browse").click()


    def clickOpenButtonOfOpenBrowseWindow(self):
        browse_window_element = self.driver.find_element_by_class_name("#32770")
        elements =  browse_window_element.find_elements_by_class_name("Button")
        for element in elements:
            if element.get_attribute("Name") == "Open":
                element.click()

    def getAccessPopUpWindowElement(self):
        parent_element = self.driver.find_element_by_name("Access")
        open_dialog_element = parent_element.find_element_by_class_name("#32770")
        return open_dialog_element
    
    def clickOKOnAccessOpenDialog(self):
        open_dialog_element = self.getAccessPopUpWindowElement()
        open_dialog_element.find_element_by_name("OK").click()
    
    def clickCancelOnAccessOpenDialog(self):
        browse_window_element = self.driver.find_element_by_class_name("#32770")
        elements =  browse_window_element.find_elements_by_class_name("Button")
        for element in elements:
            if element.get_attribute("Name") == "Cancel":
                element.click()
    
    def clickAccessAppBackButton(self):
        parent_element = self.driver.find_element_by_class_name("OMain")
        parent_element.find_element_by_name("Back").click()    
    
    def clickCloseButtonAccessApp(self):
        parent_element = self.driver.find_element_by_class_name("OMain")
        pane_element = parent_element.find_element_by_class_name("FullpageUIHost")
        pane_child_element = pane_element.find_element_by_class_name("NetUIFullpageUIWindow")
        pane_child_element.find_element_by_name("Close").click()

    def enterFileNameInOpenBrowseWindow(self, fileName):
        browse_window_element = self.driver.find_element_by_class_name("#32770")
        browse_window_element.find_element_by_class_name("Edit").click()
        browse_window_element.find_element_by_class_name("Edit").send_keys(fileName) 
     
    