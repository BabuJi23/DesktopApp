'''
Created on Sep 7, 2018

@author: bkotamahanti
'''
from selenium import webdriver
from definitions import ONENOTE_PROCESS
#import pyautogui

class OneNoteMainClass(object):
    __process = ONENOTE_PROCESS
    
    def __init__(self, utilsRef, loggerRef):
        self.utils = utilsRef
        self.logger = loggerRef
    
    def killProcess(self):
        self.utils.killProcess(self.__process)
        
    def startOneNote_via_StartMenu(self):
        self.utils.startProgramByName_via_StartMenu("OneNote")
    
    def isOneNoteProcessExist(self):
        return self.utils.isProcessExist(self.__process)
        
    def connectOneNoteToWebDriver(self):
        self.driver = webdriver.Remote(command_executor="http://localhost:9999",
                                       desired_capabilities={ 
#                                                             "app" : r'C:\Program Files (x86)\Microsoft Office\Office16\ONENOTE.EXE',
                                                             "app" : r'ONENOTE.EXE',
                                                             "debugConnectToRunningApp": 'true',
                                                             "args": "-port 345"
                                                             }
                                       )
        
    def openApplication(self, fileName):
        self.utils.openApplication(fileName)
    
    def getOneNoteMainWindowDoctopPaneElement(self):
        parent_element = self.driver.find_element_by_class_name("Framework::CFrame")
        MsoDockTop_element = parent_element.find_element_by_name("MsoDockTop")
        return MsoDockTop_element
    
    def isOneNoteElementExist(self, webelement, **property1):
        return self.utils.isElementPresent( webelement,  **property1)
    
    def clickCloseButtonOfOneNoteApp(self):
        mso_doctop_element = self.getOneNoteMainWindowDoctopPaneElement()
        mso_doctop_element.find_element_by_name("Close").click()
        #pyautogui.hotkey('alt','f4')
        
    def clickNewOneNote(self):
        mso_doctop_element = self.getOneNoteMainWindowDoctopPaneElement()
        mso_doctop_element.find_element_by_name("File Tab").click()
        self.utils.sleepUntil(2)
        back_stage_view_element = self.driver.find_element_by_name("Backstage view")
        back_stage_view_element.find_element_by_name("New").click()
        
    
    def getOneNoteNewGroupSavingfeaturesElement(self):
        back_stage_view_element = self.driver.find_element_by_name("Backstage view")
        return   back_stage_view_element.find_element_by_name("Saving Features")
    
    def clickBrowseElementInNewTemplate(self):
        back_stage_view_element = self.driver.find_element_by_name("Backstage view")
        saving_features_element = back_stage_view_element.find_element_by_name("Saving Features")
        saving_features_element.find_element_by_name("Browse").click()
    
    def getOneNoteNewGroupPickAFolderElement(self):
        back_stage_view_element = self.driver.find_element_by_name("Backstage view")
        return back_stage_view_element.find_element_by_name("Pick a Folder")
#         new_group_element = back_stage_view_element.find_element_by_name("New")
#         return new_group_element.find_element_by_name("Pick a Folder")
    
    def clickSecondBrowseElementInNewTemplate(self):
        back_stage_view_element = self.driver.find_element_by_name("Backstage view")
        pick_a_folder_element = back_stage_view_element.find_element_by_name("Pick a Folder")
        
        elements = pick_a_folder_element.find_elements_by_name("Browse")
        for element in elements:
            if "NetUIStickyButton" == element.get_attribute("ClassName"):
                element.click()
        
        
#         pick_a_folder_element.find_element_by_name("NetUIStickyButton").click()
#         pick_a_folder_element.find_element_by_name("Enter a notebook name:").send_keys(fileName)
#         self.utils.sleepUntil(2)
#         pick_a_folder_element.find_element_by_name("Create Notebook").click()
    
    def enterFileNameInNewBrowseWindow(self, fileName):
        browse_window_element = self.driver.find_element_by_name("Create New Notebook")
        browse_pane = browse_window_element.find_element_by_class_name("DUIViewWndClassName")
        browse_pane.find_element_by_class_name("Edit").click()
        browse_pane.find_element_by_class_name("Edit").send_keys(fileName)
        self.utils.sleepUntil(2)
        browse_window_element.find_element_by_name("Create").click()
    
    
    def enterFileNameInOpenBrowseWindow(self, fileName):
        browse_window_element = self.driver.find_element_by_class_name("#32770")
        browse_window_element.find_element_by_name("Open Notebook").find_element_by_class_name("Edit").click()
        browse_window_element.find_element_by_name("Open Notebook").find_element_by_class_name("Edit").send_keys(fileName) 

    
    
    def clickOpenButtonOfOpenBrowseWindow(self):
        browse_window_element = self.driver.find_element_by_class_name("#32770")
        browse_window_element.find_element_by_name("Open").click()
#         browse_window_element = self.driver.find_element_by_name("Open Notebook")
#         browse_pane = browse_window_element.find_element_by_class_name("ComboBoxEx32")
#         browse_pane.find_element_by_name("Open").click()
      
    def getOneNoteNotebookDropDownElement(self):
        parent_element = self.driver.find_element_by_class_name("Framework::CFrame")
        return parent_element.find_element_by_class_name("NetUINotebookDropdownButton")
    
    def enterText(self, textToEnter):
        parent_element = self.driver.find_element_by_class_name("Framework::CFrame")    
        onenote_main_window_element = parent_element.find_element_by_class_name("OneNote::MainWindow")
        onenote_workspace_pane_element = onenote_main_window_element.find_element_by_class_name("OneNote::CWorkspace")
        onenote_document_element = onenote_workspace_pane_element.find_element_by_class_name("")
#         onenote_document_element = onenote_workspace_pane_element.find_element_by_class_name("OneNote::DocumentCanvas")
        onenote_current_page_document_element = onenote_document_element.find_element_by_name("OneNote current page")
        onenote_current_page_document_element.click()
        onenote_current_page_document_element.send_keys(textToEnter)
    
    def isOneNoteOutlineElementExist(self, elementName):
        parent_element = self.driver.find_element_by_class_name("Framework::CFrame")    
        onenote_main_window_element = parent_element.find_element_by_class_name("OneNote::MainWindow")
        onenote_workspace_pane_element = onenote_main_window_element.find_element_by_class_name("OneNote::CWorkspace")
#         onenote_document_element = onenote_workspace_pane_element.find_element_by_class_name("OneNote::DocumentCanvas")
        onenote_document_element = onenote_workspace_pane_element.find_element_by_class_name("")
        onenote_current_page_document_element = onenote_document_element.find_element_by_name("OneNote current page")
        outline_element = onenote_current_page_document_element.find_element_by_name("Content block")
        if outline_element.text == elementName :
            return True
        return False    
    
    def removeFile(self, fileName):
        self.utils.removeDir(fileName)
    
    def clickOpenOneNoteMenuOption(self):
        mso_doctop_element = self.getOneNoteMainWindowDoctopPaneElement()
        mso_doctop_element.find_element_by_name("File Tab").click()
        self.utils.sleepUntil(2)
        back_stage_view_element = self.driver.find_element_by_name("Backstage view")
        back_stage_view_element.find_element_by_name("Open").click()
    
    
    def clickScollDownButton(self):
        back_stage_view_element = self.driver.find_element_by_name("Backstage view")
        scroll_bar_element = back_stage_view_element.find_element_by_class_name("NetUIScrollBar")
        for i in range(10):
            scroll_bar_element.find_element_by_name("Line down").click()
            
    
    def clickBrowseElementOfOpenTemplate(self):
        back_stage_view_element = self.driver.find_element_by_name("Backstage view")
        elements = back_stage_view_element.find_elements_by_name("Open")
        for element in elements:
            if element.get_attribute("ClassName") == 'NetUIAdaptiveStylingContainer' :
               element.find_element_by_class_name("NetUIElement").find_element_by_name("Browse").click()
               break
        #2nd solution in case the above solution don't work
            #         open_group_elements = back_stage_view_element.find_elements_by_class_name("NetUIElement")
            #         len(open_group_elements)
            #         for open_group_element in open_group_elements:
            #             if open_group_element.get_attribute("Name") == "":
            #                open_group_element.find_element_by_name("Browse").click()
            #                break 
            
            
    def getOneNoteMicrosoftOneNoteDialog(self):
        return self.driver.find_element_by_name("Microsoft OneNote")
         
    
    def clickOKOnMicrosoftOneNoteDialog(self):
        microsoft_oneNote_dialog = self.driver.find_element_by_name("Microsoft OneNote") 
        microsoft_oneNote_dialog.find_element_by_name("OK").click()
    
    
    def clickOneNoteAppBackButton(self):
        parent_element = self.driver.find_element_by_class_name("Framework::CFrame") 
        parent_element.find_element_by_name("Back").click()