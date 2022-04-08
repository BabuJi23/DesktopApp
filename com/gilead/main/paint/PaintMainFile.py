'''
Created on Aug 7, 2018

@author: bkotamahanti
'''
from selenium import webdriver
import pyautogui
from definitions import ROOT_DIR, PAINT_MAIN_DIR, PAINT_PROCESS
from cv2 import cv2
import numpy

class PaintMainClass(object):
    __process =  PAINT_PROCESS
    __root_dir = ROOT_DIR
    __paint_main_dir = PAINT_MAIN_DIR
    minSearchTime = 2
    
    def __init__(self, utilsRef, loggerRef):
        self.utils = utilsRef
        self.logger = loggerRef
        self.driver = ""
        self.images = self.__root_dir+self.__paint_main_dir+ "images\\"
        
    def killProcess(self):
        self.utils.killProcess(self.__process)
    
    def openPaint(self, filePath):
        self.utils.openApplication(filePath) 
    
    
    ''' Initializing webdriver - by opening new Paint application'''
    def startPaintViaWebdriver(self):
        self.driver = webdriver.Remote(command_executor='http://localhost:9999',
                                       desired_capabilities={
                                               "app" :  r"C:\\Windows\\System32\\mspaint.exe",
                                               "debugConnectToRunningApp": 'false',
                                               "args": "-port 345"
                                               }
                                       )
        self.logger.info("Running winium Driver which opens new Paint application.")
    
    
    ''' Initializing webdriver - by connecting already running Paint application'''
    def connectPaintToWebDriver(self):
        self.driver = webdriver.Remote(command_executor='http://localhost:9999',
                                       desired_capabilities={
                                               "debugConnectToRunningApp": "true",
                                               "app" :  r"C:\\Windows\\System32\\mspaint.exe",
                                               "args": "-port 345"
                                               }
                                       )
        self.logger.info("Connect the already running paint process to winium Driver instead of opening new one.")

    def startPaint_via_StartMenu(self):
        self.utils.startProgramByName_via_StartMenu("paint")
    
    def isPaintProcessExist(self):
        return self.utils.isProcessExist(self.__process)
        
    def getPaintAppName(self):
        return self.driver.find_element_by_class_name("MSPaintApp").get_attribute("Name")
    
    
    def drawSomethingOnEmptyPaintBoard(self):
        distance = 200
        pyautogui.click(503,403) #click to put drawing program in focus
        while distance > 0:
            pyautogui.dragRel(distance, 0, duration=0.2, button='left')#move right 
            distance = distance - 50
            pyautogui.dragRel(0, distance, duration=0.2, button='left')#move down
            pyautogui.dragRel(-distance, 0, duration=0.2, button='left')#move left
            distance = distance - 50
            pyautogui.dragRel(0, -distance, duration=0.2, button='left')#move up
        
        self.logger.info("Drawing is done....")   
    
    
    def takeScreenshot(self, left, top, width, height, fileName):
        pyautogui.screenshot(fileName, (left, top, width, height))
    
    def areImagesSame(self, img1, img2):
        self.logger.info("Reading the images to compare..")
        image1 = cv2.imread(img1)
        image2 = cv2.imread(img1)
        difference = cv2.subtract(image1, image2)
        if not numpy.any(difference):
            return True
        return False
        
    def getTextFromPopUpWindow(self):
        #self.driver.find_element_by_xpath("//*[@AutomationId='MainInstruction']").get_attribute("Name")
        return self.driver.find_element_by_class_name("Element").get_attribute("Name")
    
    
    def clickSaveOnPopUp(self):
        self.driver.find_element_by_name("Save").click()
    
    def clickDontSaveOnPopUp(self):
        self.driver.find_element_by_name("Don't Save").click()
    
    def clickOpenMenuOptionViaWebdriver(self):
        self.driver.find_element_by_name("File tab").click()
        self.driver.find_element_by_name("Open").click()
        self.logger.info("Clicked File->Open menu option ...")
    
    def clickPaintCloseButton(self):
        parent_element = self.driver.find_element_by_name("Untitled - Paint")
        child_element = parent_element.find_element_by_xpath("//*[@AutomationId='TitleBar']")
        child_element.find_element_by_xpath("//*[@AutomationId='Close']").click()
    
    
    ''' Don't want to use this method as it is taking time to maximize window bcaz of xpath known winium webdriver issue '''    
    def maximizePaintApp(self):
        parent_element = self.driver.find_element_by_name("Untitled - Paint")
        child_element = parent_element.find_element_by_xpath("//*[@AutomationId='TitleBar']")
        child_element.find_element_by_xpath("//*[@AutomationId='Maximize-Restore']").click()
    
    
    def getAttributesValuesOfSaveAsDialog(self, *args):
        list_of_attributes = list()
        for arg1,arg2 in args:
            list_of_attributes.append(self.driver.find_element_by_class_name(arg1).get_attribute(arg2))
        
        '''converting list to tuple '''
        return tuple(list_of_attributes)   
    
    def removePaintFile(self, filename):
        self.utils.removeFile(filename)   
    
    def openPreviousSavedFile(self, fileName):
        open_dialog_element = self.driver.find_element_by_name("Open")
        self.utils.sleepUntil(2)
        open_dialog_element.find_element_by_class_name("Edit").send_keys(fileName)
        self.logger.info("Entered the file name {}".format(fileName))
        self.utils.sleepUntil(3)
        open_dialog_element.find_element_by_class_name("Button").click()
        #self.driver.find_element_by_xpath("//*[@AutomationId='1' and @ClassName='Button']").click()
        self.logger.info("Clicked the Open button in Open dialog")
    
    def getCoordinatesByLocatingGivenImageOnScreen(self,imageFile):
        print("==================>"+self.images+imageFile)
        return self.utils.getCoordinatesByLocatingGivenImageOnScreen(self.images+imageFile, self.minSearchTime)
    
    
    def clickOK(self):
        self.driver.find_element_by_class_name("CCPushButton").click()
    
    
    def clickCancel(self):
        self.driver.find_element_by_name("Cancel").click()
    
    
    def enterFileNameAndClickSave(self, fileName):
        self.driver.find_element_by_class_name("Edit").send_keys(fileName)
        self.logger.info("Entered the file name {}".format(fileName))
        self.driver.find_element_by_name("Save").click()
        self.logger.info("Clicked the Save button in Save As dialog")
    
    def clickExitMenuOptionInPaintApp(self):
        self.driver.find_element_by_name("File tab").click()
        self.driver.find_element_by_name("Exit").click()
        self.logger.info("Clicked File->Exit menu option ...")
    
    def clickSaveMenuOptionInPaintApp(self):
        self.driver.find_element_by_name("File tab").click()
        self.driver.find_element_by_name("Save").click()
        self.logger.info("Clicked File->Save menu option ...")    