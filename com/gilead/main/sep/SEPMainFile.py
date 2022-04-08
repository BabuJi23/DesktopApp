'''
Created on Jul 25, 2018

@author: bkotamahanti
'''
import pyautogui
import time
from definitions import SEP_PROCESS_NAME
from definitions import SEP_RUN_ACTIVE_SCAN_SUB_PROCESS
from definitions import ROOT_DIR, SEP_MAIN_DIR

class SEPMainClass(object):
    minSearchTime = 5
    wait_time = 2
    __process = SEP_PROCESS_NAME
    __sub_process = SEP_RUN_ACTIVE_SCAN_SUB_PROCESS
    __root_dir = ROOT_DIR
    __sep_main_dir = SEP_MAIN_DIR
    
    def __init__(self, utilsRef, loggerRef):
        self.images = self.__root_dir + self.__sep_main_dir + "images\\"
        self.utils = utilsRef
        self.logger = loggerRef
    
    def startSEP_via_StartMenu(self):
        return self.utils.startProgramByName_via_StartMenu("Symantec")
    
    def is_SEP_run_active_scan_sub_process_started(self):
        return self.utils.isProcessExist( self.__sub_process )    
    
    def is_SEP_processExist(self):    
        return self.utils.isProcessExist(self.__process)
        
    def is_SEP_app_started(self, imageFile):
        return self.utils.isAppStarted(self.images+imageFile, self.minSearchTime)   
        
    def getCoordinatesByLocatingGivenImageOnScreen(self,imageFile):
        return self.utils.getCoordinatesByLocatingGivenImageOnScreen(self.images+imageFile, self.minSearchTime)
    
    def moveToLocationAndClick(self, x_coor, y_coor):
        self.utils.moveToLocationAndClick(x_coor, y_coor)

    def closeSEP(self):
        #pyautogui.moveRel(90,-15)
        pyautogui.moveTo(1166, 113)
        time.sleep(self.wait_time)
        self.logger.info("Closing SEP application" )
        pyautogui.click()