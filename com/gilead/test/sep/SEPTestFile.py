'''
Created on Sep 13, 2018

@author: bkotamahanti
'''
import inspect
from com.gilead.main.sep.SEPMainFile import SEPMainClass

class SEPTestClass():
    
    def __init__(self, utilsRef, loggerRef):
        self.utils = utilsRef
        self.logger = loggerRef
        self.sep_maincode = SEPMainClass(utilsRef, loggerRef)
    
    def getLocationOfGivenImageMoveAndClickElement(self, imageFile, x_offset, y_offset):
        x_coordinate, y_coordinate = self.sep_maincode.getCoordinatesByLocatingGivenImageOnScreen(imageFile)
        self.logger.info("trying to locate the image "+imageFile)
        assert x_coordinate != 0 and y_coordinate != 0, "either of the coordinates are empty"
        self.logger.info("Located the image "+imageFile+" at coordinates ("+str(x_coordinate)+","+str(y_coordinate)+")")
        self.sep_maincode.moveToLocationAndClick(x_coordinate + x_offset, y_coordinate+y_offset)
    
    def symantec_endpoint_protection_open_close_testcase_id_W10DA_8(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            #TC Step 1
            self.utils.killProcess("SymCorpUI.exe") #Need to check how to force kill the security process-Not working
            
            #TC Step 2
            self.utils.log_assert(self.sep_maincode.startSEP_via_StartMenu(), 
               "SEP app not started properly as it failed to enter program name in windows search bar", self.logger, 
               verbose=True) 
            self.utils.sleepUntil(10)
            assert self.sep_maincode.is_SEP_processExist(), "SEP process didn't start properly"

            #TC Step 3
            assert self.sep_maincode.is_SEP_app_started("202_screen1.PNG"), "SEP application not started properly"  

            #TC Step 4
            x_coordinate, y_coordinate = self.sep_maincode.getCoordinatesByLocatingGivenImageOnScreen("202_screen1.PNG")
            assert (x_coordinate != 0 and y_coordinate !=0) ,"either of the coordinates are empty as it could not locate image 202_screen1.PNG "
            
            self.sep_maincode.moveToLocationAndClick(x_coordinate , y_coordinate)
            #self.sep_maincode.moveToLocationAndClick(1160, 118)
            
            self.sep_maincode.closeSEP()
            self.utils.sleepUntil(3)

            assert self.sep_maincode.is_SEP_processExist()== False,"SymCorpUI.exe didn't close"
            
            return True, "", func.co_name
        except (AssertionError, AttributeError, TypeError) as e:
            return False, e, func.co_name
    
    '''TC: W10DA-26 '''
    def symantec_endpoint_protection_run_active_scan_testcase_id_W10DA_26(self):
        try:
            func = inspect.currentframe().f_back.f_code
            self.logger.info("Starting execution of function {} :".format(func.co_name))
            #TC Step 1
            self.utils.killProcess("SymCorpUI.exe") #Need to check how to force kill the security process-Not working
            
            #TC Step 2
            self.utils.log_assert(self.sep_maincode.startSEP_via_StartMenu(), 
               "SEP app not started properly as it failed to enter program name in windows search bar", self.logger, 
               verbose=True) 
            self.utils.sleepUntil(10)
            assert self.sep_maincode.is_SEP_processExist(), "SEP process didn't start properly"
            assert self.sep_maincode.is_SEP_app_started("202_screen1.PNG"), "SEP application not started properly" 

            #TC Step 3
            # Identify and click on the image 'Scan for Threats' 
            self.getLocationOfGivenImageMoveAndClickElement("SEP_Scan_For_Threats.PNG", 30, 0)
            
            self.utils.sleepUntil(6)
            # Validate images on 'Scan for Threats' screen              
            x_coordinate, y_coordinate = self.sep_maincode.getCoordinatesByLocatingGivenImageOnScreen("SEP_Run_full_scan_text.PNG")
            assert (x_coordinate != 0 and y_coordinate !=0) ,"either of the coordinates are empty"
            x_coordinate, y_coordinate = self.sep_maincode.getCoordinatesByLocatingGivenImageOnScreen("SEP_Scan_For_Threats_text.PNG")
            assert (x_coordinate != 0 and y_coordinate !=0) ,"either of the coordinates are empty"

            #click on Run Active Scan Element 
            self.getLocationOfGivenImageMoveAndClickElement("SEP_Run_active_scan.PNG", 60, 13)
            
            self.utils.sleepUntil(6)

            #validate 'SEP Run active scan sub process'started or not
            assert self.sep_maincode.is_SEP_run_active_scan_sub_process_started(), "SEP Run active scan sub process didn't start"

            #TC Step 4
            #click on Pause scan button in sub process window
            self.getLocationOfGivenImageMoveAndClickElement("SEP_Run_active_scan_Pause_scan_image.PNG", 60, 13)
            self.utils.sleepUntil(6)

            #TC Step 5
            #click on Resume Scan button in sub process window
            self.getLocationOfGivenImageMoveAndClickElement("SEP_Run_active_scan_Resume_scan_image.PNG", 60, 13)
            self.utils.sleepUntil(6)
            #validate after Resume blue color Pause button image shows up, Completed text and Close button shows up
            x_coordinate, y_coordinate = self.sep_maincode.getCoordinatesByLocatingGivenImageOnScreen("SEP_Run_active_scan_PauseBlueColor_scan_image.PNG")
            assert (x_coordinate != 0 and y_coordinate !=0) ,"either of the coordinates are empty or 0"
            self.utils.sleepUntil(10)
            x_coordinate, y_coordinate = self.sep_maincode.getCoordinatesByLocatingGivenImageOnScreen("SEP_Run_active_scan_completed_image.PNG")
            assert (x_coordinate != 0 and y_coordinate !=0) ,"either of the coordinates are empty or 0"

            #TC Step 6
#           Close the sub-process window and validate   
            self.utils.sleepUntil(6)
            self.getLocationOfGivenImageMoveAndClickElement("SEP_Run_active_scan_close_image.PNG", 30, 13)
            self.utils.sleepUntil(6)
            #validate the sub process is closed or not
#             self.utils.sleepUntil(2)
            assert (self.sep_maincode.is_SEP_run_active_scan_sub_process_started()== False), "SEP Run active scan sub process didn't close"

            #TC Step 7
            #close the SEP process
            self.getLocationOfGivenImageMoveAndClickElement("202_screen1.PNG", 0, 0)
            self.utils.sleepUntil(6)
            self.sep_maincode.closeSEP()
            
            return True, "", func.co_name
        except (AssertionError, AttributeError, TypeError) as e:
            return False, e, func.co_name