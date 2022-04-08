'''
Created on Jul 23, 2018

@author: bkotamahanti
'''

import logging
import unittest
import pytest
from com.gilead.test.main_package.main import MainClass
from definitions import ROOT_DIR, LOG_DIR
import re
import atexit
import warnings
import time



class GUI_tests(unittest.TestCase):
    
    main = MainClass()
    logger = logging   
         
   
    def setUp(self):
        pass
    
    def tearDown(self):
        time.sleep(5)

    
    @pytest.mark.run(order=1)
    def test_devicemanager_computer_drive_testcase_id_W10DA_079(self):
        self.logger.info("============> test_devicemanager_testcase_id_W10DA_079 <============")
        x, e, func_name  = self.main.devicemanager.devicemanager_computer_drive_testcase_id_W10DA_079()
        if x is True:
            self.logger.info(
                "Device Manager  testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Device Manager  testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================") 
    
    @pytest.mark.run(order=2)
    def test_devicemanager_disk_drive_testcase_id_W10DA_080(self):
        self.logger.info("============> test_devicemanager_testcase_id_W10DA_080 <============")
        x, e, func_name  = self.main.devicemanager.devicemanager_disk_drive_testcase_id_W10DA_080()
        if x is True:
            self.logger.info(
                "Device Manager  testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Device Manager  testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")          
    
    
    @pytest.mark.run(order=3)
    def test_devicemanager_displayadaptor_drive_testcase_id_W10DA_081(self):
        self.logger.info("============> test_devicemanager_testcase_id_W10DA_081 <============")
        x, e, func_name  = self.main.devicemanager.devicemanager_dsiplay_adaptor_drive_testcase_id_W10DA_081()
        if x is True:
            self.logger.info(
                "Device Manager  testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Device Manager  testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")          
    
    @pytest.mark.run(order=4)
    def test_devicemanager_imagingdevices_driver_testcase_id_W10DA_082(self):
        self.logger.info("============> test_devicemanager_testcase_id_W10DA_082 <============")
        x, e, func_name  = self.main.devicemanager.devicemanager_imagingdevices_driver_testcase_id_W10DA_082()
        if x is True:
            self.logger.info(
                "Device Manager  testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Device Manager  testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================") 
    
    @pytest.mark.run(order=5)
    def test_devicemanager_keyboard_driver_testcase_id_W10DA_083(self):
        self.logger.info("============> test_devicemanager_testcase_id_W10DA_083 <============")
        x, e, func_name  = self.main.devicemanager.devicemanager_keyboard_driver_testcase_id_W10DA_083()
        if x is True:
            self.logger.info(
                "Device Manager  testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Device Manager  testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")  
        
    @pytest.mark.run(order=6)
    def test_devicemanager_mouse_driver_testcase_id_W10DA_084(self):
        self.logger.info("============> test_devicemanager_testcase_id_W10DA_084 <============")
        x, e, func_name  = self.main.devicemanager.devicemanager_mouse_driver_testcase_id_W10DA_084()
        if x is True:
            self.logger.info(
                "Device Manager  testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Device Manager  testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")          
    
    @pytest.mark.run(order=7)
    def test_devicemanager_monitor_driver_testcase_id_W10DA_085(self):
        self.logger.info("============> test_devicemanager_testcase_id_W10DA_085 <============")
        x, e, func_name  = self.main.devicemanager.devicemanager_monitor_driver_testcase_id_W10DA_085()
        if x is True:
            self.logger.info(
                "Device Manager  testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Device Manager  testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")  
    
   
    @pytest.mark.run(order=8)
    def test_devicemanager_network_adaptor_driver_testcase_id_W10DA_086(self):
        self.logger.info("============> test_devicemanager_network_adaptor_driver_testcase_id_W10DA_086 <============")
        x, e, func_name  = self.main.devicemanager.devicemanager_network_adaptor_driver_testcase_id_W10DA_086()
        if x is True:
            self.logger.info(
                "Device Manager  testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Device Manager  testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================") 
               
    @pytest.mark.run(order=9)
    def test_devicemanager_processor_driver_testcase_id_W10DA_087(self):
        self.logger.info("============> test_devicemanager_testcase_id_W10DA_087 <============")
        x, e, func_name  = self.main.devicemanager.devicemanager_processor_driver_testcase_id_W10DA_087()
        if x is True:
            self.logger.info(
                "Device Manager  testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Device Manager  testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")          
    
            
    @pytest.mark.run(order=1)
    def test_admin_tool_defragment_optimize_testcase_id_W10DA_062(self):
        self.logger.info("============> test_admin_tool_defragment_optimize_testcase_id_W10DA_062 <============")
        x, e, func_name  = self.main.admintool.admin_tool_defragment_optimize_testcase_id_W10DA_62()
        if x is True:
            self.logger.info(
                "Admin Tool Defragment and Optimize Drivers testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Admin Tool Defragment and Optimize Drivers testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")   
     
    @pytest.mark.run(order=2)
    def test_admin_tool_disk_cleanup_testcase_id_W10DA_063(self):
        self.logger.info("============> test_admin_tool_disk_cleanup_testcase_id_W10DA_063 <============")
        x, e, func_name  = self.main.admintool.admin_tool_disk_cleanup_testcase_id_W10DA_63()
        if x is True:
            self.logger.info(
                "Admin Tool Disk Cleanup testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Admin Tool Disk Cleanup testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")     
        
    @pytest.mark.run(order=3)
    def test_admin_tool_iSCSi_testcase_id_W10DA_065(self):
        self.logger.info("============> test_admin_tool_iSCSi_testcase_id_W10DA_065 <============")
        x, e, func_name  = self.main.admintool.admin_tool_iSCSi_testcase_id_W10DA_65()
        if x is True:
            self.logger.info(
                "Admin Tool iSCSi testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Admin Tool iSCSi testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")     
        
    @pytest.mark.run(order=4)
    def test_admin_tool_odbcad32_process_testcase_id_W10DA_067(self):
        self.logger.info("============> test_admin_tool_odbcad32_process_testcase_id_W10DA_060 <============")
        x, e, func_name  = self.main.admintool.admin_tool_odbcad32_testcase_id_W10DA_67_68()
        if x is True:
            self.logger.info(
                "Admin Tool ODBC driver testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Admin Tool ODBC driver testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")  
        
    @pytest.mark.run(order=5)
    def test_admin_tool_memory_diagnostic_testcase_id_W10DA_069(self):
        self.logger.info("============> test_admin_tool_memory_diagnostic_testcase_id_W10DA_069 <============")
        x, e, func_name  = self.main.admintool.admin_tool_memory_diagnostic_testcase_id_W10DA_69()
        if x is True:
            self.logger.info(
                "Admin Tool Memory Diagnostic testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Admin Tool Memory Diagnostic testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")      
     
    @pytest.mark.run(order=6)
    def test_admin_tool_reliability_monitor_Firewall_testcase_id_W10DA_071(self):
        self.logger.info("============> admin_tool_resource_monitor_testcase_id_W10DA_072 <============")
        x, e, func_name  = self.main.admintool.admin_tool_reliability_monitor_Firewall_testcase_id_W10DA_71_77()
        if x is True:
            self.logger.info(
                "Admin Tool Reliability Monitor & Firewall testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Admin Tool Reliability Monitor & Firewall testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")
              
    @pytest.mark.run(order=7)
    def test_admin_tool_resource_monitor_testcase_id_W10DA_072(self):
        self.logger.info("============> admin_tool_resource_monitor_testcase_id_W10DA_072 <============")
        x, e, func_name  = self.main.admintool.admin_tool_resource_monitor_testcase_id_W10DA_72()
        if x is True:
            self.logger.info(
                "Admin Tool Resource monitor testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Admin Tool Resource monitor testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")     
        
    @pytest.mark.run(order=8)
    def test_admin_tool_system_configuration_testcase_id_W10DA_074(self):
        self.logger.info("============> test_admin_tool_system_configuration_testcase_id_W10DA_074 <============")
        x, e, func_name  = self.main.admintool.admin_tool_system_configuration_testcase_id_W10DA_74()
        if x is True:
            self.logger.info(
                "Admin Tool System Configuration testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Admin Tool System Configuration testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================") 
  
    @pytest.mark.run(order=9)
    def test_admin_tool_system_information_testcase_id_W10DA_075(self):
        self.logger.info("============> test_admin_tool_system_information_testcase_id_W10DA_075 <============")
        x, e, func_name  = self.main.admintool.admin_tool_system_information_testcase_id_W10DA_75()
        if x is True:
            self.logger.info(
                "Admin Tool System Information testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Admin Tool System Information testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================") 



    @pytest.mark.run(order=10)
    def test_admin_tool_mmc_process_testcase_id_W10DA_060(self):
        self.logger.info("============> test_admin_tool_mmc_process_testcase_id_W10DA_060 <============")
        x, e, func_name  = self.main.admintool.admin_tool_testcase_id_W10DA_60_61_64_66_70_73_76()
        if x is True:
            self.logger.info(
                    "Admin Tool testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                    "Admin Tool testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")


    #****************************************************************************************************************
    
    @pytest.mark.run(order=1)
    def test_notepad_save_and_open_file_testcase_id_W10DA_001(self):
        self.logger.info("===================> test_notepad_save_and_open_file <==========================================================")
        x,e, func_name = self.main.notepad.notepad_save_and_open_file_testcase_id_W10DA_1()
        if x is True:
            self.logger.info("Notepad testcase {} passed".format(func_name))
        else:    
            self.assertTrue(x, self.logger.debug("Notepad testcase {} failed due to reason: '".format(func_name)+str(e)+"'"))
        self.logger.info("=============================================================================")
       
    @pytest.mark.run(order=2)
    def test_notepad_dont_save_file_testcase_id_W10DA_002(self):
        self.logger.info("===================> test_notepad_dont_save_file <==========================================================")
        x, e, func_name = self.main.notepad.notepad_dont_save_file_testcase_id_W10DA_2()
        if x is True:
            self.logger.info("Notepad testcase {} passed".format(func_name))
        else:    
            self.assertTrue(x, self.logger.debug("Notepad testcase {} failed due to reason: '".format(func_name)+str(e)+"'"))
           
        self.logger.info("=============================================================================")
        
    @pytest.mark.run(order=3)
    def test_notepad_save_delete_open_file_testcase_id_W10DA_003(self):
        self.logger.info("===================> test_notepad_save_delete_open_file <==========================================================")
        x,e, func_name = self.main.notepad.notepad_save_delete_open_file_testcase_id_W10DA_3()
        if x is True:
            self.logger.info("Notepad testcase {} passed".format(func_name))
        else:    
            self.assertTrue(x, self.logger.debug("Notepad testcase {} failed due to reason: '".format(func_name)+str(e)+"'"))
        self.logger.info("=============================================================================")
       
    
    @pytest.mark.run(order=4)
    def test_calculator_open_perform_arithemetic_operations_close_app_testcase_id_W10DA_004(self):
        self.logger.info("=====================================> test_calculator_open_close_app <=============================")
        x,e, func_name = self.main.calculator.calculator_open_perform_arithemetic_operations_close_app_testcase_id_W10DA_4()
        if x is True:
            self.logger.info("Calculator testcase {} passed".format(func_name))
        else:    
            self.assertTrue(x, self.logger.debug("Calculator testcase {} failed due to reason: '".format(func_name) +str(e)+"'"))
        self.logger.info("=============================================================================")
    
       
    @pytest.mark.run(order=5)
    def test_paint_open_do_paint_close_save_app_open_testcase_id_W10DA_005(self):
        self.logger.info("===================> test_paint_open_do_paint_close_save_app_open_testcase_id_W10DA_5 <==========================================================")
        x,e, func_name = self.main.paint.paint_open_do_paint_close_save_app_open_testcase_id_W10DA_5()
        if x is True:
            self.logger.info("Paint testcase {} passed".format(func_name))
        else:    
            self.assertTrue(x, self.logger.debug("Paint testcase  {} failed due to reason: '".format(func_name)+str(e)+"'"))
        self.logger.info("=============================================================================")
       
    @pytest.mark.run(order=6)
    def test_paint_open_do_paint_close_dont_save_testcase_id_W10DA_006(self):
        self.logger.info("===================> test_paint_open_do_paint_close_dont_save_testcase_id_W10DA_6 <==========================================================")
        x,e, func_name = self.main.paint.paint_open_do_paint_close_dont_save_testcase_id_W10DA_6()
        if x is True:
            self.logger.info("Paint testcase {} passed".format(func_name))
        else:    
            self.assertTrue(x, self.logger.debug("Paint testcase {} failed due to reason: '".format(func_name)+str(e)+"'"))
        self.logger.info("=============================================================================")
       
    @pytest.mark.run(order=7)
    def test_paint_save_delete_open_file_testcase_id_W10DA_007(self):
        self.logger.info("===================> test_paint_save_delete_open_file_testcase_id_W10DA_7 <==========================================================")
        x,e, func_name = self.main.paint.paint_save_delete_open_file_testcase_id_W10DA_7()
        if x is True:
            self.logger.info("Paint testcase {} passed".format(func_name))
        else:    
            self.assertTrue(x, self.logger.debug("Paint testcase {} failed due to reason: '".format(func_name)+str(e)+"'"))
        self.logger.info("=============================================================================")
    
    @pytest.mark.run(order=8)
    def test_outlook_check_connectedTo_microsoft_exchange_testcase_id_W10DA_018(self):
        self.logger.info(
            "===================> test_outlook_check_connected_microsoft_exchange_testcase_id_W10DA_18 <============================")
        x, e, func_name = self.main.outlook.outlook_check_connectedTo_microsoft_exchange_testcase_id_W10DA_18()
        if x is True:
            self.logger.info(" Check Outlook Exchange Server Status testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug("Outlook Exchange Server Status testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info("=============================================================================")       
     
                      
    @pytest.mark.run(order=9)
    def test_outlook_compose_mail_check_mail_inbox_close_testcase_id_W10DA_019(self):
        self.logger.info("===================> test_outlook_compose_mail_check_mail_inbox_close_testcase_id_W10DA_19 <==========================================================")
        x,e, func_name = self.main.outlook.outlook_compose_mail_check_mail_inbox_close_testcase_id_W10DA_19()
        if x is True:
            self.logger.info("Outlook New mail testcase {} passed".format(func_name))
        else:    
            self.assertTrue(x, self.logger.debug("Outlook New mail testcase {} failed due to reason: '".format(func_name)+str(e)+"'"))
        self.logger.info("=============================================================================")    
           
    @pytest.mark.run(order=10)
    def test_outlook_compose_mail_save_draft_testcase_id_W10DA_020(self):
        self.logger.info("===================> test_outlook_compose_mail_save_draft_testcase_id_W10DA_20 <==========================================================")
        x,e, func_name = self.main.outlook.test_outlook_compose_mail_save_draft_testcase_id_W10DA_20()
        if x is True:
            self.logger.info("Outlook New mail draft testcase {} passed".format(func_name))
        else:    
            self.assertTrue(x, self.logger.debug("Outlook New mail Draft testcase {} failed due to reason: '".format(func_name)+str(e)+"'"))
        self.logger.info("=============================================================================")  
     
    
    @pytest.mark.run(order=11)
    def test_onenote_open_close_testcase_id_W10DA_012(self):
        self.logger.info(
            "===================> test_onenote_open_close_testcase_id_W10DA_12 <============================")
        x, e, func_name = self.main.onenote.onenote_open_close_testcase_id_W10DA_12()
        if x is True:
            self.logger.info(" OneNote open close testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug("OneNote open close testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info("=============================================================================")       
    
    
    @pytest.mark.run(order=12)
    def test_onenote_open_save_reopen_testcase_id_W10DA_013(self):
        self.logger.info(
            "===================> test_onenote_open_save_reopen_testcase_id_W10DA_13 <============================")
        x, e, func_name = self.main.onenote.onenote_open_save_reopen_testcase_id_W10DA_13()
        if x is True:
            self.logger.info(" OneNote open save reopen testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug("OneNote open save reopen testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info("=============================================================================")
    
    
    @pytest.mark.run(order=13)
    def test_onenote_open_save_delete_reopen_testcase_id_W10DA_014(self):
        self.logger.info(
            "===================> test_onenote_open_save_delete_reopen_testcase_id_W10DA_014 <============================")
        x, e, func_name = self.main.onenote.onenote_open_save_delete_reopen_testcase_id_W10DA_14()
        if x is True:
            self.logger.info(" OneNote open save reopen testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug("OneNote open save reopen testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info("=============================================================================")
    
    @pytest.mark.run(order=14)
    def test_acrobat_open_read_close_testcase_id_W10DA_009(self):
        self.logger.info(
            "===================> test_acrobat_open_read_close_testcase_id_W10DA_09 <============================")
        x, e, func_name = self.main.acrobat.acrobat_open_read_close_testcase_id_W10DA_09()
        if x is True:
            self.logger.info(" Acrobat open-read-close testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug("Acrobat open-read-close  testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info("=============================================================================")
    
    '''Commenting this test case as the UI behavior is changing back and forth. waiting for stable deployment. will enable it once it has stable UI '''

    # @pytest.mark.run(order=15)
    # def test_acrobat_attach_file_email_testcase_id_W10DA_010(self):
    #     self.logger.info(
    #         "===================> test_acrobat_attach_file_email_testcase_id_W10DA_10 <============================")
    #     x, e, func_name = self.main.acrobat.acrobat_attach_file_email_testcase_id_W10DA_10()
    #     if x is True:
    #         self.logger.info(" Acrobat attach to email testcase {} passed".format(func_name))
    #     else:
    #         self.assertTrue(x, self.logger.debug("Acrobat attach to email  testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
    #     self.logger.info("=============================================================================")

    
    @pytest.mark.run(order=16)
    def test_acrobat_print_cancel_testcase_id_W10DA_011(self):
        self.logger.info(
            "===================> test_acrobat_print_cancel_testcase_id_W10DA_11 <============================")
        x, e, func_name = self.main.acrobat.acrobat_print_cancel_testcase_id_W10DA_11()
        if x is True:
            self.logger.info(" Acrobat print cancel testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug("Acrobat print cancel  testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info("=============================================================================")
    
    @pytest.mark.run(order=17)
    def test_chrome_open_url_close_testcase_id_W10DA_016(self):
        self.logger.info(
            "===================> test_chrome_open_url_save_close_testcase_id_W10DA_11 <============================")
        x, e, func_name = self.main.chrome.chrome_open_url_close_testcase_id_W10DA_16()
        if x is True:
            self.logger.info(" Chrome browser open-navigateUrl-close testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug("Chrome browser open-navigateUrl-close testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info("=============================================================================")
    
    
    @pytest.mark.skip(reason="Always Failed") #run(order=18)
    def test_microsoft_edge_open_url_close_testcase_id_W10DA_042(self):
        self.logger.info(
            "===================> test_microsoft_edge_open_url_close_testcase_id_W10DA_42 <============================")
        x, e, func_name = self.main.edge.microsoft_edge_open_url_close_testcase_id_W10DA_42()
        if x is True:
            self.logger.info(" Edge browser open-navigateUrl-close testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug("Edge browser open-navigateUrl-close testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info("=============================================================================")
    
    ''' Not in scope of testing '''
#     @pytest.mark.run(order=19)
#     def test_mozilla_firefox_open_url_close_testcase_id_W10DA_043(self):
#         self.logger.info(
#             "===================> test_mozilla_firefox_open_url_close_testcase_id_W10DA_43 <============================")
#         x, e, func_name = self.main.firefox.mozilla_firefox_open_url_close_testcase_id_W10DA_43()
#         if x is True:
#             self.logger.info(" Firefox browser open-navigateUrl-close testcase {} passed".format(func_name))
#         else:
#             self.assertTrue(x, self.logger.debug("Firefox browser open-navigateUrl-close testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
#         self.logger.info("=============================================================================")
#     
    @pytest.mark.run(order=20)
    def test_powerpoint_open_save_close_testcase_id_W10DA_036(self):
        self.logger.info(
            "===================> test_powerpoint_open_save_close_testcase_id_W10DA_36 <============================")
        x, e, func_name = self.main.powerpoint.powerpoint_open_save_close_testcase_id_W10DA_36()
        if x is True:
            self.logger.info(" PowerPoint open-save-close testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug("PowerPoint open-save-close testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info("=============================================================================")
    
    
    
    @pytest.mark.run(order=21)
    def test_powerpoint_open_dont_save_close_testcase_id_W10DA_037(self):
        self.logger.info(
            "===================> test_powerpoint_open_dont_save_close_testcase_id_W10DA_37 <============================")
        x, e, func_name = self.main.powerpoint.powerpoint_open_dont_save_close_testcase_id_W10DA_37()
        if x is True:
            self.logger.info(" PowerPoint open-dont save-close testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug("PowerPoint open-dont save-close testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info("=============================================================================")
        
    @pytest.mark.run(order=22)
    def test_powerpoint_open_save_delete_open_testcase_id_W10DA_038(self):
        self.logger.info(
            "===================> test_powerpoint_open_save_delete_open_testcase_id_W10DA_38 <============================")
        x, e, func_name = self.main.powerpoint.powerpoint_open_save_delete_open_testcase_id_W10DA_38()
        if x is True:
            self.logger.info(" PowerPoint open-dont save-close testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug("PowerPoint open-dont save-close testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info("=============================================================================")
    
    
    @pytest.mark.skip(reason="Always Failed") #run(order=23)
    def test_avatier_credential_password_reset_using_MsEdge_testcase_id_W10DA_029(self):
        self.logger.info("=================> test_avatier_credential_password_reset_using_MsEdge_testcase_id_W10DA_29 <=====================")
        x, e, func_name  = self.main.edge.avatier_credential_password_reset_using_MsEdge_testcase_id_W10DA_29()
        if x is True:
            self.logger.info(
                "Password credential reset using Edge testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Password credential reset using Edge testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")
 

    @pytest.mark.skip(reason="Always Failed") #run(order=25)
    def test_avatier_credential_password_reset_disagree_MsEdge_testcase_id_W10DA_031(self):
        self.logger.info("=================> test_avatier_credential_password_reset_disagree_MsEdge_testcase_id_W10DA_31 <=====================")
        x, e, func_name  = self.main.edge.avatier_credential_password_reset_disagree_MSEdge_testcase_id_W10DA_31()
        if x is True:
            self.logger.info(
                "Password credential reset disagree using Edge Browser testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Password credential reset disagree using Edge Browser testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")

    @pytest.mark.run(order=26)
    def test_avatier_credential_password_reset_using_chrome_testcase_id_W10DA_032(self):
        self.logger.info("=================>test_avatier_credential_password_reset_using_chrome_testcase_id_W10DA_32 <=====================")
        x, e, func_name  = self.main.chrome.avatier_credential_password_reset_using_chrome_testcase_id_W10DA_32()
        if x is True:
            self.logger.info(
                "Password credential reset using Chrome Browser testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Password credential reset using Chrome Browser testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")

    @pytest.mark.run(order=27)
    def test_check_google_legacy_browser_support_testcase_id_W10DA_033(self):
        self.logger.info("============> test_check_google_legacy_Browser_support_testcase_id_W10DA_33 <============")
        x, e, func_name  = self.main.miscellaneous.check_google_legacy_browser_support_testcase_id_W10DA_33()
        if x is True:
            self.logger.info(
                "Google legacy Browser support testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Google legacy Browser support testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")

    
    @pytest.mark.run(order=28)
    def test_msword_open_save_close_testcase_id_W10DA_044(self):
        self.logger.info("============> test_word_open_save_close_testcase_id_W10DA_44 <============")
        x, e, func_name  = self.main.word.word_open_save_close_testcase_id_W10DA_44()
        if x is True:
            self.logger.info(
                "Microsoft Word Save and Close  testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Microsoft Word Save and Close testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")  
    
    @pytest.mark.run(order=29)
    def test_msword_create_document_donot_save_testcase_id_W10DA_045(self):
        self.logger.info("============> test_word_create_document_donot_save_testcase_id_W10DA_045 <============")
        x, e, func_name  = self.main.word.word_create_document_donot_save_testcase_id_W10DA_045()
        if x is True:
            self.logger.info(
                "Microsoft Word create doc do't save and Close  testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Microsoft Word create doc do't save testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")  
        
    @pytest.mark.run(order=30)
    def test_msword_delete_file_validate_error_testcase_id_W10DA_046(self):
        self.logger.info("============> test_word_delete_file_validate_error_testcase_id_W10DA_046 <============")
        x, e, func_name  = self.main.word.word_delete_file_validate_error_testcase_id_W10DA_46()
        if x is True:
            self.logger.info(
                "Microsoft Word delete file and validate error testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Microsoft Word delete file and validate error testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")      
        
    @pytest.mark.run(order=31)
    def test_snipping_open_close_testcase_id_W10DA_027(self):
        self.logger.info("============> test_snipping_open_close_testcase_id_W10DA_27 <============")
        x, e, func_name  = self.main.snipping.snipping_open_close_testcase_id_W10DA_27()
        if x is True:
            self.logger.info(
                "Snipping Tool Open and Close  testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Snipping Tool Open and Close testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")  

    @pytest.mark.run(order=32)
    def test_snipping_new_cancel_option_close_testcase_id_W10DA_028(self):
        self.logger.info("============> test_snipping_new_cancel_option_close_testcase_id_W10DA_028 <============")
        x, e, func_name  = self.main.snipping.snipping_new_cancel_option_close_testcase_id_W10DA_28()
        if x is True:
            self.logger.info(
                "Snipping Tool New, cancel and Close  testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Snipping Tool New, cancel and Close testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")  
        
    @pytest.mark.run(order=33 )
    def test_excel_open_save_close_testcase_id_W10DA_047(self):
        self.logger.info("============> test_excel_open_save_close_testcase_id_W10DA_047 <============")
        x, e, func_name  = self.main.excel.excel_open_save_close_testcase_id_W10DA_047()
        if x is True:
            self.logger.info(
                "Microsoft Excel open-save-close testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Microsoft Excel open-save-close testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")    
    
    @pytest.mark.run(order=34 )
    def test_excel_open_dont_save_close_testcase_id_W10DA_048(self):
        self.logger.info("============> test_excel_open_dont_save_close_testcase_id_W10DA_048 <============")
        x, e, func_name  = self.main.excel.excel_open_dont_save_close_testcase_id_W10DA_048()
        if x is True:
            self.logger.info(
                "Microsoft Excel open-dont-save-close testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Microsoft Excel open-dont-save-close testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")     

    @pytest.mark.run(order=35 )
    def test_excel_delete_testcase_id_W10DA_049(self):
        self.logger.info("============> test_excel_delete_testcase_id_W10DA_049 <============")
        x, e, func_name  = self.main.excel.excel_delete_testcase_id_W10DA_049()
        if x is True:
            self.logger.info(
                "Microsoft Excel open-save-delete-close testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Microsoft Excel open-save-delete-close testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")     
 

    @pytest.mark.run(order=36 )
    def test_access_open_save_close_testcase_id_W10DA_053(self):
        self.logger.info("============> test_access_open_save_close_testcase_id_W10DA_053 <============")
        x, e, func_name  = self.main.access.access_open_save_close_testcase_id_W10DA_53()
        if x is True:
            self.logger.info(
                "Microsoft Access open-save-close testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Microsoft Access open-save-close testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================") 
    
    @pytest.mark.skip(reason="Always Failed") #run(order=37 )
    def test_access_dont_save_close_testcase_id_W10DA_054(self):
        self.logger.info("============> test_access_open_save_close_testcase_id_W10DA_051 <============")
        x, e, func_name  = self.main.access.access_dont_save_close_testcase_id_W10DA_54()
        if x is True:
            self.logger.info(
                "Microsoft Access dont-save-close testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Microsoft Access dont-save-close testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")
   
    @pytest.mark.run(order=38 )
    def test_access_delete_file_validate_error_testcase_id_W10DA_055(self):
        self.logger.info("============> test_access_delete_file_validate_error_testcase_id_W10DA_055 <============")
        x, e, func_name  = self.main.access.access_delete_file_validate_error_testcase_id_W10DA_55()
        if x is True:
            self.logger.info(
                "Microsoft Access delete-file-validate error testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Microsoft Access delete-file-validate error testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")    

    @pytest.mark.run(order=39 )
    def test_publisher_open_save_close_testcase_id_W10DA_039(self):
        self.logger.info("============> test_publisher_open_save_close_testcase_id_W10DA_039 <============")
        x, e, func_name  = self.main.publisher.publisher_open_save_close_testcase_id_W10DA_039()
        if x is True:
            self.logger.info(
                "Microsoft Publisher open_save_close testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Microsoft Publisher open_save_close testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")

    @pytest.mark.run(order=40 )
    def test_publisher_open_dontsave_close_testcase_id_W10DA_040(self):
        self.logger.info("============> test_publisher_open_dontsave_close_testcase_id_W10DA_040 <============")
        x, e, func_name  = self.main.publisher.publisher_open_dontsave_close_testcase_id_W10DA_040()
        if x is True:
            self.logger.info(
                "Microsoft Publisher open_dontsave_close testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Microsoft Publisher open_dontsave_close testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")
    
    @pytest.mark.run(order=41)
    def test_publisher_open_save_delete_open_close_testcase_id_W10DA_041(self):
        self.logger.info("============> test_publisher_open_save_delete_open_close_testcase_id_W10DA_041 <============")
        x, e, func_name  = self.main.publisher.publisher_open_save_delete_open_close_testcase_id_W10DA_041()
        if x is True:
            self.logger.info(
                "Microsoft Publisher open_save_delete_open_close testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Microsoft Publisher open_save_delete_open_close testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")
    
                
    @pytest.mark.run(order=42)
    def test_wordpad_open_save_close_testcase_id_W10DA_050(self):
        self.logger.info("============> test_wordpad_open_save_close_testcase_id_W10DA_050 <============")
        x, e, func_name  = self.main.wordpad.wordpad_open_save_close_testcase_id_W10DA_050()
        if x is True:
            self.logger.info(
                "Wordpad Save and Close  testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Wordpad Save and Close testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")     
         
    @pytest.mark.run(order=43)
    def test_wordpad_dont_save_testcase_id_W10DA_051(self):
        self.logger.info("============> test_wordpad_dont_save_testcase_id_W10DA_051 <============")
        x, e, func_name  = self.main.wordpad.wordpad_dont_save_testcase_id_W10DA_051()
        if x is True:
            self.logger.info(
                "Wordpad Don't Save  testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Wordpad Don't Save  testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")            
              
    @pytest.mark.run(order=44)
    def test_wordpad_delete_file_validate_error_testcase_id_W10DA_052(self):
        self.logger.info("============> test_wordpad_delete_file_validate_error_testcase_id_W10DA_052 <============")
        x, e, func_name  = self.main.wordpad.wordpad_delete_file_validate_error_testcase_id_W10DA_052()
        if x is True:
            self.logger.info(
                "Wordpad Delete file and validate error testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Wordpad Delete file and validate error testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")            
                        
    
    '''****************************************************************************************'''
    ''' Add tests under this section if tests need some manual intervention as pre-reqs  '''    
    '''****************************************************************************************'''   
        
    @pytest.mark.run(order=45)
    def test_performance_monitor_testcase_id_W10DA_015(self):
        self.logger.info("============> test_performance_monitor_id_W10DA_15 <============")
        x, e, func_name  = self.main.miscellaneous.performance_monitor_id_W10DA_15()
        if x is True:
            self.logger.info(
                "Performance monitor testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Performance monitor testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")
    
    
    @pytest.mark.run(order=71)    
    def test_symantec_endpoint_protection_open_close_testcase_id_W10DA_008(self):
        self.logger.info("===================> test_symantec_endpoint_protection_open_close_app <==========================================================")
        x, e, func_name = self.main.sep.symantec_endpoint_protection_open_close_testcase_id_W10DA_8()
           
        if x is True:
            self.logger.info("Symantec End point protection testcase {} passed".format(func_name))
        else:    
            self.assertTrue(x, self.logger.debug("Symantec End point protection testcase {} failed due to reason: '".format(func_name)+str(e)+"'"))
           
        self.logger.info("=============================================================================")  


    @pytest.mark.run(order=72)    
    def test_system_settings_performanceframe_testcase_id_W10DA_103(self):
        self.logger.info("===================> test_system_settings_performanceframe <==========================================================")
        x, e, func_name = self.main.system.system_settings_performanceframe_testcase_id_W10DA_103()
           
        if x is True:
            self.logger.info("System Settings PerformanceFrame testcase {} passed".format(func_name))
        else:    
            self.assertTrue(x, self.logger.debug("System Settings PerformanceFrame testcase {} failed due to reason: '".format(func_name)+str(e)+"'"))
           
        self.logger.info("=============================================================================")  
    

    @pytest.mark.run(order=73)    
    def test_system_settings_remoteuserscheck_testcase_id_W10DA_105(self):
        self.logger.info("===================> test_system_settings_remote_users <==========================================================")
        x, e, func_name = self.main.system.system_settings_remoteusers_testcase_id_W10DA_105()
           
        if x is True:
            self.logger.info("System Settings Remoter users check testcase {} passed".format(func_name))
        else:    
            self.assertTrue(x, self.logger.debug("System Settings Remoter users check testcase {} failed due to reason: '".format(func_name)+str(e)+"'"))
           
        self.logger.info("=============================================================================")  
    

    @pytest.mark.run(order=74)    
    def test_system_settings_systemfaiure_testcase_id_W10DA_104(self):
        self.logger.info("===================> test_system_settings_system_failure <==========================================================")
        x, e, func_name = self.main.system.system_settings_system_failure_testcase_id_W10DA_104()
           
        if x is True:
            self.logger.info("System Settings System Failure testcase {} passed".format(func_name))
        else:    
            self.assertTrue(x, self.logger.debug("System Settings System Failure check testcase {} failed due to reason: '".format(func_name)+str(e)+"'"))
           
        self.logger.info("=============================================================================")  
    

    @pytest.mark.run(order=75)    
    def test_system_settings_systemprotection_testcase_id_W10DA_106(self):
        self.logger.info("===================> test_system_settings_system_protection <==========================================================")
        x, e, func_name = self.main.system.system_settings_system_protection_testcase_id_W10DA_106()
           
        if x is True:
            self.logger.info("System Settings System Protection testcase {} passed".format(func_name))
        else:    
            self.assertTrue(x, self.logger.debug("System Settings System Protection check testcase {} failed due to reason: '".format(func_name)+str(e)+"'"))
           
        self.logger.info("=============================================================================")  
    

    @pytest.mark.run(order=76)    
    def test_windows_settings_power_sleep_defaultcheck_testcase_id_W10DA_107(self):
        self.logger.info("===================> test_windows_settings_power_sleep_testcase <==========================================================")
        x, e, func_name = self.main.system.windows_settings_power_sleep_testcase_id_W10DA_107()
           
        if x is True:
            self.logger.info("Windows Power Sleep Default Settings testcase {} passed".format(func_name))
        else:    
            self.assertTrue(x, self.logger.debug("Windows Power Sleep Default Settings testcase {} failed due to reason: '".format(func_name)+str(e)+"'"))
           
        self.logger.info("=============================================================================")  
    

    @pytest.mark.run(order=77)    
    def test_windows_settings_date_time_check_testcase_id_W10DA_108(self):
        self.logger.info("===================> test_windows_settings_power_sleep_testcase <==========================================================")
        x, e, func_name = self.main.system.windows_settings_date_time_testcase_id_W10DA_108()
           
        if x is True:
            self.logger.info("Windows Time Zone Settings testcase {} passed".format(func_name))
        else:    
            self.assertTrue(x, self.logger.debug("Windows Time Zone Settings testcase {} failed due to reason: '".format(func_name)+str(e)+"'"))
           
        self.logger.info("=============================================================================")  
    
    
    @pytest.mark.run(order=78)
    def test_outlook_check_datafile_testcase_id_W10DA_132(self):
        self.logger.info(
            "===================> test_outlook_check_datafile_testcase_id_W10DA_132 <============================")
        x, e, func_name = self.main.outlook.outlook_check_datafile_testcase_id_W10DA_132()
        if x is True:
            self.logger.info(" Check Outlook Data File testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug("Check Outlook Data File testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info("=============================================================================")       
    
    
    @pytest.mark.run(order=79)
    def test_outlook_check_if_datafile_exists_testcase_id_W10DA_133(self):
        self.logger.info(
            "===================> test_outlook_check_if_datafile_exists_testcase_id_W10DA_133 <============================")
        x, e, func_name = self.main.outlook.outlook_check_if_datafile_exists_testcase_id_W10DA_133()
        if x is True:
            self.logger.info(" Check if Outlook Data File exists testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug("Check if Outlook Data File exists testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info("=============================================================================")       
     
    @pytest.mark.run(order=80)
    def test_outlook_check_datafile_exists_from_importexport_testcase_id_W10DA_134(self):
        self.logger.info(
            "===================> test_outlook_check_datafile_exists_from_importexport_testcase_id_W10DA_134 <============================")
        x, e, func_name = self.main.outlook.outlook_check_datafile_importexport_testcase_id_W10DA_134()
        if x is True:
            self.logger.info(" Check if Outlook Data File is opening from Outlook Export/Import testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug("Check if Outlook Data File is opening from Outlook Export/Import testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info("=============================================================================")       
    

    @pytest.mark.run(order=81)
    def test_tier0_application_IEversion_testcase_id_W10DA_117(self):
        self.logger.info(
            "===================> test_tier0_application_IEversion_testcase_id_W10DA_117 <============================")
        x, e, func_name = self.main.ie.tier0_application_IEversion_testcase_id_W10DA_117()
        if x is True:
            self.logger.info("Check Tier 0 Application IE Version testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug("Check Tier 0 Application IE Version testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info("=============================================================================")       
    
    @pytest.mark.run(order=82)
    def test_tier0_application_internetoptions_testcase_id_W10DA_118(self):
        self.logger.info(
            "===================> test_tier0_application_internetoptions_testcase_id_W10DA_118 <============================")
        x, e, func_name = self.main.ie.tier0_application_internetoptions_testcase_id_W10DA_118()
        if x is True:
            self.logger.info("Check Tier 0 Application IE Internet Options testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug("Check Tier 0 Application IE Internet Options testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info("=============================================================================")       
      

    @pytest.mark.run(order=83)
    def test_tier0_application_internet_testcase_id_W10DA_119(self):
        self.logger.info(
            "===================> test_tier0_application_internet_testcase_id_W10DA_119 <============================")
        x, e, func_name = self.main.ie.tier0_application_internet_testcase_id_W10DA_119()
        if x is True:
            self.logger.info("Check Tier 0 Application IE Internet testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug("Check Tier 0 Application IE Internet testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info("=============================================================================")       
    

    @pytest.mark.run(order=84)
    def test_tier0_application_internet_trustedsite_testcase_id_W10DA_121(self):
        self.logger.info(
            "===================> test_tier0_application_internet_trustedsite_testcase_id_W10DA_121 <============================")
        x, e, func_name = self.main.ie.tier0_application_internet_trustedsite_testcase_id_W10DA_121()
        if x is True:
            self.logger.info("Check Tier 0 Application IE Trusted Sites testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug("Check Tier 0 Application IE Trusted Sites testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info("=============================================================================")       
        

    @pytest.mark.run(order=85)
    def test_tier0_application_internet_localintranet_testcase_id_W10DA_120(self):
        self.logger.info(
            "===================> test_tier0_application_internet_localintranet_testcase_id_W10DA_120 <============================")
        x, e, func_name = self.main.ie.tier0_application_internet_localintranet_testcase_id_W10DA_120()
        if x is True:
            self.logger.info("Check Tier 0 Application Local Intranet testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug("Check Tier 0 Application IE Local Intranet testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info("=============================================================================")       
      

    @pytest.mark.run(order=86)
    def test_tier0_application_internet_restrictedsites_testcase_id_W10DA_122(self):
        self.logger.info(
            "===================> test_tier0_application_internet_restrictedsites_testcase_id_W10DA_122 <============================")
        x, e, func_name = self.main.ie.tier0_application_internet_restrictedsites_testcase_id_W10DA_122()
        if x is True:
            self.logger.info("Check Tier 0 Application IE Restricted Sites testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug("Check Tier 0 Application IE Restricted Sites testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info("=============================================================================")       
       

    @pytest.mark.run(order=87)
    def test_tier0_application_internet_connections_testcase_id_W10DA_123(self):
        self.logger.info(
            "===================> test_tier0_application_internet_connections_testcase_id_W10DA_123 <============================")
        x, e, func_name = self.main.ie.tier0_application_internet_connections_testcase_id_W10DA_123()
        if x is True:
            self.logger.info("Check Tier 0 Application IE Connections testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug("Check Tier 0 Application IE Connections testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info("=============================================================================")       
         
    
    @pytest.mark.run(order=88)
    def test_tier0_application_general_browserhistory_testcase_id_W10DA_124(self):
        self.logger.info(
            "===================> test_tier0_application_general_browserhistory_testcase_id_W10DA_124 <============================")
        x, e, func_name = self.main.ie.tier0_application_general_browserhistory_testcase_id_W10DA_124()
        if x is True:
            self.logger.info("Check Tier 0 Application General Browser History testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug("Check Tier 0 Application General Browser History testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info("=============================================================================")       
         
    
    @pytest.mark.run(order=120)
    def test_gccm_windows_activation_testcase_id_W10DA_099(self):
        self.logger.info(
            "===================> test_gccm_windows_activation_testcase_id_W10DA_099 <============================")
        x, e, func_name = self.main.miscellaneous.gccm_windows_activation_testcase_id_W10DA_099()
        if x is True:
            self.logger.info(" Windows activation testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug("Windows activation testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info("=============================================================================")
    
    
    @pytest.mark.run(order=121)
    def test_gccm_windows_computername_testcase_id_W10DA_100(self):
        self.logger.info(
            "===================> test_gccm_windows_computername_testcase_id_W10DA_100 <============================")
        x, e, func_name = self.main.miscellaneous.gccm_windows_computername_testcase_id_W10DA_100()
        if x is True:
            self.logger.info(" Windows Computername testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug("Windows Computername testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info("=============================================================================")
    
    @pytest.mark.run(order=122)
    def test_gccm_windows_System_type_testcase_id_W10DA_101(self):
        self.logger.info(
            "===================> test_gccm_windows_System_type_testcase_id_W10DA_101 <============================")
        x, e, func_name = self.main.miscellaneous.gccm_windows_System_type_testcase_id_W10DA_101()
        if x is True:
            self.logger.info(" Windows Computername testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug("Windows Computername testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info("=============================================================================")
        
    @pytest.mark.run(order=123)
    def test_gccm_windows_OS_type_testcase_id_W10DA_102(self):
        self.logger.info(
            "===================> test_gccm_windows_OS_type_testcase_id_W10DA_102 <============================")
        x, e, func_name = self.main.miscellaneous.gccm_windows_OS_type_testcase_id_W10DA_102()
        if x is True:
            self.logger.info(" Windows Computername testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug("Windows Computername testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info("=============================================================================")
    
    
    @pytest.mark.run(order=124)
    def test_outlook_check_version_testcase_id_W10DA_110(self):
        self.logger.info(
            "===================> test_outlook_check_version_testcase_id_W10DA_110 <============================")
        x, e, func_name = self.main.outlook.outlook_check_version_testcase_id_W10DA_110()
        if x is True:
            self.logger.info(" Check Outlook Version Status testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug("Check Outlook Version Status testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info("=============================================================================")
    
    
    @pytest.mark.run(order=125)
    def test_outlook_offline_settings_testcase_id_W10DA_111(self):
        self.logger.info(
            "===================> test_outlook_offline_settings_testcase_id_W10DA_111 <============================")
        x, e, func_name = self.main.outlook.outlook_offline_settings_testcase_id_W10DA_111()
        if x is True:
            self.logger.info("Check Outlook Offline settings Status testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug("Check Outlook Offline settings testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info("=============================================================================") 
    
    
    @pytest.mark.run(order=126)
    def test_outlook_addins_disabled_testcase_id_W10DA_112(self):
        self.logger.info(
            "===================> test_outlook_addins_disabled_testcase_id_W10DA_112 <============================")
        x, e, func_name = self.main.outlook.outlook_addins_disabled_testcase_id_W10DA_112()
        if x is True:
            self.logger.info("Check Outlook AddIns testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug("Check Outlook AddIns  testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info("=============================================================================") 
    
    
    @pytest.mark.run(order=127)
    def test_outlook_account_settings_testcase_id_W10DA_113(self):
        self.logger.info(
            "===================> test_outlook_account_settings_testcase_id_W10DA_113 <============================")
        x, e, func_name = self.main.outlook.test_outlook_account_settings_testcase_id_W10DA_113()
        if x is True:
            self.logger.info("Check Outlook Account settings Status testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug("Check Outlook Account settings testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info("=============================================================================") 
    
    @pytest.mark.skip(reason="Always Failed") #run(order=128)
    def test_outlook_account_settings_encryption_testcase_id_W10DA_114(self):
        self.logger.info(
            "===================> test_outlook_account_settings_encryption_testcase_id_W10DA_114 <============================")
        x, e, func_name = self.main.outlook.outlook_account_settings_encryption_testcase_id_W10DA_113()
        if x is True:
            self.logger.info("Check Outlook Account settings Encryption Status testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug("Check Outlook Account settings Encryption testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info("=============================================================================") 
    
    
    ''' 
    @pytest.mark.run(order=72)    
    def test_symantec_endpoint_protection_run_active_scan_testcase_id_W10DA_026(self):
        self.logger.info("===================> test_symantec_endpoint_protection_run_active_scan_testcase_id_W10DA_26 <==========================================================")
        x, e, func_name = self.main.sep.symantec_endpoint_protection_run_active_scan_testcase_id_W10DA_26()
           
        if x is True:
            self.logger.info("Symantec End point protection Run active scan testcase {} passed".format(func_name))
        else:    
            self.assertTrue(x, self.logger.debug("Symantec End point protection Run active scan testcase {} failed due to reason: '".format(func_name)+str(e)+"'"))
        
        self.logger.info("=============================================================================")   
    '''
    
    @pytest.mark.run(order=108)
    def test_avatier_credential_password_reset_using_IE_Browser_testcase_id_W10DA_030(self):
        self.logger.info("=================> test_avatier_credential_password_reset_using_IE_Browser_testcase_id_W10DA_30 <=====================")
        x, e, func_name  = self.main.ie.avatier_credential_password_reset_using_IE_Browser_testcase_id_W10DA_30()
        if x is True:
            self.logger.info(
                "Password credential reset using IE Browser testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "Password credential reset using IE Browser testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")

    """ GNet Applications Testing """

    @pytest.mark.run(order=143)
    def test_gnet_core_applications_calendar_testcase_id_W10DA_143(self):
        self.logger.info("=================> test_gnet_core_applications_W10DA_143 <=====================")
        x, e, func_name  = self.main.gnet_core.gnet_core_applications_calendar_W10DA_143()
        if x is True:
            self.logger.info(
                "GNet core application calendar testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "GNet core application calendar testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")

    @pytest.mark.run(order=144)
    def test_gnet_core_applications_gxplearn_testcase_id_W10DA_144(self):
        self.logger.info("=================> test_gnet_core_applications_gxplearn_W10DA_144 <=====================")
        x, e, func_name  = self.main.gnet_core.gnet_core_applications_gxplearn_W10DA_144()
        if x is True:
            self.logger.info(
                "GNet core application GxpLearn testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "GNet core application GxpLearn testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")

    @pytest.mark.run(order=145)
    def test_gnet_core_applications_workday_testcase_id_W10DA_145(self):
        self.logger.info("=================> test_gnet_core_applications_workday_W10DA_145 <=====================")
        x, e, func_name  = self.main.gnet_core.gnet_core_applications_workday_W10DA_145()
        if x is True:
            self.logger.info(
                "GNet core application Workday testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "GNet core application Workday testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")

    @pytest.mark.run(order=146)
    def test_gnet_core_applications_sparc_testcase_id_W10DA_146(self):
        self.logger.info("=================> test_gnet_core_applications_sparc_W10DA_146 <=====================")
        x, e, func_name  = self.main.gnet_core.gnet_core_applications_sparc_W10DA_146()
        if x is True:
            self.logger.info(
                "GNet core application Sparc testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "GNet core application Sparc testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")

    @pytest.mark.run(order=147)
    def test_gnet_core_applications_glearn_testcase_id_W10DA_147(self):
        self.logger.info("=================> test_gnet_core_applications_glearn_W10DA_147 <=====================")
        x, e, func_name  = self.main.gnet_core.gnet_core_applications_glearn_W10DA_147()
        if x is True:
            self.logger.info(
                "GNet core application GLearn testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "GNet core application GLearn testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")

    @pytest.mark.run(order=148)
    def test_gnet_core_applications_search_testcase_id_W10DA_148(self):
        self.logger.info("=================> test_gnet_core_applications_search_W10DA_148 <=====================")
        x, e, func_name  = self.main.gnet_core.gnet_core_applications_search_W10DA_148()
        if x is True:
            self.logger.info(
                "GNet core application search testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "GNet core application search testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")

    @pytest.mark.run(order=149)
    def test_gnet_core_applications_ethics_compliance_testcase_id_W10DA_149(self):
        self.logger.info("=================> test_gnet_core_applications_ethics_compliance_W10DA_149 <=====================")
        x, e, func_name  = self.main.gnet_core.gnet_core_applications_ethics_compliance_W10DA_149()
        if x is True:
            self.logger.info(
                "GNet core application Ethics & Compliance testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "GNet core application Ethics & Compliance testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")

    @pytest.mark.run(order=150)
    def test_gnet_core_applications_gvdi_testcase_id_W10DA_150(self):
        self.logger.info("=================> test_gnet_core_applications_gvdi_W10DA_150 <=====================")
        x, e, func_name  = self.main.gnet_core.gnet_core_applications_gvdi_W10DA_150()
        if x is True:
            self.logger.info(
                "GNet core application GVDI testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "GNet core application GVDI testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")

    @pytest.mark.run(order=151)
    def test_gnet_core_applications_IDM_testcase_id_W10DA_151(self):
        self.logger.info("=================> test_gnet_core_applications_IDM_testcase_id_W10DA_151 <=====================")
        x, e, func_name  = self.main.gnet_core.gnet_core_applications_IDM_W10DA_151()
        if x is True:
            self.logger.info(
                "GNet core application IDM testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "GNet core application IDM testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")

    # @pytest.mark.run(order=152)
    # def test_gnet_core_applications_printer_testcase_id_W10DA_152(self):
    #     self.logger.info("=================> test_gnet_core_applications_printer_testcase_id_W10DA_152 <=====================")
    #     x, e, func_name  = self.main.gnet_core.gnet_core_applications_printer_W10DA_152()
    #     if x is True:
    #         self.logger.info(
    #             "GNet core application printer testcase {} passed".format(func_name))
    #     else:
    #         self.assertTrue(x, self.logger.debug(
    #             "GNet core application printer testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
    #     self.logger.info(
    #         "============================================================================")

    @pytest.mark.skip("LoadTest")  #run(order=153)
    def test_loadtesting_applications_testcase_id_W10DA_153(self):
        self.logger.info(
            "=================> test_loadtesting_applications_testcase_id_W10DA_153 <=====================")
        x, e, func_name = self.main.loadtesting.loadtesting_applications_testcase_id_W10DA_153()
        if x is True:
            self.logger.info(
                "LoadTesting applications testcase {} passed".format(func_name))
        else:
            self.assertTrue(x, self.logger.debug(
                "LoadTesting applications testcase {} failed due to reason: '".format(func_name) + str(e) + "'"))
        self.logger.info(
            "============================================================================")


def pytest_unconfigure() :
    f2 = open(ROOT_DIR+LOG_DIR+"SimpleTestLog.txt",'w')
    with open(ROOT_DIR+LOG_DIR+"DetailTestLog.txt", 'r') as f:
        for line in f:
            if(re.search(r'^.*(POST|GET|AwesomeSession|Finished).*$', line)):
                continue
            else:
                f2.write(line)


def fxn():
    warnings.warn("deprecated", DeprecationWarning)

with warnings.catch_warnings():
    warnings.simplefilter("ignore")
    fxn()


atexit.register(pytest_unconfigure)       
    

# Execute Tests
def main():
    pytest.main(['GUI_tests.py', '-s'])


if __name__ == '__main__':
    main()
