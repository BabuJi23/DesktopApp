import pytest
import logging

from com.gilead.test.network_drive.network_main import NetworkMainClass
from definitions import LOCAL_DRIVE, SJ_DRIVE

class TestCase():
    
    main=NetworkMainClass()
    logger=logging
    
    @pytest.mark.run(order=1)
    @pytest.mark.parametrize( "src_drive, dest_drive, filename", [(LOCAL_DRIVE, SJ_DRIVE, "Test_txtfile_from_localdrive_546KB.txt"), 
                                                                  (LOCAL_DRIVE, SJ_DRIVE, "Test_docfile_from_localdrive_22KB.docx"),
                                                                  (LOCAL_DRIVE, SJ_DRIVE, "Test_pptfile_from_localdrive_5MB.ppt"),
                                                                  (SJ_DRIVE, LOCAL_DRIVE, "Test_pptfile_from_sanjose_drive_5MB.ppt")] 
                            )
    def test_copy_files_between_drives_testcase_id_W10DA_001(self, src_drive, dest_drive, filename):
        self.logger.info(
            "===================> test_copy_files_between_drives_testcase_id_W10DA_001 <============================")
        
        x, e, func_name = self.main.network.copy_files_between_drives_testcase_id_W10DA_001(src_drive, dest_drive, filename)
        if x is True:
            self.logger.info("Copy files to network share drive testcase {} passed".format(func_name))
        else:
            assert x, self.logger.debug("Copy files to network share drive testcase {} failed due to reason: '".format(func_name) + str(e) + "'")
        self.logger.info("=============================================================================")
    
    
    @pytest.mark.run(order=2)
    @pytest.mark.parametrize( "src_drive, dest_drive, filename", [(SJ_DRIVE, LOCAL_DRIVE, "Test_docfile_from_remote_drive_to_be_modified_47KB.docx"),(SJ_DRIVE, LOCAL_DRIVE, "Test_txtfile_from_remote_drive_to_be_modified_107KB.txt")] ) 
    def test_copy_file_to_localdrive_modify_copyback_to_remotedrive_testcase_id_W10DA_002(self, src_drive, dest_drive, filename):
        self.logger.info(
            "===================> test_copy_file_to_localdrive_modify_copyback_to_remotedrive_testcase_id_W10DA_002 <============================")
         
        x, e, func_name = self.main.network.copy_file_to_localdrive_modify_copyback_to_remotedrive_testcase_id_W10DA_002(src_drive, dest_drive, filename)
        if x is True:
            self.logger.info("Copy and modify files between drives testcase {} passed".format(func_name))
        else:
            assert x, self.logger.debug("Copy and modify files between drives testcas {} failed due to reason: '".format(func_name) + str(e) + "'")
        self.logger.info("=============================================================================")
         
        
    @pytest.mark.run(order=3)
    @pytest.mark.parametrize( "src_drive, dest_drive, filename1, filename2", [(LOCAL_DRIVE, SJ_DRIVE, "Test_txtfile1_from_localdrive.txt", "Test_txtfile2_from_localdrive.txt"), (LOCAL_DRIVE, SJ_DRIVE, "Test1_Global_Infra_Auto_Framework_Setup_9MB.docx", "Test2_Global_Infra_Auto_Framework_Setup_9MB.docx")] ) 
    def test_copy_multiplefiles_between_drives_testcase_id_W10DA_003(self, src_drive, dest_drive, filename1, filename2):
        self.logger.info(
            "===================> test_copy_multiplefiles_between_drives_testcase_id_W10DA_003 <============================")
         
        x, e, func_name = self.main.network.copy_files_between_drives_testcase_id_W10DA_001(src_drive, dest_drive, filename1, filename2)
        if x is True:
            self.logger.info("Copy and modify files between drives testcase {} passed".format(func_name))
        else:
            assert x, self.logger.debug("Copy and modify files between drives testcas {} failed due to reason: '".format(func_name) + str(e) + "'")
        self.logger.info("=============================================================================")
         
        
    
    
    
