'''
Created on Mar 09, 2020
@author: jnguyen19
'''

import inspect
import os
import pywinauto

from pywinauto import MatchError
from definitions import LOADTESTING_TEST_DIR, LOADTESTING_MAIN_DIR, LOADTESTING_DATA_DIR, \
    NOTEPAD_TEST_DIR, ACROBAT_READER_TEST_DIR, ACROBAT_READER_MAIN_DIR, ACROBAT_READER_IMAGES_DIR, \
    OUTLOOK_MAIN_DIR, OUTLOOK_TEST_DIR, OUTLOOK_IMAGES_DIR, EXCEL_FILE_LOCATION1, EXCEL_FILE_LOCATION2, \
    EXCEL_BINARY_PATH, WORD_FILE_LOCATION, WORD_TEST_DIR, ROOT_DIR, LOOPING, LOADTEST_TXT_LOCATION, LOADTEST_WORD_LOCATION, \
    LOADTEST_PDF_LOCATION, LOADTEST_XLS_LOCATION, LOADTEST_PPT_LOCATION, cpu_mem_stats

from selenium.common.exceptions import NoSuchElementException, WebDriverException

from com.gilead.main.loadtesting.LoadTestingMainFile import LoadTestingMainClass
from definitions import email_recepient_id, email_content, email_subject
from com.gilead.main.outlook.OutlookMainFile import OutlookMainClass


class LoadTestingTestClass():
    __root_dir = ROOT_DIR
    __notepad_test_dir = NOTEPAD_TEST_DIR
    __word_test_dir = WORD_TEST_DIR

    __file_location1 = EXCEL_FILE_LOCATION1
    __file_location2 = EXCEL_FILE_LOCATION2


    __loadtesting_main_dir = LOADTESTING_MAIN_DIR
    __loadtesting_test_dir = LOADTESTING_TEST_DIR
    __loadtesting_data_dir = LOADTESTING_DATA_DIR

    _cpu_mem_stats = cpu_mem_stats

    def __init__(self, utilsRef, loggerRef):
        self.utils = utilsRef
        self.logger = loggerRef
        self.loadtesting_maincode = LoadTestingMainClass(utilsRef, loggerRef)


    '''
    This test will start X applications in series, then loops for LOOPING amount.
    It will start with Notepad, then Acrobat Reader then....
    It will track start and end time, memory and CPU usage after each app start.
    Will give summary of memory usage as well as CPU.
    '''
    def loadtesting_applications_testcase_id_W10DA_153(self):
        # Record start time & memory
        cpu_mem_stats['start_time'] = self.utils.getTimeCounterNow()
        cpu_mem_stats['mem_start'] = self.utils.getSystemAvailablePhysicalMemory()
        cpu_mem_stats['start_cpu'] = self.utils.getCPUAndMemoryStatistics()
        self.logger.info("LOADTEST: LT_START START_MEM : {}".format(cpu_mem_stats['mem_start']))
        self.logger.info("LOADTEST: LT_START START_CPU : {}".format(cpu_mem_stats['start_cpu']))
        self.logger.info("LOADTEST: LT_START START_TIME : {}".format(cpu_mem_stats['start_time']))

        #Log current Windows Version
        #(OSName, OSVersion) = self.utils.getWindowsOSBuildVersion()

        for loop in range(LOOPING):
            try:
                func = inspect.currentframe().f_back.f_code
                self.logger.info("Starting execution of function: {} :".format(func.co_name))
                self.logger.info("LOADTESTING- of APPs Loop: {} ".format(str(loop)))

                #LOAD UP FIRST APP: NOTEPAD
                self.logger.info("Starting Loop: {} : NotePad ".format(str(loop)))
                startKey = "app1_"+str(loop)+"_start_time"
                cpu_mem_stats[startKey] = self.utils.getTimeCounterNow()
                self.notePad_App_Launcher(loop)
                self.get_lapsed_time(startKey, "NotePad", 1, loop)


                #LOAD UP 2nd APP: ACROBAT READER
                self.logger.info("Starting Loop: {} : Acrobat Reader ".format(str(loop)))
                startKey = "app2_" + str(loop) + "_start_time"
                cpu_mem_stats[startKey] = self.utils.getTimeCounterNow()
                self.acrobat_App_Launcher(loop)
                self.get_lapsed_time(startKey, "Acrobat Reader", 2, loop)


                #LOAD UP 3rd APP:
                #Disabling for now due to intermittently not allowing multiple instances.
                self.logger.info("Starting Loop: {} : Outlook ".format(str(loop)))
                startKey = "app3_" + str(loop) + "_start_time"
                cpu_mem_stats[startKey] = self.utils.getTimeCounterNow()
                self.outlook_App_Launcher(loop)
                self.get_lapsed_time(startKey, "Outlook", 3, loop)


                # LOAD UP 4th APP:
                self.logger.info("Starting Loop: {} : Excel ".format(str(loop)))
                startKey = "app4_" + str(loop) + "_start_time"
                cpu_mem_stats[startKey] = self.utils.getTimeCounterNow()
                self.excel_App_Launcher(loop)
                self.get_lapsed_time(startKey, "Excel", 4, loop)


                # LOAD UP 5th APP:
                self.logger.info("Starting Loop: {} : Word  ".format(str(loop)))
                startKey = "app4_" + str(loop) + "_start_time"
                cpu_mem_stats[startKey] = self.utils.getTimeCounterNow()
                self.word_App_Launcher(loop)
                self.get_lapsed_time(startKey, "Word", 5, loop)

            except (AssertionError, AttributeError, TypeError, WebDriverException, NoSuchElementException) as e:
                return False, e, func.co_name


        return True, "", func.co_name

    def mem_cpu_stats(self, appCount, loop, message, summary=False):
        # Mem & Stuffs:
        self._cpu_mem_stats['mem_after_app'+str(appCount)+'_'+str(loop)] = self.utils.getSystemAvailablePhysicalMemory()
        self._cpu_mem_stats['end_time'] = self.utils.getTimeCounterNow()
        self._cpu_mem_stats['mem_end'] = self.utils.getSystemAvailablePhysicalMemory()
        cpu_mem_stats['end_cpu'] = self.utils.getCPUAndMemoryStatistics()

        # Display Mem & Time results:
        self.utils.displayTestProfileForTimeAndMemoryUsage(message, self._cpu_mem_stats, appCount, loop, summary)

    def get_lapsed_time(self, initialKey, appName, appCount, loop):
        endKey = 'app'+str(appCount)+'_' +str(loop)+'_end_time'
        cpu_mem_stats[endKey] = self.utils.getTimeCounterNow()
        lapsedTime = cpu_mem_stats[endKey] - cpu_mem_stats[initialKey]
        self.logger.info("LOADTEST: Lapsed Time for: {} : app : loop {} : is {} seconds".format(appName, loop, lapsedTime))
#***********************************************************************************************************************
    def notePad_App_Launcher(self, loop):
        self.logger.info("LOADTEST- Starting LOADTESTING of: Notepad App Loop: {} ".format(str(loop)))

        # TC Step 1  Check if existing process is running, clean it up.
        self.logger.info(
            "Checking and killing the process. if any before executing the test case")
        self.loadtesting_maincode.killProcess(self.loadtesting_maincode._LoadTestingMainClass__notepad)
        self.utils.removeFile(self.__root_dir + self.__notepad_test_dir + "SavedTextFile" + str(loop) + ".txt")

        # TC Step 2 - Load up a 500KB file
        self.loadtesting_maincode._LoadTestingMainClass__notepad.startNotepadProcess()
        assert self.loadtesting_maincode._LoadTestingMainClass__notepad.isNotepadProcessExist(), "Notepad.exe is not started"
        self.logger.info("Notepad process started successfully")
        self.logger.info("LOADTEST- Notepad about to load up a 500KB Text file...")

        self.loadtesting_maincode._LoadTestingMainClass__notepad.openFile(LOADTEST_TXT_LOCATION)

        # Mem & Stuffs:
        message = "*** End-of-App1 Notepad loop {} ***".format(str(loop))
        appCount = 1
        self.mem_cpu_stats(appCount, loop, message)
        return

    def acrobat_App_Launcher(self, loop):
        self.logger.info("LOADTEST- Starting LOADTESTING of: Acrobat Reader App Loop: {} ".format(str(loop)))

        '''start webdriver '''
        pid = self.utils.startWiniumDesktopDriver()
        self.logger.info("LOADTEST- Launched the WebDriver of AcrobatReader app... ")

        # TC Step 1
        # self.logger.info(
        #     "Checking and killing the process AcroRd32.exe if any before executing the test case")
        # self.loadtesting_maincode._LoadTestingMainClass__acrobatReader.killProcess()
        # self.utils.sleepUntil(5)

        self.loadtesting_maincode._LoadTestingMainClass__acrobatReader.startAcrobatReader_via_StartMenu()
        self.utils.sleepUntil(5)
        assert self.loadtesting_maincode._LoadTestingMainClass__acrobatReader.isAcrobatReaderProcessExist(), "Acrobat Reader process AcroRd32.exe didn't start properly"
        self.logger.info("Acrobat Reader app started and process AcroRd32.exe exist in taskmanager")

        ''' Connect the already running Acrobat Reader process to winium Driver instead of opening new one.'''
        self.loadtesting_maincode._LoadTestingMainClass__acrobatReader.connectAcrobatReaderToWebDriver()
        self.logger.info("connected Acrobat reader app to WebDriver ")

        #TC Step 3:
        assert self.loadtesting_maincode._LoadTestingMainClass__acrobatReader.isAcrobatReaderElementExist(
            self.loadtesting_maincode._LoadTestingMainClass__acrobatReader.getAcrobatReaderMainWindowMenuElement(),
            **{'Name': 'File'}), "Acrobat Reader app didn't open properly"
        self.logger.info("Checked the basic window elements present in Acrobat Reader GUI")

        # TC Step 3
        self.loadtesting_maincode._LoadTestingMainClass__acrobatReader.clickAcrobatReaderOpenMenuOption()
        self.logger.info("Clicked AcrobatReaderOpenMenuOption")
        self.utils.sleepUntil(3)

        # TC Step 4
        self.logger.info("LOADTEST- Acrobat Reader about to load up a 10MB PDF document...")
        self.loadtesting_maincode._LoadTestingMainClass__acrobatReader.enterFileNameAndClickToOpen(LOADTEST_PDF_LOCATION)

        assert self.loadtesting_maincode._LoadTestingMainClass__acrobatReader.isAcrobatReaderElementExist(self.loadtesting_maincode._LoadTestingMainClass__acrobatReader.driver, \
                                                                                                          **{'Name': '10MB_PDF_Document.pdf - Adobe Acrobat Reader DC'}), \
            "Acrobat reader 10MB_PDF_Document.pdf didnt open properly"
        self.logger.info("Able to find the 10MB_PDF_Document.pdf file element")

        pid.kill()

        # Mem & Stuffs:
        message = "*** End-of-App2 Acrobat loop {} ***".format(str(loop))
        appCount = 2
        self.mem_cpu_stats(appCount, loop, message)
        return

    def outlook_App_Launcher(self, loop):
        self.logger.info("LOADTEST- Starting LOADTESTING of: Outlook App Loop: {} ".format(str(loop)))

        '''start webdriver '''
        pid = self.utils.startWiniumDesktopDriver()
        self.utils.sleepUntil(5)
        self.logger.info("Launched the WebDriver of Outlook app... ")

        # TC Step 1
        #self.logger.info("Checking and killing the process OUTLOOK.exe if any before executing the test case")
        #self.loadtesting_maincode._LoadTestingMainClass__outlook.killProcess()

        # TC Step 2
        self.loadtesting_maincode._LoadTestingMainClass__outlook.openApplication()
        self.utils.sleepUntil(8)
        assert self.loadtesting_maincode._LoadTestingMainClass__outlook.isOutlookProcessExist(), "Outlook process OUTLOOK.EXE didn't exist"

        ''' Connect the already running Outlook process to winium Driver instead of opening new one.'''
        self.loadtesting_maincode._LoadTestingMainClass__outlook.connectOutlookToWebDriver()
        self.utils.sleepUntil(8)
        assert self.loadtesting_maincode._LoadTestingMainClass__outlook.isOutlookElementExist(
            self.loadtesting_maincode._LoadTestingMainClass__outlook.getOutlookMainWindowDoctopPaneElement(),
            **{'Name': 'New Email'}), "Outlook app didnt open properly"

        # TC Step 3
        # self.loadtesting_maincode._LoadTestingMainClass__outlook.clickNewItems()
        # self.utils.sleepUntil(5)
        #
        # # TC Step 4
        # self.loadtesting_maincode._LoadTestingMainClass__outlook.clickMoreItems()
        # self.utils.sleepUntil(5)
        #
        # # TC Step 5
        # self.loadtesting_maincode._LoadTestingMainClass__outlook.checkDataFile()
        # self.utils.sleepUntil(5)
        # self.logger.info(" No Outlook Data File in list")

        pid.kill()

        #Mem & Stuffs:
        message = "*** End-of-App3 Outlook loop {} ***".format(str(loop))
        appCount = 3
        self.mem_cpu_stats(appCount, loop, message)
        return

    def excel_App_Launcher(self, loop):

        self.logger.info("LOADTEST- Starting LOADTESTING of: Excel App Loop: {} ".format(str(loop)))

        '''start webdriver '''
        pid = self.utils.startWiniumDesktopDriver()
        self.logger.info("LOADTEST- Launched the WebDriver of MS-Excel app...{} ".format(pid))
        self.utils.sleepUntil(5)

        self.logger.info("LOADTEST- Excel about to load up 6MB XLSX file...")

        # TC Step 9
        handle = self.utils.openFile(LOADTEST_XLS_LOCATION)
        self.utils.sleepUntil(5)

        assert self.loadtesting_maincode._LoadTestingMainClass__excel.isExcelProcessExist(), "Excel process EXCEL.EXE didn't start properly"
        ''' Connect the already running Excel process to winium Driver instead of opening new one.'''
        self.loadtesting_maincode._LoadTestingMainClass__excel.connectExcelToWebDriver()
        self.logger.info("connected Excel app to WebDriver ")


        assert self.loadtesting_maincode._LoadTestingMainClass__excel.isExcelElementExist(
            self.loadtesting_maincode._LoadTestingMainClass__excel.driver, **{
                'ClassName': 'NetUIOfficeCaption'}), "Excel GUI Element is not present"
        self.logger.info("Checked the basic Excel elements present in Excel GUI")

        pid.kill()

        # Mem & Stuffs:
        message = "*** End-of-App4 Excel loop {} ***".format(str(loop))
        appCount = 4
        self.mem_cpu_stats(appCount, loop, message)
        return

    def word_App_Launcher(self, loop):

        self.logger.info("LOADTEST- Starting LOADTESTING of: Word App Loop: {} ".format(str(loop)))

        self.logger.info("LOADTEST- Word about to load up 2MB DOCX file...")

        #self.logger.info(
        #    "Checking and killing the process WINWORD.EXE if any before executing the test case")
        self.loadtesting_maincode._LoadTestingMainClass__word.killProcess()

        '''start webdriver '''
        pid = self.utils.startWiniumDesktopDriver()
        self.logger.info("Launched the Winium Driver for Word application. ")

        self.loadtesting_maincode._LoadTestingMainClass__word.createWordDocument(LOADTEST_WORD_LOCATION)
        self.utils.sleepUntil(3)

        assert self.utils.isFileExists(LOADTEST_WORD_LOCATION), "Word Doc {} is not created ".format(LOADTEST_WORD_LOCATION)
        self.logger.info("Microsoft Word is created at {}".format(LOADTEST_WORD_LOCATION))

        self.utils.openFile(LOADTEST_WORD_LOCATION)
        self.utils.sleepUntil(3)

        assert self.loadtesting_maincode._LoadTestingMainClass__word.isWordProcessExist(), "Microsoft Word WINWORD.EXE didn't start properly"
        ''' Connect the already running Word process to winium Driver instead of opening new one.'''

        self.loadtesting_maincode._LoadTestingMainClass__word.connectWordToWebDriver()
        self.logger.info("connected Word application to WebDriver")

        assert self.loadtesting_maincode._LoadTestingMainClass__word.isWordElementExist(self.loadtesting_maincode._LoadTestingMainClass__word.driver,
                                                     **{'Name': 'MsoDockTop'}), "Word Element is not present"
        self.logger.info("Checked the basic Word Doc elements present in Word GUI")

        pid.kill()

        # Mem & Stuffs:
        message = "*** End-of-App5 Word loop {} ***".format(str(loop))
        appCount = 5
        self.mem_cpu_stats(appCount, loop, message, True)
        return