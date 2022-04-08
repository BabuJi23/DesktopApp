'''
Created on Mar 09, 2020
@author: jnguyen19
'''

from com.gilead.main.word.WordMainFile import WordMainClass
from com.gilead.main.excel.ExcelMainFile import ExcelMainClass
from com.gilead.main.outlook.OutlookMainFile import OutlookMainClass
from com.gilead.main.acrobat.AcrobatReaderMainFile import AcrobatReaderMainClass
from com.gilead.main.notepad.NotepadMainFile import NotepadMainClass
from com.gilead.main.utils.DesktopAppUtilsFile import DesktopAppCommonFunctionsClass

from _overlapped import NULL
from definitions import LOADTESTING_TEST_DIR, LOADTESTING_MAIN_DIR, LOADTESTING_DATA_DIR, \
    NOTEPAD_MAIN_DIR, NOTEPAD_PROCESS, ACROBAT_READER_PROCESS, ACROBAT_READER_MAIN_DIR, \
    ACROBAT_READER_IMAGES_DIR, OUTLOOK_MAIN_DIR, OUTLOOK_PROCESS, EXCEL_PROCESS, EXCEL_MAIN_DIR, \
    WORD_PROCESS, ROOT_DIR

from selenium import webdriver
from docx import Document
import os
import pywinauto
import pyautogui
from selenium.common.exceptions import WebDriverException

class LoadTestingMainClass():
    __process = ""
    __processNotePad = NOTEPAD_PROCESS
    __processAcrobat = ACROBAT_READER_PROCESS
    __processOutlook = OUTLOOK_PROCESS
    __processExcel = EXCEL_PROCESS
    __processWord = WORD_PROCESS

    #Class attribute to hold instances of started apps.
    __word = ""
    __excel = ""
    __outlook = ""
    __acrobatReader = ""
    __notepad = ""

    __notepad_main_dir = NOTEPAD_MAIN_DIR
    __acrobat_main_dir = ACROBAT_READER_MAIN_DIR
    __image_dir = ROOT_DIR + ACROBAT_READER_IMAGES_DIR
    __outlook_main_dir = ROOT_DIR + OUTLOOK_MAIN_DIR
    __excel_main_dir = ROOT_DIR + EXCEL_MAIN_DIR
    __loadtesting_main_dir = LOADTESTING_MAIN_DIR


    def __init__(self, utilsRef, loggerRef):
        self.app = pywinauto.application.Application()
        self.utils = utilsRef
        self.logger = loggerRef
        self.driver = NULL
        self.__notepad = NotepadMainClass(utilsRef, loggerRef)
        self.__acrobatReader = AcrobatReaderMainClass(utilsRef, loggerRef)
        self.__outlook = OutlookMainClass(utilsRef, loggerRef)
        self.__excel = ExcelMainClass(utilsRef, loggerRef)
        self.__word = WordMainClass(utilsRef, loggerRef)


    def startAppOpenFile(appName, file2Open):
        def app_process(appName):
            switcher = {
                         'Word': WORD_PROCESS,
                        'Excel': EXCEL_PROCESS,
                'AcrobatReader': ACROBAT_READER_PROCESS,
                      'Notepad': NOTEPAD_PROCESS
            }
            return switcher.get(appName, "Invalid application Name provided!")

    def killProcess(self, appProcess):
        self.utils.killProcess(appProcess)

    def isProcessExist(self, appProcess):
        return self.utils.isProcessExist(appProcess)