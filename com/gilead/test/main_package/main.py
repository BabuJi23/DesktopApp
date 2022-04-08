'''
Created on Aug 17, 2018
@author: bkotamahanti
'''

''' built-in modules'''
import os
import logging

'''user-defined modules'''
from com.gilead.test.notepad.NotepadTestFile import NotepadTestClass
from com.gilead.main.utils.DesktopAppUtilsFile import DesktopAppCommonFunctionsClass
from definitions import ROOT_DIR, LOG_DIR
from com.gilead.test.calculator.CalculatorTestFile import CalculatorTestClass
from com.gilead.test.paint.PaintTestFile import PaintTestClass
from com.gilead.test.onenote.OneNoteTestFile import OneNoteTestClass
from com.gilead.test.outlook.OutlookTestFile import OutlookTestClass
from com.gilead.test.sep.SEPTestFile import SEPTestClass
from com.gilead.test.acrobat.AcrobatReaderTestFile import AcrobatReaderTestClass
from com.gilead.test.chrome_browser.ChromeBrowserTestFile import ChromeBrowserTestClass
from com.gilead.test.miscellaneous.MiscellaneousTestFile import MiscellaneousTestClass
from com.gilead.test.edge_browser.EdgeBrowserTestFile import EdgeBrowserTestClass
from com.gilead.test.firefox_browser.FirefoxBrowserTestFile import FirefoxBrowserTestClass
from com.gilead.test.power_point.PowerPointTestFile import PowerPointTestClass
from com.gilead.test.ie_browser.InternetExplorerTestFile import InternetExplorerTestClass
from com.gilead.test.word.WordTestFile import WordTestClass
from com.gilead.test.snipping.SnippingTestFile import SnippingTestClass
from com.gilead.test.excel.ExcelTestFile import ExcelTestClass
from com.gilead.test.access.AccessTestFile import AccessTestClass
from com.gilead.test.wordpad.WordpadTestFile import WordpadTestClass
from com.gilead.test.publisher.PublisherTestFile import PublisherTestClass
from com.gilead.test.admintool.AdminToolTestFile import AdminToolTestClass
from com.gilead.test.system.SystemsettingsTestFile import SystemsettingsTestClass
from com.gilead.test.gnet_core.gnetcoreTestFile import gnetcoreTestClass
from com.gilead.test.devicemanager.DeviceManagerTestFile import DeviceManagerTestClass
from com.gilead.test.loadtesting.LoadTestingTestFile import LoadTestingTestClass

class MainClass():
    def __init__(self):
        '''public variables'''
        
        self.detailLogFile = ROOT_DIR + LOG_DIR + "DetailTestLog.txt"
        ''' Checking and removing the log file before opening and writing to it...'''
        if (os.path.exists(self.detailLogFile)):
            os.remove(self.detailLogFile) 
        
        self.simpleLogFile = ROOT_DIR + LOG_DIR + "SimpleTestLog.txt"
        
        ''' Checking and removing the log file before opening and writing to it...'''
        if (os.path.exists(self.simpleLogFile)):
            os.remove(self.simpleLogFile) 
          
        self.logger = logging.getLogger("root")
        for handler in logging.root.handlers[:]:
            logging.root.removeHandler(handler)
        
        logging.basicConfig(filename = self.detailLogFile, 
                                level=logging.DEBUG, 
                                format='%(asctime)s : %(levelname)s :  %(filename)s :  %(funcName)s()  %(message)s'
#                                 format='[%(asctime)s : %(levelname)s : filename ( %(filename)s ): line number ( %(lineno)s )-------function name ( %(funcName)s()) ] %(message)s' 
                                )
        ''' private variables'''
        self.__num1 = 0
        self.__num2 = 0
      
        ''' object initialization '''
        self.utils = DesktopAppCommonFunctionsClass()
        self.notepad = NotepadTestClass(self.utils, self.logger)
        self.calculator = CalculatorTestClass(self.utils, self.logger)
        self.paint = PaintTestClass(self.utils, self.logger)
        self.onenote = OneNoteTestClass(self.utils, self.logger)
        self.outlook = OutlookTestClass(self.utils, self.logger)
        self.sep = SEPTestClass(self.utils, self.logger)
        self.acrobat = AcrobatReaderTestClass(self.utils, self.logger)
        self.chrome = ChromeBrowserTestClass(self.utils, self.logger)
        self.edge = EdgeBrowserTestClass(self.utils, self.logger)
        self.firefox = FirefoxBrowserTestClass(self.utils, self.logger)
        self.miscellaneous = MiscellaneousTestClass(self.utils, self.logger)
        self.powerpoint = PowerPointTestClass(self.utils, self.logger)
        self.ie = InternetExplorerTestClass(self.utils, self.logger)
        self.word = WordTestClass(self.utils, self.logger)
        self.snipping = SnippingTestClass(self.utils, self.logger)
        self.excel = ExcelTestClass(self.utils, self.logger)
        self.access = AccessTestClass(self.utils, self.logger)
        self.wordpad = WordpadTestClass(self.utils, self.logger)
        self.publisher = PublisherTestClass(self.utils, self.logger)
        self.admintool = AdminToolTestClass(self.utils, self.logger)
        self.devicemanager = DeviceManagerTestClass(self.utils, self.logger)
        self.system = SystemsettingsTestClass(self.utils, self.logger)
        self.gnet_core = gnetcoreTestClass(self.utils, self.logger)
        self.loadtesting = LoadTestingTestClass(self.utils, self.logger)
        
        
        ''' Kill the webdriver process if any before being used by test cases'''
        self.utils.killProcess("Winium.Desktop.Driver.exe")


