'''
Created on Jul 26, 2018

@author: bkotamahanti
'''

import os

ROOT_DIR = os.path.dirname(os.path.realpath(__file__))
LOG_DIR = "\\com\\gilead\\test\\logs\\"
WINIUM_DRIVER = "\\com\\gilead\\main\\utils\\winium_driver\\Winium.Desktop.Driver.exe"

DATA_DIR = ROOT_DIR + "\\com\\gilead\\data\\"

'''Admin Tool properties '''
ADMIN_TOOL_MMC_PROCESS = "mmc.exe"
ADMIN_TOOL_ODBC_PROCESS = "odbcad32.exe"

'''Notepad properties '''
NOTEPAD_MAIN_DIR = "\\com\\gilead\\main\\notepad\\"
NOTEPAD_TEST_DIR = "\\com\\gilead\\test\\notepad\\"
NOTEPAD_PROCESS = "notepad.exe"

'''Calculator properties '''
CALCULATOR_MAIN_DIR = "\\com\\gilead\\main\\calculator\\"
CALCULATOR_TEST_DIR = "\\com\\gilead\\test\\calculator\\"
CALCULATOR_PROCESS_NAME = "Calculator.exe"
CALCULATOR_PROCESS_EXECUTABLE = "Calc.exe"

'''Google Support legacy directory file '''
GOOGLE_LEGACY_BROWSER_FILE = "C:\\Program Files (x86)\\Google\\Legacy Browser Support\\browser_switcher_bho.dll"

cal_dict = {    0   : {
                        'Name'  :  "Zero",
                        'AutomationId' : "num0Button",
                        'IsEnabled' :  True,
                        'ClassName' :  "Button" },
                1   : {
                        'Name'  :  "One",
                        'AutomationId' : "num1Button",
                        'IsEnabled' :   True,
                        'ClassName' :  "Button" },
                2   : {   
                        'Name'  :  "Two",
                        'AutomationId' : "num2Button",
                        'IsEnabled' :   True,
                        'ClassName' :   "Button" },
                3   : {
                        'Name'  :  "Three",
                        'AutomationId' : "num3Button",
                        'IsEnabled' :   True,
                        'ClassName' :   "Button" },
                4   : {
                        'Name'  :  "Four",
                        'AutomationId' : "num4Button",
                        'IsEnabled' :   True,
                        'ClassName' :   "Button"},
                5   : {
                        'Name'  :  "Five",
                        'AutomationId' : "num5Button",
                        'IsEnabled' :   True,
                        'ClassName' :   "Button"},
                6   : {
                        'Name'  :  "Six",
                        'AutomationId' : "num6Button",
                        'IsEnabled' :   True,
                        'ClassName' :   "Button"},
                7   : {
                        'Name'  :  "Seven",
                        'AutomationId' : "num7Button",
                        'IsEnabled' :   True,
                        'ClassName' :   "Button"},
                8   : {
                        'Name'  :  "Eight",
                        'AutomationId' : "num8Button",
                        'IsEnabled' :   True,
                        'ClassName' :   "Button"},
                9   : {
                        'Name'  :  "Nine",
                        'AutomationId' : "num9Button",
                        'IsEnabled' :   True,
                        'ClassName' :   "Button"}          
            }

'''LoadTesting properties'''
LOADTESTING_TEST_DIR = "\\com\\gilead\\test\\loadtesting\\"
LOADTESTING_MAIN_DIR = "\\com\\gilead\\main\\loadtesting\\"
LOADTESTING_DATA_DIR = "\\com\\gilead\\data\\"
LOADTEST_TXT_LOCATION = ROOT_DIR + LOADTESTING_DATA_DIR + "500KB_TXT_Document.txt"
LOADTEST_WORD_LOCATION = ROOT_DIR + LOADTESTING_DATA_DIR + "2MB_WORD_Document.docx"
LOADTEST_PDF_LOCATION = ROOT_DIR + LOADTESTING_DATA_DIR + "10MB_PDF_Document.pdf"
LOADTEST_XLS_LOCATION = ROOT_DIR + LOADTESTING_DATA_DIR + "6MB_XLS_Document.xlsx"
LOADTEST_PPT_LOCATION = ROOT_DIR + LOADTESTING_DATA_DIR + "1MB_PPT_Document.ppt"
LOOPING = 5

'''OneNote properties '''
ONENOTE_TEST_DIR = "\\com\\gilead\\test\\onenote\\"
ONENOTE_MAIN_DIR = "\\com\\gilead\\main\\onenote\\"
ONENOTE_PROCESS = "ONENOTE.EXE"
# ONENOTE_BINARY_PATH = "C:\\Program Files (x86)\\Microsoft Office\\Office16\\ONENOTE.EXE"
ONENOTE_BINARY_PATH = "ONENOTE.EXE"



'''Paint properties '''
PAINT_TEST_DIR="\\com\\gilead\\test\\paint\\"
PAINT_MAIN_DIR="\\com\\gilead\\main\\paint\\"
PAINT_PROCESS = "mspaint.exe" 

'''SEP properties '''
SEP_MAIN_DIR = "\\com\\gilead\\main\\sep\\"
SEP_TEST_DIR = "\\com\\gilead\\test\\sep\\"
SEP_PROCESS_NAME = "SymCorpUI.exe"
SEP_RUN_ACTIVE_SCAN_SUB_PROCESS = "SavUI.exe"

'''Outlook properties '''
OUTLOOK_TEST_DIR = "\\com\\gilead\\test\\outlook\\"
OUTLOOK_MAIN_DIR = "\\com\\gilead\\main\\outlook\\"
OUTLOOK_PROCESS = "OUTLOOK.EXE"
email_recepient_id = "joseph.nguyen19@gilead.com;nagendrababu.garimella@gilead.com;Igor.Shmakov@gilead.com"
email_subject = "This is test Subject"
email_content = "Hello World!! This is Sample message to the recipient"

'''AcrobatReader properties '''
ACROBAT_READER_MAIN_DIR = "\\com\\gilead\\main\\acrobat\\"
ACROBAT_READER_TEST_DIR = "\\com\\gilead\\test\\acrobat\\"
ACROBAT_READER_PROCESS = "AcroRd32.exe"
ACROBAT_READER_IMAGES_DIR ="\\com\\gilead\\main\\acrobat\\images\\"
ACROBAT_READER_TEST_FILE = DATA_DIR + "GIAT_test.pdf"

'''Google Chrome Browser properties '''
CHROME_BROWSER_MAIN_DIR = "\\com\\gilead\\main\\chrome_browser\\" 
CHROME_BROWSER_TEST_DIR = "\\com\\gilead\\test\\chrome_browser\\"
CHROME_BROWSER_PROCESS = "chrome.exe"
CHROME_DRIVER = "\\com\\gilead\\main\\utils\\chrome_driver\\chromedriver.exe"

''' websites '''
GILEAD_URL = "http://gilead.com"
GOOGLE_URL="http://google.com"
TRUVADA_URL="http://www.truvada.com"
AVATIER_URL="http://reset.gilead.com"

''' Microsoft Edge browser properties '''
EDGE_BROWSER_MAIN_DIR = "\\com\\gilead\\main\\edge_browser\\" 
EDGE_BROWSER_TEST_DIR = "\\com\\gilead\\test\\edge_browser\\"
EDGE_BROWSER_PROCESS = "MicrosoftEdge.exe"
EDGE_DRIVER_16999 = "\\com\\gilead\\main\\utils\\edge_driver\\MicrosoftWebDriver_16999.exe"
EDGE_DRIVER_15063 = "\\com\\gilead\\main\\utils\\edge_driver\\MicrosoftWebDriver_15063.exe"

'''Mozilla Firefox browser properties '''
FIREFOX_BROWSER_MAIN_DIR = "\\com\\gilead\\main\\firefox_browser\\" 
FIREFOX_BROWSER_TEST_DIR = "\\com\\gilead\\test\\firefox_browser\\"
FIREFOX_BROWSER_PROCESS = "firefox.exe"
FIREFOX_DRIVER = "\\com\\gilead\\main\\utils\\firefox_driver\\geckodriver.exe"
FIREFOX_DRIVER_ABS_PATH= ROOT_DIR+FIREFOX_DRIVER


'''MS Office Power point properties'''
POWERPOINT_MAIN_DIR = "\\com\\gilead\\main\\power_point\\" 
POWERPOINT_TEST_DIR = "\\com\\gilead\\test\\power_point\\"
POWERPOINT_PROCESS = "POWERPNT.EXE"
POWERPOINT_FILE_LOCATION1 = ROOT_DIR+POWERPOINT_TEST_DIR+"scriptCreatedPowerPoint.pptx"
POWERPOINT_FILE_LOCATION2 = ROOT_DIR+POWERPOINT_TEST_DIR+"samplePowerPoint2.pptx"
POWERPOINT_IMAGE_FILE = ROOT_DIR+POWERPOINT_MAIN_DIR+"image.jpg"

'''Internet Explorer browser properties '''
IE_BROWSER_PROCESS = "iexplore.exe"
# IE_DRIVER = "\\com\\gilead\\main\\utils\\ie_driver\\IEDriverServer.exe"
IE_DRIVER = "\\com\\gilead\\main\\utils\\ie_driver\\IEDriverServer3.4_64.exe"


'''Miscellaneous '''
PERFORMANCE_MONITOR_PROCESS = "mmc.exe"

'''Snipping Tool properties '''
SNIP_TOOL_PROCESS = "SnippingTool.exe" 

'''Microsoft Word properties'''
WORD_TEST_DIR = "\\com\\gilead\\test\\word\\"
WORD_PROCESS = "WINWORD.EXE"
WORD_FILE_LOCATION = ROOT_DIR+WORD_TEST_DIR+"mydoc1.docx"

'''Wordpad properties'''
WORDPAD_TEST_DIR = "\\com\\gilead\\test\\wordpad\\"
WORDPAD_PROCESS = "wordpad.exe"
WORDPAD_FILE_LOCATION = ROOT_DIR+WORDPAD_TEST_DIR
WORDPADFILE = "wordpaddoc1"

'''Microsoft Excel properties '''
EXCEL_TEST_DIR = "\\com\\gilead\\test\\excel\\"
EXCEL_MAIN_DIR = "\\com\\gilead\\main\\excel\\"
EXCEL_PROCESS = "EXCEL.EXE"
EXCEL_FILE_LOCATION1 = ROOT_DIR+EXCEL_TEST_DIR+"MSExcelDocument.xlsx"
EXCEL_FILE_LOCATION2 = ROOT_DIR+EXCEL_TEST_DIR+"MSExcelDocument2.xlsx"
# EXCEL_BINARY_PATH="C:\\Program Files (x86)\\Microsoft Office\\Office16\\EXCEL.EXE"
EXCEL_BINARY_PATH="EXCEL.EXE"

'''Microsoft Access DB properties'''
ACCESS_TEST_DIR = "\\com\\gilead\\test\\access\\"
ACCESS_PROCESS = "MSACCESS.EXE"
ACCESS_FILE_LOCATION = ROOT_DIR+ACCESS_TEST_DIR+"TestAccessDB.accdb"

'''Microsoft Publisher properties '''
PUBLISHER_BINARY_PATH="MSPUB.EXE"
PUBLISHER_PROCESS = "MSPUB.EXE"
PUBLISHER_TEST_DIR = "\\com\\gilead\\test\\publisher\\"
PUBLISHER_FILE_LOCATION1 = ROOT_DIR+PUBLISHER_TEST_DIR+"NewPublisher1.pub"
PUBLISHER_FILE_LOCATION2 = ROOT_DIR+PUBLISHER_TEST_DIR+"NewPublisher2.pub"
PUBLISHER_TEXT_TO_ENTER = "Hello World! this is sample Publisher text for demo. First Blank template of Microsoft Publisher App. Very useful to create themes like BirthdayCards, Thank you notes etc.,"


'''Device Manager properties '''
DEVICE_MANAGER_PROCESS = "mmc.exe"
DEVICE_MANAGER_BINARY_PATH="C:\\Windows\\System32\\devmgmt.msc"
DEVICE_MANAGER_IMAGES_DIR="\\com\\gilead\\main\\devicemanager\\images\\" 

''' System Settings '''

SYSTEM_IMAGES_DIR="\\com\\gilead\\main\\system\\images\\" 

''' Outlook Data File '''
OUTLOOK_IMAGES_DIR="\\com\\gilead\\main\\outlook\\images\\"

IE_IMAGES_DIR="\\com\\gilead\\main\\ie_browser\\images\\"

''' GCCM Settings'''
MISC_IMAGES_DIR="\\com\\gilead\\main\\miscellaneous\\images\\"

''' GNet Core Application '''
GNET_IMAGES_DIR="\\com\\gilead\\main\\gnet_core\\images\\"

''' Global Cluster drives'''
LOCAL_DRIVE=ROOT_DIR+"\\docs\\"
SJ_DRIVE="\\\\fcgcsnsvmn05\\GIAT"
# SJ_DRIVE="\\\\fcgcsnsvmn05\\laboperations"
CAMBRIDGE_DRIVE="\\\\cbgcsnsvmg01\\GIAT"
CORK_DRIVE="\\\\eugcsnprdclg01\\GIAT"
HONGKONG_DRIVE="\\\\hggcsnsvmg01\\GIAT"
TOKYO_DRIVE="\\\\tygcsnsvmg01\\GIAT"

'''CPU & Memory Usage Statistics'''
cpu_mem_stats = {
    #Start x is fixed, end_x is transient.
    'start_time'        : "",
    'end_time'          : "",
    'mem_start'         : "",
    'mem_end'           : "",
    'start_cpu'         : "",
    'end_cpu'           : "",

    #Transient during run
    'cummulative_mem_usage' : "",
    'cummulative_cpu_usage' : "",
    'cummulative_time_elapsed' : "",

    # Overall result
    'overall_mem_usage': "",
    'overall_cpu_usage': "",
    'overall_time_elapsed': "",

    #Mem Per App
    'mem_after_app1_0'   : "",
    'mem_after_app2_0'   : "",
    'mem_after_app3_0'   : "",
    'mem_after_app4_0'   : "",
    'mem_after_app5_0'   : "",
    'mem_after_app1_1'   : "",
    'mem_after_app2_1'   : "",
    'mem_after_app3_1'   : "",
    'mem_after_app4_1'   : "",
    'mem_after_app5_1'   : "",
    'mem_after_app1_2'   : "",
    'mem_after_app2_2'   : "",
    'mem_after_app3_2'   : "",
    'mem_after_app4_2'   : "",
    'mem_after_app5_2'   : "",
    'mem_after_app1_3': "",
    'mem_after_app2_3': "",
    'mem_after_app3_3': "",
    'mem_after_app4_3': "",
    'mem_after_app5_3': "",
    'mem_after_app1_4': "",
    'mem_after_app2_4': "",
    'mem_after_app3_4': "",
    'mem_after_app4_4': "",
    'mem_after_app5_4': "",

    #Time per App:
    'app1_0_start_time': "",
    'app1_0_end_time': "",
    'app1_1_start_time': "",
    'app1_1_end_time': "",
    'app1_2_start_time': "",
    'app1_2_end_time': "",
    'app1_3_start_time': "",
    'app1_3_end_time': "",
    'app1_4_start_time': "",
    'app1_4_end_time': "",

    'app2_0_start_time': "",
    'app2_0_end_time': "",
    'app2_1_start_time': "",
    'app2_1_end_time': "",
    'app2_2_start_time': "",
    'app2_2_end_time': "",
    'app2_3_start_time': "",
    'app2_3_end_time': "",
    'app2_4_start_time': "",
    'app2_4_end_time': "",

    'app3_0_start_time': "",
    'app3_0_end_time': "",
    'app3_1_start_time': "",
    'app3_1_end_time': "",
    'app3_2_start_time': "",
    'app3_2_end_time': "",
    'app3_3_start_time': "",
    'app3_3_end_time': "",
    'app3_4_start_time': "",
    'app3_4_end_time': "",

    'app4_0_start_time': "",
    'app4_0_end_time': "",
    'app4_1_start_time': "",
    'app4_1_end_time': "",
    'app4_2_start_time': "",
    'app4_2_end_time': "",
    'app4_3_start_time': "",
    'app4_3_end_time': "",
    'app4_4_start_time': "",
    'app4_4_end_time': "",

    'app5_0_start_time': "",
    'app5_0_end_time': "",
    'app5_1_start_time': "",
    'app5_1_end_time': "",
    'app5_2_start_time': "",
    'app5_2_end_time': "",
    'app5_3_start_time': "",
    'app5_3_end_time': "",
    'app5_4_start_time': "",
    'app5_4_end_time': "",
    }

app_list = {
    'NotePad' : 1,
    'Acrobat Reader' : 2,
    'Outlook'   :3,
    'Excel' : 4,
    'Word'  :5,
}