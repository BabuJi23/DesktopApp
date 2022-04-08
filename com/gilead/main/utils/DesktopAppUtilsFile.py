'''
Created on Jul 23, 2018

@author: bkotamahanti
'''
import psutil
import pyautogui
import time
import logging
import inspect
import subprocess
from definitions import ROOT_DIR, WINIUM_DRIVER, cpu_mem_stats
from _overlapped import NULL
import os
import shutil
import win32api
import sys
import wmi
comp = wmi.WMI()
from datetime import datetime

class DesktopAppCommonFunctionsClass(object):
    wait_time = 4
    __root_dir = ROOT_DIR
    __winium_driver = WINIUM_DRIVER
    time_counter = 0
    
    def __init__(self):
        self.logger = logging
        self.driver = NULL
        self.time_counter = time.perf_counter()
    
    def killProcess(self, process_name):
        flag = 0
        for proc in psutil.process_iter():
            # check whether the process_name name matches
            if proc.name() == process_name:
                flag = 1
                self.logger.info("Found the process "+process_name+ " in process list and killing it now")
                proc.kill()
                
                if self.isProcessExist(process_name):
                    self.logger.info("waiting for some more time until the process is completely killed")
                    self.sleepUntil(5)
                self.logger.info("Killed the process "+process_name)
        
        if flag == 0:
            self.logger.info("There is no process {} in process list".format(process_name))
              
                
    def startWiniumDesktopDriver(self):
#         subprocess._cleanup()
        self.killProcess(self.__root_dir+self.__winium_driver)
        self.sleepUntil(5)
        return subprocess.Popen(self.__root_dir+self.__winium_driver)         
    
    def openApplication(self, filepath ):
        os.startfile(filepath, "open")
        self.sleepUntil(3)
        self.logger.info("started application {}".format(filepath))
        
    def pageDown(self):
        for _ in range(5):
            pyautogui.press('down',2)
    
    def executeSystemCommands(self, command ):
        os.system(command)
        self.sleepUntil(4)
        self.logger.info("Executed command  {}".format(command))
    
    def isFileExists(self, filename):
        if os.path.exists(filename):
            self.logger.info("File "+str(filename)+" exists")
            return True
        else:
            self.logger.info("File "+str(filename)+" doesn't exist")
            return False 
                 
    def openFile(self, fileName):
        handle = win32api.ShellExecute(0, 'open', fileName, '', '', 1)
        return handle
    
    def removeFile(self, filename):
        if os.path.exists(filename):
            os.remove(filename)
            self.logger.info("File {} is deleted successfully ".format(filename))
        
    def removeDir(self, filename):
        if os.path.exists(filename):
            shutil.rmtree(filename)
    
    def isElementPresent(self, driver_element, **kargs):
        if 'Name' in kargs:
            list_of_webelements= driver_element.find_elements_by_name(kargs['Name'])
        if 'ClassName' in kargs:
            list_of_webelements= driver_element.find_elements_by_class_name(kargs['ClassName'])
        
#         print("---->number of elements {}".format(len(list_of_webelements)))
        return len(list_of_webelements) > 0
    
    def isElementTextPresent(self, driver_element, text, **kargs):
        if 'Name' in kargs:
            list_of_webelements= driver_element.find_elements_by_name(kargs['Name'])
        if 'ClassName' in kargs:
            list_of_webelements= driver_element.find_elements_by_class_name(kargs['ClassName'])
        
        for element in list_of_webelements:
            if text in element.get_attribute("Name"):
                return True
            
        return False
    
    def getOSBuild(self):
        return str(sys.getwindowsversion().build)
    
    def isProcessExist(self,process_name):
        for proc in psutil.process_iter():
            # check whether the process_name name matches
            if proc.name() == process_name:
                self.logger.info("process "+str(process_name)+" exists")
                return True
        self.logger.info("process "+str(process_name)+" doesn't exist ")
        return False
    
    def startProgramByName_via_StartMenu(self, progName):
        pyautogui.press("winleft")
        self.sleepUntil(self.wait_time)
        pyautogui.typewrite(progName, interval=0.25)
        pyautogui.press("enter")
        self.sleepUntil(self.wait_time)
        self.logger.info("Started program "+str(progName)+" via Windows start icon")
        return True
    
    def getCoordinatesByLocatingGivenImageOnScreen(self,imageFile, minSearchTime):
        r=pyautogui.locateOnScreen(imageFile, minSearchTime)
        self.logger.info("image is located at coordinates on screen ==> "+str(r)+" <==")
        time.sleep(self.wait_time)
        a= r[0] 
        b= r[1]
        self.logger.info("returning coordinates of image==> "+str(imageFile) + " (" +str(a)+","+str(b)+") <==")
        return a,b
    
    def moveToLocationAndClick(self, x_coor, y_coor):
        self.logger.info("moving to location ( {} {} ) on application screen:".format(x_coor, y_coor))
        pyautogui.moveTo(x_coor , y_coor)
        self.logger.info("clicking on location ( {} {} ) of application screen:".format(x_coor, y_coor))
        pyautogui.click(x_coor, y_coor, clicks=2)
    
    def isAppStarted(self, imageFile, searchTime):
        r = pyautogui.locateOnScreen(imageFile, 20)
        if(r is not None):
            self.logger.info("able to locate image {} on the application".format(imageFile))
            return True
        self.logger = self.logger.getLogger()
        self.logger.debug("not able to locate image {} on the application".format(imageFile))
        return False
    
    def sleepUntil(self, wait):
        self.logger.info("waiting for {} seconds for application to sync up or the element to show up".format(wait))
        time.sleep(wait)
    
    def close_window(self):
        pyautogui.hotkey('alt', 'f4')
        self.logger.info("Closed the window with alt+f4 key")
    
    def getSystemHostName(self):
        import socket
        return socket.getfqdn()
    
    def log_assert(self, bool_, message="", logger=None, logger_name="", verbose=False ):
        """Use this as a replacement for assert if you want the failing of the
    assert statement to be logged."""
        if logger is None:
            logger = logging.getLogger(logger_name)
        try:
            assert bool_, message
        except AssertionError as e:
            last_stackframe = inspect.stack()[-2]
            source_file, line_no, func = last_stackframe[1:4]
            source = "Traceback (most recent call last):\n" + \
                    '  File "%s", line %s, in %s\n    ' % (source_file, line_no, func)
            if verbose:
                # include more lines than that where the statement was made
                source_code = open(source_file).readlines()
                source += "".join(source_code[line_no - 3:line_no + 1])
            else:
                source += last_stackframe[-2][0].strip()
            logger.debug("%s\n%s" % (message+":"+str(e), source))
            raise AssertionError("%s\n%s" % (message, source))

    def getSystemTotalPhysicalMemory(self):
        for i in comp.Win32_ComputerSystem():
            totalPhysicalMemory = int(i.TotalPhysicalMemory) / 1024000
            print(totalPhysicalMemory, "Total Physical Memory in MB")
            return totalPhysicalMemory

    def getSystemAvailablePhysicalMemory(self):
        for os in comp.Win32_OperatingSystem():
            availablePhysicalMemory = int(os.FreePhysicalMemory) / 1024
            print(availablePhysicalMemory, "Available Physical Memory in MB")
            return availablePhysicalMemory

    def getCPUAndMemoryStatistics(self):
        CPU_Usage = psutil.cpu_percent(interval=1, percpu=False)
        CPU_Count = psutil.cpu_count()
        CPU_Usage_Overall = CPU_Usage * CPU_Count
        #self.logger.info('LOADTEST: CPU % Usage Overall = {}'.format(CPU_Usage_Overall))
        #self.logger.info('LOADTEST: {}'.format(psutil.virtual_memory()))  # physical memory usage
        #self.logger.info('LOADTEST: memory % used:'.format(psutil.virtual_memory()[2]))
        return CPU_Usage_Overall

    '''
    @usage: retrieve the current time counter, which can be saved off,
    future readings and taking the differences will give the time for the X process, you're trying to track.
    @return: current time counter now
    '''
    def getTimeCounterNow(self):
        time_counter = time.perf_counter()
        return time_counter

    @staticmethod
    def getWindowsOSBuildVersion():
        cmd = "systeminfo | findstr /b /c:\"OS Name\" /c:\"OS Version\" "
        WinBuildVer = subprocess.check_output(cmd,  shell=True).strip()
        attributes = WinBuildVer.split(b'\r\n')
        OSName = (attributes[0].split(b':')[1].strip()).decode('ascii')
        OSVersion = (attributes[1].split(b':')[1].strip()).decode('ascii')
        return (OSName, OSVersion)

    @staticmethod
    def getPythonVersion():
        cmd = "python -V"
        pythonversion = subprocess.check_output(cmd,  shell=True).strip()
        return (pythonversion).decode('ascii')

    @staticmethod
    def getDateTimeString():
        now = datetime.now()
        return(now.strftime("%m-%d-%y-%H-%M-%S"))

    @staticmethod
    def getDate():
        now = datetime.now()
        return(now.strftime("%m-%d-%y"))

    def displayTestProfileForTimeAndMemoryUsage(self, message, cpu_mem_stats, appCount, loop, summary=False):
        processTime =  cpu_mem_stats['end_time'] - cpu_mem_stats['start_time']
        processMemUsage = cpu_mem_stats['mem_start'] - cpu_mem_stats['mem_end']
        processCPUUsage = cpu_mem_stats['end_cpu'] #Taking Curret CPU count, not delta.- cpu_mem_stats['start_cpu']

        self.logger.info('LOADTEST: '+message)
        self.logger.info('LOADTEST: MemStart: {}: MemEnd: {} : TimeStart: {} : TimeEnd: {}\n' \
                         .format(cpu_mem_stats['mem_start'], cpu_mem_stats['mem_end'],\
                                 cpu_mem_stats['start_time'], cpu_mem_stats['end_time'],
                                 cpu_mem_stats['start_cpu'], cpu_mem_stats['end_cpu']))
        self.logger.info('LOADTEST: Cummulative Processing Time for app: {} : loop: {} : is: {} : in secs.'.format(appCount, loop, processTime))
        self.logger.info('LOADTEST: Cummulative Memory Usage  for app: {} : loop: {} : is: {} : in MB'.format(appCount, loop, processMemUsage))
        self.logger.info('LOADTEST: Cummulative CPU Usage  for app: {} : loop: {} : is: {} : in %.'.format(appCount, loop, processCPUUsage))


        self.logger.info('\n')
        for key in cpu_mem_stats:
            if cpu_mem_stats[key]:
                self.logger.info("LOADTEST: CPU_MEM_STATS For:  \'{}\':\'{}\' ".format(key, cpu_mem_stats[key]))
        self.logger.info('\n')

        if summary:
            self.logger.info("LOADTEST: OVERALL-MEM-USAGE: {} : (in MB)".format(processMemUsage))
            self.logger.info("LOADTEST: OVERALL-CPU-USAGE: {} : (in %)".format(processCPUUsage))
            self.logger.info("LOADTEST: OVERALL-TIME-ELAPSED: {} : (in minutes)".format(processTime/60) )

