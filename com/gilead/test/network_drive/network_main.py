'''
Created on Aug 21, 2019

@author: bkotamahanti
'''
from definitions import ROOT_DIR, LOG_DIR
import logging
import os
from com.gilead.test.network_drive.NetworkFileShareTestFile import NetworkFileShareTestClass

class NetworkMainClass():
    
    # Initializing 
    def __init__(self):
        
        self.networkDetailLogFile=ROOT_DIR + LOG_DIR +"NetWorkDetailTestLog.txt"
        ''' Checking and removing the log file before opening and writing to it...'''
        if (os.path.isfile(self.networkDetailLogFile)):
            os.remove(self.networkDetailLogFile)
        
        self.logger=logging.getLogger("root")
        for handler in self.logger.handlers[:]:
            self.logger.removeHandler(handler)
            self.logger.info(str(handler)+"========================")
        
        logging.basicConfig(filename = self.networkDetailLogFile, 
                                    level=logging.DEBUG, 
                                    format='%(asctime)s : %(levelname)s :  %(filename)s :  %(funcName)s()  %(message)s'
    #                                 format='[%(asctime)s : %(levelname)s : filename ( %(filename)s ): line number ( %(lineno)s )-------function name ( %(funcName)s()) ] %(message)s' 
                                    )
        
             
              
        self.network= NetworkFileShareTestClass(self.logger)

    # Calling destructor
    def __del__(self):
        logging.shutdown()
    