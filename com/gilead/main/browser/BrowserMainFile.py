'''
Created on Oct 2, 2018

@author: bkotamahanti
'''

class BrowserParentClass():
    
    def __init__(self):
        self.driver = ""    
    
    def navigateToURL(self, url):
        self.driver.get(url) 
    
    def getPageTitle(self):
        return self.driver.title
    
    def getElementById(self, id):
        return self.driver.find_element_by_id(id)
       
    def isElementPresentByName(self, name):
        elements = self.driver.find_elements_by_name(name)
        if len(elements) > 0:
            return True
        else:
            return False
    
    def isElementPresentByClassName(self, classname):
        elements = self.driver.find_elements_by_class_name(classname)
        if len(elements) > 0:
            return True
        else:
            return False
    
    def closeBrowser(self):
        self.driver.quit()