'''
Created on Mar 18, 2020
@author: jnguyen19
'''

import inspect
import os

inputFile = 'DetailTestLog.txt'
outputFile = 'LOADTEST.txt'
matchString = 'LOADTEST:'

#To output only LOADTEST logging related.
with open(outputFile, 'w') as fo:
	with open(inputFile, 'r') as f:
		line = f.readline()
		while line:
			if matchString in line:
				fo.write(line)
			line = f.readline()
	f.close()
fo.close()	

