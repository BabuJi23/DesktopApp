''''
Created on Mar 19, 2020
@author: jnguyen19
'''

import os
import glob
from jinja2 import Environment, FileSystemLoader
from definitions import cpu_mem_stats, app_list
from com.gilead.main.utils.DesktopAppUtilsFile import DesktopAppCommonFunctionsClass

#To parse already pre-processed LOADTEST parsed into HTML table and html output:

# Create the jinja2 environment.
current_directory = os.path.dirname(os.path.abspath(__file__))
env = Environment(loader=FileSystemLoader(current_directory))
timestamp = DesktopAppCommonFunctionsClass.getDateTimeString()
inputFile = 'LOADTEST.txt'
outputFile = 'Load-Testing_Summary_Report-'+timestamp+'.html'


# Find all files with the j2 extension in the current directory
templates = glob.glob('*.j2')

#class Generate_ReportClass():
def render_template(filename, cpu_mem_stats):
    (OSName, OSVersion) = DesktopAppCommonFunctionsClass.getWindowsOSBuildVersion()

    return env.get_template(filename).render(
        winBuildVer = str(OSName)+str(OSVersion),
        pyVer=DesktopAppCommonFunctionsClass.getPythonVersion(),
        date=DesktopAppCommonFunctionsClass.getDate(),
        app_one='Notepad',
        app_two='Acrobat Reader',
        app_three='Outlook',
        app_four='Excel',
        app_five='Word',

        mem_start=convert2FloatAndRounded(cpu_mem_stats['mem_start']),
        cpu_start=cpu_mem_stats['start_cpu'],
        time_start=convert2FloatAndRounded(cpu_mem_stats['start_time']),

        #APP 1, loops 1-5:
        mem_app1_loop1=convert2FloatAndRounded(cpu_mem_stats['mem_after_app1_0']),
        mem_app1_loop2=convert2FloatAndRounded(cpu_mem_stats['mem_after_app1_1']),
        mem_app1_loop3=convert2FloatAndRounded(cpu_mem_stats['mem_after_app1_2']),
        mem_app1_loop4=convert2FloatAndRounded(cpu_mem_stats['mem_after_app1_3']),
        mem_app1_loop5=convert2FloatAndRounded(cpu_mem_stats['mem_after_app1_4']),

        cpu_app1_loop1=cpu_mem_stats['CPUUsageApp1-Loop0'],
        cpu_app1_loop2=cpu_mem_stats['CPUUsageApp1-Loop1'],
        cpu_app1_loop3=cpu_mem_stats['CPUUsageApp1-Loop2'],
        cpu_app1_loop4=cpu_mem_stats['CPUUsageApp1-Loop3'],
        cpu_app1_loop5=cpu_mem_stats['CPUUsageApp1-Loop4'],

        et_app1_loop1=convert2FloatAndRounded(cpu_mem_stats['ElapsedTimeApp1-Loop0']),
        et_app1_loop2=convert2FloatAndRounded(cpu_mem_stats['ElapsedTimeApp1-Loop1']),
        et_app1_loop3=convert2FloatAndRounded(cpu_mem_stats['ElapsedTimeApp1-Loop2']),
        et_app1_loop4=convert2FloatAndRounded(cpu_mem_stats['ElapsedTimeApp1-Loop3']),
        et_app1_loop5=convert2FloatAndRounded(cpu_mem_stats['ElapsedTimeApp1-Loop4']),

        # APP 2, loops 1-5:
        mem_app2_loop1=convert2FloatAndRounded(cpu_mem_stats['mem_after_app2_0']),
        mem_app2_loop2=convert2FloatAndRounded(cpu_mem_stats['mem_after_app2_1']),
        mem_app2_loop3=convert2FloatAndRounded(cpu_mem_stats['mem_after_app2_2']),
        mem_app2_loop4=convert2FloatAndRounded(cpu_mem_stats['mem_after_app2_3']),
        mem_app2_loop5=convert2FloatAndRounded(cpu_mem_stats['mem_after_app2_4']),

        cpu_app2_loop1=cpu_mem_stats['CPUUsageApp2-Loop0'],
        cpu_app2_loop2=cpu_mem_stats['CPUUsageApp2-Loop1'],
        cpu_app2_loop3=cpu_mem_stats['CPUUsageApp2-Loop2'],
        cpu_app2_loop4=cpu_mem_stats['CPUUsageApp2-Loop3'],
        cpu_app2_loop5=cpu_mem_stats['CPUUsageApp2-Loop4'],

        et_app2_loop1=convert2FloatAndRounded(cpu_mem_stats['ElapsedTimeApp2-Loop0']),
        et_app2_loop2=convert2FloatAndRounded(cpu_mem_stats['ElapsedTimeApp2-Loop1']),
        et_app2_loop3=convert2FloatAndRounded(cpu_mem_stats['ElapsedTimeApp2-Loop2']),
        et_app2_loop4=convert2FloatAndRounded(cpu_mem_stats['ElapsedTimeApp2-Loop3']),
        et_app2_loop5=convert2FloatAndRounded(cpu_mem_stats['ElapsedTimeApp2-Loop4']),

        # APP 3, loops 1-5:
        mem_app3_loop1=convert2FloatAndRounded(cpu_mem_stats['mem_after_app3_0']),
        mem_app3_loop2=convert2FloatAndRounded(cpu_mem_stats['mem_after_app3_1']),
        mem_app3_loop3=convert2FloatAndRounded(cpu_mem_stats['mem_after_app3_2']),
        mem_app3_loop4=convert2FloatAndRounded(cpu_mem_stats['mem_after_app3_3']),
        mem_app3_loop5=convert2FloatAndRounded(cpu_mem_stats['mem_after_app3_4']),

        cpu_app3_loop1=cpu_mem_stats['CPUUsageApp3-Loop0'],
        cpu_app3_loop2=cpu_mem_stats['CPUUsageApp3-Loop1'],
        cpu_app3_loop3=cpu_mem_stats['CPUUsageApp3-Loop2'],
        cpu_app3_loop4=cpu_mem_stats['CPUUsageApp3-Loop3'],
        cpu_app3_loop5=cpu_mem_stats['CPUUsageApp3-Loop4'],

        et_app3_loop1=convert2FloatAndRounded(cpu_mem_stats['ElapsedTimeApp3-Loop0']),
        et_app3_loop2=convert2FloatAndRounded(cpu_mem_stats['ElapsedTimeApp3-Loop1']),
        et_app3_loop3=convert2FloatAndRounded(cpu_mem_stats['ElapsedTimeApp3-Loop2']),
        et_app3_loop4=convert2FloatAndRounded(cpu_mem_stats['ElapsedTimeApp3-Loop3']),
        et_app3_loop5=convert2FloatAndRounded(cpu_mem_stats['ElapsedTimeApp3-Loop4']),

        # APP 4, loops 1-5:
        mem_app4_loop1=convert2FloatAndRounded(cpu_mem_stats['mem_after_app4_0']),
        mem_app4_loop2=convert2FloatAndRounded(cpu_mem_stats['mem_after_app4_1']),
        mem_app4_loop3=convert2FloatAndRounded(cpu_mem_stats['mem_after_app4_2']),
        mem_app4_loop4=convert2FloatAndRounded(cpu_mem_stats['mem_after_app4_3']),
        mem_app4_loop5=convert2FloatAndRounded(cpu_mem_stats['mem_after_app4_4']),

        cpu_app4_loop1=cpu_mem_stats['CPUUsageApp4-Loop0'],
        cpu_app4_loop2=cpu_mem_stats['CPUUsageApp4-Loop1'],
        cpu_app4_loop3=cpu_mem_stats['CPUUsageApp4-Loop2'],
        cpu_app4_loop4=cpu_mem_stats['CPUUsageApp4-Loop3'],
        cpu_app4_loop5=cpu_mem_stats['CPUUsageApp4-Loop4'],

        et_app4_loop1=convert2FloatAndRounded(cpu_mem_stats['ElapsedTimeApp4-Loop0']),
        et_app4_loop2=convert2FloatAndRounded(cpu_mem_stats['ElapsedTimeApp4-Loop1']),
        et_app4_loop3=convert2FloatAndRounded(cpu_mem_stats['ElapsedTimeApp4-Loop2']),
        et_app4_loop4=convert2FloatAndRounded(cpu_mem_stats['ElapsedTimeApp4-Loop3']),
        et_app4_loop5=convert2FloatAndRounded(cpu_mem_stats['ElapsedTimeApp4-Loop4']),

        # APP 5, loops 1-5:
        mem_app5_loop1=convert2FloatAndRounded(cpu_mem_stats['mem_after_app5_0']),
        mem_app5_loop2=convert2FloatAndRounded(cpu_mem_stats['mem_after_app5_1']),
        mem_app5_loop3=convert2FloatAndRounded(cpu_mem_stats['mem_after_app5_2']),
        mem_app5_loop4=convert2FloatAndRounded(cpu_mem_stats['mem_after_app5_3']),
        mem_app5_loop5=convert2FloatAndRounded(cpu_mem_stats['mem_after_app5_4']),

        cpu_app5_loop1=cpu_mem_stats['CPUUsageApp5-Loop0'],
        cpu_app5_loop2=cpu_mem_stats['CPUUsageApp5-Loop1'],
        cpu_app5_loop3=cpu_mem_stats['CPUUsageApp5-Loop2'],
        cpu_app5_loop4=cpu_mem_stats['CPUUsageApp5-Loop3'],
        cpu_app5_loop5=cpu_mem_stats['CPUUsageApp5-Loop4'],

        et_app5_loop1=convert2FloatAndRounded(cpu_mem_stats['ElapsedTimeApp5-Loop0']),
        et_app5_loop2=convert2FloatAndRounded(cpu_mem_stats['ElapsedTimeApp5-Loop1']),
        et_app5_loop3=convert2FloatAndRounded(cpu_mem_stats['ElapsedTimeApp5-Loop2']),
        et_app5_loop4=convert2FloatAndRounded(cpu_mem_stats['ElapsedTimeApp5-Loop3']),
        et_app5_loop5=convert2FloatAndRounded(cpu_mem_stats['ElapsedTimeApp5-Loop4']),

        #Summary Results:
        overall_mem_usage=convert2FloatAndRounded(cpu_mem_stats['overall_mem_usage']),
        overall_cpu_usage=cpu_mem_stats['overall_cpu_usage'],
        overall_time_lapsed=convert2FloatAndRounded(cpu_mem_stats['overall_time_elapsed']),
    )# End of Render Template


def parseLoadTestLog(cpu_mem_stats):

    currentDir = os.path.dirname(os.path.realpath(__file__))
    print(currentDir)

    with open(currentDir+"\\"+inputFile, 'r') as fIn:
        line = fIn.readline()
        while line:
            tokens = line.split(':')

            if len(tokens) == 9:
                if ' CPU_MEM_STATS For' in tokens:
                    cpu_mem_stats[removeTicks(tokens[7])] = removeTicks(tokens[8])
                if ' OVERALL-MEM-USAGE' in tokens:
                    cpu_mem_stats['overall_mem_usage'] = cleanString(tokens[7])
                if ' OVERALL-CPU-USAGE' in tokens:
                    cpu_mem_stats['overall_cpu_usage'] = cleanString(tokens[7])
                if ' OVERALL-TIME-ELAPSED' in tokens:
                    cpu_mem_stats['overall_time_elapsed'] = cleanString(tokens[7])
            if len(tokens) == 11:
                if ' Lapsed Time for' in tokens:
                    cpu_mem_stats['ElapsedTimeApp'+str(app_list[cleanString(tokens[7])])+'-Loop'+str(parseLoopCount(tokens[9]))] = getElapsedTime(tokens[10])
                pass
            if len(tokens) == 13:
                if ' Cummulative CPU Usage  for app' in tokens:
                    cpu_mem_stats['CPUUsageApp'+str(cleanString(tokens[7]))+'-Loop'+str(cleanString(tokens[9]))] = cleanString(tokens[11])
            line = fIn.readline()
    fIn.close()
    return cpu_mem_stats

def verify_dictionary_content(cpu_mem_stats):
    for key in cpu_mem_stats:
        if cpu_mem_stats[key]:
            print("PARSER_TEST:   \'{}\':\'{}\' ".format(key, cpu_mem_stats[key]))
            pass

def removeTicks(string):
    newString = string.strip().replace("'", "")
    return newString

#' string '
def cleanString(string):
    newString = string.strip()
    return newString
#' loop 3'
def parseLoopCount(string):
    list = string.strip().split(' ')
    return list[1]

#' is 39.033811499999956 seconds'
def getElapsedTime(string):
    list = string.strip().split(' ')
    return list[1]

def convert2FloatAndRounded(stringIn):
    return round(float(stringIn), 2)

def main():
    #Parse LoadTestLog file and load up defined Data Structure for load testing
    parseLoadTestLog(cpu_mem_stats)

    #For debugging:
    #verify_dictionary_content(cpu_mem_stats)

    #Render and update Template and write out results for new report
    for f in templates:
        rendered_string = render_template(f, cpu_mem_stats)
        f = open(outputFile, 'w')
        print(rendered_string, file=f)

if __name__ == "__main__":
    main()

