# -*- coding: utf-8 -*-
#!-*- encoding: utf-8 -*- 
import logging 
import wmi 
import os 
import time 
from ConfigParser import ConfigParser 
 
CONFIGFILE='./config.ini' 
config = ConfigParser() 
config.read(CONFIGFILE) 
 
ProgramPath = config.get('MonitorProgramPath','ProgramPath') 
ProcessName = config.get('MonitorProcessName','ProcessName') 
 
 
c = wmi.WMI() 
 
def main(): 
  
   ProList = []             #如果在main()函数之外ProList 不会清空列表内容. 
    for process in c.Win32_Process(): 
        ProList.append(str(process.Name)) 
 
    if ProcessName in ProList: 
        print "Service " + ProcessName + " is running...!!!" 
        if os.path.isdir("c:\MonitorWin32Process"): 
            pass 
        else: 
            os.makedirs("c:\MonitorWin32Process") 
 
    else: 
        print "Service " + ProcessName + " error ...!!!" 
        os.startfile(ProgramPath) 
 
if __name__ == "__main__": 
    while True: 
        main() 
        time.sleep(300)
