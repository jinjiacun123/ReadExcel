# -*- coding: utf-8 -*-
from xlwings import Range
import time

def main_xlwings():
    #Range('A1').value = 1000
    try:
        print Range('Sheet1','A1:A3').value
    except Exception,e:
        print e.message

def main_com():
    import win32com
    xl = win32com.client.Dispatch("Excel.Application")
    try:
        print xl.Sheets('Sheet1').Range('A4:K4').value
    except Exception,e:
        print e.message
    del xl

if __name__ == "__main__":
    while True:
        begin = time.clock()
        main_com()
        end = time.clock()
        print "com time:%.03f\n"%(end-begin)

        begin = time.clock()
        main_xlwings()
        end = time.clock()
        print "xlwings time:%.03f\n"%(end-begin)

        time.sleep(3)