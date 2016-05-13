# -*- coding: utf-8 -*-
import time
import win32com.client
import logging
import random
import datetime
import thread
import threading
import os, sys



class timer(threading.Thread): #The timer class is derived from the class threading.Thread
    def __init__(self, num, interval):
        threading.Thread.__init__(self)
        self.thread_num = num
        self.interval = interval
        self.thread_stop = False

    def run(self): #Overwrite run() method, put what you want the thread do here
        xl = win32com.client.DispatchEx("Excel.Application")
        for i in range(4,710):
            tag = 'A'+str(i)
            xl.Range(tag).value
        #xl.Quit()
        # while not self.thread_stop:
        #     print 'Thread Object(%d), Time:%s\n' %(self.thread_num, time.ctime())
        #     time.sleep(self.interval)

    def stop(self):
        self.thread_stop = True

def main():
    xl = win32com.client.Dispatch("Excel.Application")
    #print type(xl)
    #print xl
    for i in range(4,710):
        tag = 'A'+str(i)
        xl.Range(tag).value

def main_mul_thread():
    thread1 = timer(1, 1)
    #thread2 = timer(2, 2)
    thread1.start()
    #thread2.start()
    time.sleep(10)
    thread1.stop()
    #thread2.stop()
    return

def datediff_ex(endDate):
   # begin = time.mktime(time.strptime(beginDate,"%Y-%m-%d"))
    end = time.mktime(time.strptime(endDate,"%Y-%m-%d"))
    return int((end)/86400)

def get_err_message(err_number):
    import win32api
    e_msg = win32api.FormatMessage(err_number)
    #print e_msg.decode('CP1251')
    print e_msg.decode('gbk')

def getpwd():
    ddir = sys.path[0]
    if os.path.isfile(ddir):
        ddir,filen = os.path.split(ddir)
    os.chdir(ddir)

if __name__ == "__main__":
     getpwd()
     logging.basicConfig(level=logging.DEBUG,
                format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',
                datefmt='%a, %d %b %Y %H:%M:%S',
                filename='./log/'+time.strftime('%Y-%m-%d',time.localtime(time.time()))+'_watch'+str(random.uniform(1, 10))+'.log',
                 filemode='w')

    #xl = win32com.client.Dispatch("Excel.Application")

    # while True:
    #     begin = time.clock()
    #     main(xl)
    #     end = time.clock()
    #     print "time:%.03f\n"%(end-begin)

    # begin = time.clock()
    # #main(xl)
    # main_mul_thread()
    # end = time.clock()
    # print "time:%.03f\n"%(end-begin)
    #xl.Quit()
    #time.sleep(0.01)
    #time.sleep(10)

     # beginDate = "1900-01-01"
     # endDate = time.strftime('%Y-%m-%d',time.localtime(time.time()))
     # cur_days =  datediff_ex(endDate)
     # print cur_days
     # print time.mktime((1900, 1, 1, 0, 0, 0, 0, 1, -1))

     #get_err_message(-2147352567)
     while(True):
         logging.info('info:%s\n'%("hello,world"))
         print 'hello,world\n'
        # print os.getcwd()
         time.sleep(3)
