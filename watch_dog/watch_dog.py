# -*- coding: utf-8 -*-
import time
import logging
from lib_help import watch_process_is_run
from lib_help import watch_excel_is_run
from lib_help import watch_excel_is_open
from lib_help import send_warning
import random
import gc

logging.basicConfig(level=logging.DEBUG,
            format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',
            datefmt='%a, %d %b %Y %H:%M:%S',
            filename='./log/'+time.strftime('%Y-%m-%d',time.localtime(time.time()))+'_watch'+str(random.uniform(1, 10))+'.log',
             filemode='w')
def main():
    global is_err_1,is_err_2
    #监控excel是否打开
    while(False ==watch_excel_is_open()):
        logging.info('error:%s\n'%("excel is not open"))
        print "excel is not open"
        if(is_err_1 == True):
            is_err_1 = False
            send_warning(0)
        time.sleep(2)
        
    is_err_1 = True
    #监控excel是否掉线
    while(False == watch_excel_is_run()):
        logging.info('error:%s\n'%("excel is offline"))
        print "excel is offline"
        if(is_err_2 == True):
            is_err_2 = False
            send_warning(1)
        time.sleep(2)
    
    is_err_2 = True
    #监控目标进程是否运行，没有就自动启动
    watch_process_is_run(logging)

if __name__ == "__main__":
    # import win32com
    # xl = win32com.client.Dispatch("Excel.Application")
    # print xl.Range('A4').value
    # pass
    is_err_1 = True
    is_err_2 = True
    while True:
        main()
        gc.collect()
        time.sleep(3)