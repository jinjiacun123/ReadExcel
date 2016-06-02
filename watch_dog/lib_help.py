# -*- coding: utf-8 -*-
import logging
import wmi
import os
import time
import win32com
import gc

from ConfigParser import ConfigParser

def watch_excel_is_open():
    xl = win32com.client.Dispatch("Excel.Application")
    try:
        xl.Sheets("Sheet1")
    except Exception,e:
        del xl
        gc.collect()
        return False
    del xl
    gc.collect()
    return True

def watch_excel_is_run():
    xl = win32com.client.Dispatch("Excel.Application")
    is_get = False
    is_change = False
    #获取旧值
    while(False == is_get):
        try:
            old_value = xl.sheets("Sheet2").Range('K3').value
            print old_value
        except Exception,e:
            is_get = False
            continue
        is_get = True

    #延长2秒
    time.sleep(1)

    index = 0
    while(index<=10):
        #获取新值
        is_get = False
        while(False == is_get):
            try:
               new_value =xl.sheets("Sheet2").Range('K3').value
               print new_value
            except Exception,e:
                is_get = False
                continue
            is_get = True
            if(new_value != old_value):
                del is_get
                del is_change
                del new_value
                del old_value
                del xl
                gc.collect()
                return True
        time.sleep(1)
        index = index +1

    del is_get
    del index
    del xl
    gc.collect()

    return is_change

def watch_process_is_run(logging):
    CONFIGFILE='./config.ini'
    config = ConfigParser()
    config.read(CONFIGFILE)
    ProgramPath = config.get('MonitorProgramPath','ProgramPath')
    ProcessName = config.get('MonitorProcessName','ProcessName')

    c = wmi.WMI()
    ProList = []             #如果在main()函数之外ProList 不会清空列表内容.
    for process in c.Win32_Process():
        ProList.append(str(process.Name))

    if ProcessName in ProList:
        print "Service " + ProcessName + " is running...!!!"
        logging.info('info:%s\n'%("Service " + ProcessName + " is running...!!!"))
    else:
        print "Service " + ProcessName + " error ...!!!"
        logging.info('error:%s\n'%("Service " + ProcessName + " error ...!!!"))
        print os.startfile(ProgramPath)
        #send_warning(2)
        time.sleep(5)
    del CONFIGFILE
    del config
    del ProgramPath
    del ProcessName
    del c
    del ProList
    gc.collect()

#发送警告
def send_warning(type):
    type_list = ("Excel没打开","Excel掉线","软件没打开或者奔溃")
    message = "%s "%(type_list[type])
    call_wx(u'快讯-经济指标监控', unicode(message, 'utf-8'))

#调用微信报警
def call_wx(e_type, e_description):
    from suds.client import Client
    url = 'http://112.84.186.217:8022/WeChatMonitoringService.WeChatService.svc?WSDL'
    r_password = 'zjzx2015'
    r_time = time.strftime('%Y-%m-%d %H:%M',time.localtime(time.time()))
    client=Client(url)
    result = client.service.KxOperation(r_password, e_type, e_description, r_time)
    print result

def main():
    #watch_process_is_run()
    print watch_excel_is_open()

if __name__ == "__main__":
    title = '我的测试'
    content = '测试内容'
   # title = title.encode('utf-8')
   # content = content.encode('utf-8')
    call_wx(unicode(title, 'utf-8'), unicode(content, 'utf-8'))
    pass
    '''
    while(True):
        main()
        time.sleep(3)
    '''
