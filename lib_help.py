# -*- coding: utf-8 -*-
import urllib2, urllib
import time
import datetime
import ConfigParser
from config import Db_Connector

#读取经济指标对应国家
def get_eci_country(eci):
     f = Db_Connector("eci.ini")
     return f.read('baseconf', eci)

#读取经济指标对应标题
def get_eci_title(eci):
    f = Db_Connector("eci.ini")
    return f.read('titleconf', eci)
    
#读取经济指标对应星级
def get_eci_rank(eci):
    f = Db_Connector("eci.ini")
    return f.read('rankconf', eci)

#读取经济指标对应计量单位
def get_eci_unit(eci):
    f = Db_Connector("eci.ini")
    return f.read('unitconf', eci)

#写入经济指标发布时间
def set_eci_date(eci,date):
    f = Db_Connector("eci.ini")
    f.write('baseconf', eci, date)

#写入经济指标基本值
def set_eci_basic(country_l, rank_l, title_l, unit_l):
    f = Db_Connector("eci.ini")
    f.mul_write('countryconf',country_l)
    f.mul_write('rankconf',rank_l)
    f.mul_write('titleconf', title_l)
    f.mul_write('unitconf', unit_l)

#查询经济指标发布时间
def get_eci_date(eci):
    f = Db_Connector("eci.ini")
    return f.read('baseconf', eci)
    
#获取周期
def get_cycle(cycle_flag):
    f = Db_Connector("my.ini")
    return f.read('cycleconf', cycle_flag)

#格式化excel日期
def format_excel_date(date):
    import xlrd
    r=xlrd.xldate_as_tuple(date, 0)
    return r
    
#获取计量单位
def get_unit(unit_flag):
    f = Db_Connector("my.ini")
    return f.read('unitconf', unit_flag)

#两个日期相隔多少天，例：2008-10-03和2008-10-01是相隔两天  
# def datediff(beginDate,endDate):
#     format="%Y-%m-%d";
#     bd=strtodatetime(beginDate,format)
#     ed=strtodatetime(endDate,format)
#     oneday=datetime.timedelta(days=1)
#     count=0
#     while bd!=ed:
#         ed=ed-oneday
#         count+=1
#     return count
# def strtodatetime(datestr,format):
#     return datetime.datetime.strptime(datestr,format)

def datediff_ex(beginDate,endDate):
    begin = time.mktime(time.strptime(beginDate,"%Y-%m-%d"))
    end = time.mktime(time.strptime(endDate,"%Y-%m-%d"))
    return int((end-begin)/86400)

#获取提交数据url
def get_post_data_url():
    f = Db_Connector("my.ini")
    return f.read('baseconf','data_url')

#写csv文件
def write_csv(csv_file='my.csv',data=None):
    import csv
    with open(csv_file, 'wb') as csvfile:
        spamwriter = csv.writer(csvfile, delimiter=' ',
                                quotechar='|', quoting=csv.QUOTE_MINIMAL)
        '''                                
        spamwriter.writerow(['Spam'] * 5 + ['Baked Beans'])
        spamwriter.writerow(['Spam', 'Lovely Spam', 'Wonderful Spam'])    
        '''
        for row_data in data:
            spamwriter.writerow(row_data)    
        
    
#提交测试环境
def eci_data_post(url, data):
    headers ={ 'User-Agent' :"CngoldClient/1.0",
                'Source':"cngold.com.cn",
    }
    data    = urllib.urlencode(data)
    req = urllib2.Request(url,data, headers)
    f = urllib2.urlopen(req)
    try:
        result =  f.read() 
    except Exception,e:
        print e
    return result

def eci_mul_data_post(url, data):
    headers ={ 'User-Agent' :"CngoldClient/1.0",
                'Source':"cngold.com.cn",
    }
    data    = urllib.urlencode(data)
    req = urllib2.Request(url,data, headers)
    f = urllib2.urlopen(req)
    try:
        result =  f.read()
    except Exception,e:
        print e
    return result

##日期比较
#def compare_time(l_time,end_t):#l_time-当前日期,end_t-发布日期
#    import time
#    #s_time = time.mktime(time.strptime(start_t,'%Y%m%d%H%M')) # get the seconds for specify date    
#    e_time = time.mktime(time.strptime(end_t,'%Y%m%d%H%M'))    
#    log_time = time.mktime(time.strptime(l_time,'%Y-%m-%d %H:%M:%S'))    
#    if (float(log_time) >= float(e_time):    
#        return True    
#    return False

if __name__ == "__main__":
    # url = "http://192.168.1.248:86/ReutersHandler.ashx"
    # data = {'op':'addfinancedata',
    #       'content':"中国4月消费者物价指数(月率)",
    #       'before':"-0.4%",
    #       'prediction':"-0.2%",
    #       'result':"-0.2%" ,
    #       'country':"中国",
    #       'rank':3,                       
    #     }
    # print eci_data_post(url, data)

    # beginDate = "1900-1-1"
    # endDate = time.strftime('%Y-%m-%d',time.localtime(time.time()))
    # print endDate
    # print datediff(beginDate,endDate)
    # import win32com.client
    # xl = win32com.client.Dispatch("Excel.Application")
    # data = xl.Range('H9').value
    # uint_list = data.split(' ')
    # type(uint_list[0])
    pass
