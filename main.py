# -*- coding: utf-8 -*-
import time
import datetime
import lib_help
import lib_excel
import sys
import logging
import ctypes  
from multiprocessing import Process, Queue 
import math
import random
import win32com.client
import os
import gc

#eursl
reload(sys) 
sys.setdefaultencoding("utf-8")
# whnd = ctypes.windll.kernel32.GetConsoleWindow()
# if whnd != 0:  
#     ctypes.windll.user32.ShowWindow(whnd, 0)  
#     ctypes.windll.kernel32.CloseHandle(whnd)
url = ''
is_debug = True
def my_do(row):
    csv_list =[]
    today = datetime.date.today()
    
    #读取经济指标
    try:
        eci = row[0]
    except Exception, e:
        return
    #eci = eci[0:-4]
    if('*The record could not be found'==row[1] or \
      'Access Denied: User req to PE(122)' ==row[1] or \
      'Access Denied: User req to PE(5022)' == row[1] or \
      '' == row[1]) or \
      '' == row[4]:
        return

    #获取当前日期
    if('' == row[5]):
        return
    try:
       tmp_time = lib_help.format_excel_date(float(row[5]))
    except Exception, e:
       return
    if(tmp_time[0]!= today.year \
    or tmp_time[1]!= today.month \
    or tmp_time[2]!= today.day):
        return
     
    #查询更新时间是否改变
    if('#N/A'==row[6] or '*The record could not be found'==row[6]):
        if('' == row[5]):
           return
        else:
           r = float(row[5])
    else:
       try:
        r = float(row[5]) + float(row[6]) 
       except Exception, e:
         return
    tmp_date = lib_help.format_excel_date(r)
    release_date = str(tmp_date[0])+'-'+str(tmp_date[1])+'-'+str(tmp_date[2]) \
                  +' '+str(tmp_date[3])+':'+str(tmp_date[4])+':'+str(tmp_date[5])
                     
    if(lib_help.get_eci_date(eci) == release_date):
        return
    else:
        lib_help.set_eci_date(eci, release_date)
           
    #查询对应国家及其标题
    title   = row[9]
    country = row[8]
    #title   = country+title

    #获取星级
    rank    = int(round(float(row[10])))
#       ($title, $before, $prediction, $result, $country, $rank=1)
    before     = row[2] #前值
    result     = row[4]#公布值
    prediction = row[3]#预期值

    #周期获取及其格式化标题
    cycle_flag = row[1]
    cycle_fmt = lib_help.get_cycle(cycle_flag)  
    title = country+cycle_fmt+title
      
    #计量单位获取及其计算对应值
    unit_flag = row[7] 
    unit_list = unit_flag.split(' ')
    unit_flag = unit_list[0]
    unit_fmt  = lib_help.get_unit(unit_flag)
    opt_var = lib_help.get_eci_opt(eci)
    if '' == unit_fmt:
        if('%' == unit_flag): 
            if(''!= before):
                before = before+unit_flag
            if(''!=result):
                result = result+unit_flag
            if(''!=prediction):
                prediction = prediction+unit_flag
        else:
            before = before
            result = result
            prediction = prediction
    else:
        unit_fmt = float(unit_fmt)
        opt_var = float(opt_var)
        if('' != before):
            #before = round(float(before)*unit_fmt,2)
            before = round(float(before)*opt_var, 2)
        if('' != result):
            #result = round(float(result)*unit_fmt,2)
            result = round(float(result)*opt_var,2)
        if('' != prediction):
            #prediction = round(float(prediction)*unit_fmt,2)
            prediction = round(float(prediction)*opt_var,2)
     
    data = {'op':'addfinancedata',
          'content':title,
          'before':before,
          'prediction':prediction,
          'result':result,
          'country':country,
          'rank':rank,                       
        }
       
       
    logging.info('%s %s %s %s %s %s %s %s'%(eci, title, before ,prediction, result, country, rank, release_date))
     #提交测试环境
    if(is_debug == False):
        try:
            print lib_help.eci_data_post(url, data)
        except Exception,e:
            print e.message
    print 'send '+eci+"\n"
    if(is_debug == False):
      time.sleep(0.01)

    #加入数据到csv
    if(is_debug == True):
      csv_list.append([title,before,prediction,result,country,rank])
     
    if(is_debug == True):
      lib_help.write_csv("my.csv", csv_list)

#预处理,获取当天需要处理的行数
def pretreatment(xl, begin_row, end_row, target_col, result_col):
  #计算从1900-1-1到当前的天数
  beginDate = "2016-2-25"
  endDate = time.strftime('%Y-%m-%d',time.localtime(time.time()))
  cur_days =  lib_help.datediff_ex(beginDate,endDate)
  cur_days += 42425

  #查询所有和要求天数相同的所有经济指标的行号
  line_no_list = {}
  for i in range(begin_row, end_row+1):
    try:
        days = str(xl.Cells(i, target_col).Value).strip()
        #days = str(xl.Range('F'+str(i)).value).strip()
    except Exception,e:
        print e
        continue
    if('' == days or \
        '#N/A *The record could not be found'==days):
        continue
    try:
        days = float(days)
    except Exception,e:
        # print e
        continue
    days = int(days)

    if(cur_days == days):
        try:
            result = str(xl.Cells(i, result_col).value).strip()
        except Exception,e:
            continue
        if('' == result):
            continue
        #判定是否发布过
        if(check_is_public(xl, i)):
            continue
        line_no_list[i] = lib_excel.excel_table_row_byindex_dynamic(xl, i)
  return line_no_list

'''预处理,获取当天需要处理的行数
 (一天只跑一次，当天需要公布的数据；
 然后不停循环查找需要当天公布的公布值)
'''
def pretreatment_only_day(xl, begin_row, end_row, target_col, result_col):
    global line_no_list
    #计算从1900-1-1到当前的天数
    beginDate = "2016-2-25"
    endDate = time.strftime('%Y-%m-%d',time.localtime(time.time()))
    cur_days =  lib_help.datediff_ex(beginDate,endDate)
    cur_days += 42425
    
    
    #判定是否是当天(yes)
    if(cur_days == int(lib_help.get_curdays())):
        #只跑当天需要的公布数据
        for k in line_no_list:
            #判定是否发布过
            if(check_is_public(xl, k)):
                continue
            line_no_list[k][4] = lib_excel.excel_table_row_public_value(xl, k)
                             #lib_excel.excel_table_row_byindex_dynamic(xl, k)
    else:
        lib_help.set_curdays(cur_days)
        #获取当前需要跑的eci及其对应的行号        
        for i in range(begin_row, end_row+1):
            try:
                days = str(xl.Cells(i, target_col).Value).strip()
            except Exception,e:
                print e
                continue
            if('' == days or \
               '#N/A *The record could not be found'==days):
                continue
            try:
                days = float(days)
            except Exception,e:
                continue
            days = int(days)

            if(cur_days == days):
                try:
                    result = str(xl.Cells(i, result_col).value).strip()
                except Exception,e:
                    continue
                if('' == result):
                    continue
                #判定是否发布过
                if(check_is_public(xl, i)):
                    continue
                line_no_list[i] = lib_excel.excel_table_row_byindex_dynamic(xl, i)
    return line_no_list

def pretreatment_v1(xl, begin_row, end_row, target_col, result_col):
  #计算从1900-1-1到当前的天数
  beginDate = "2016-2-25"
  endDate = time.strftime('%Y-%m-%d',time.localtime(time.time()))
  cur_days =  lib_help.datediff_ex(beginDate,endDate)
  del beginDate
  cur_days += 42425

  #查询所有和要求天数相同的所有经济指标的行号
  line_no_list = {}
  tmp_tag = "F%d:F%d"%(begin_row, end_row)

  try:
    old_list = xl.Range(tmp_tag).value
  except Exception,e:
      print e.message
      return {}
  del tmp_tag
  i = begin_row
  del begin_row
  for day in old_list:
    days = str(day[0]).strip()
    if('' == days or \
        '#N/A *The record could not be found'==days):
        del days
        continue
    try:
        days = float(days)
    except Exception,e:
        # print e
        del days
        continue
    days = int(days)

    if(cur_days == days):
        try:
            result = str(xl.Cells(i, result_col).value).strip()
        except Exception,e:
            del days
            continue
        if('' == result):
            del days
            continue
        #判定是否发布过
        if(check_is_public(xl, i)):
            del result
            del days
            continue
        line_no_list[i] = lib_excel.excel_table_row_byindex_dynamic(xl, i)
    i = i+1
    del days
  del old_list
  return line_no_list

#检查此条指标当天是否发布过
def check_is_public(xl, target_row):
    try:
        eci = str(xl.Range('A'+str(target_row)).value).strip()
    except Exception,e:
        return False
    eci = eci[0:-4]
    ini_eci_date_str = lib_help.get_eci_date(eci)
    del eci
    if('' == ini_eci_date_str):
        del ini_eci_date_str
        return False
    ini_eci_date_l = ini_eci_date_str.split(' ')
    ini_eci_date = ini_eci_date_l[0]
    del ini_eci_date_l
    cur_date = time.strftime('%Y-%m-%d',time.localtime(time.time()))
    # print 'cur_date:%s\n'%(cur_date)
    # print 'ini_eci_date:%s\n'%(ini_eci_date)
    if(0 == lib_help.datediff_ex(cur_date, ini_eci_date)):
        del cur_date
        del ini_eci_date
        return True
    del cur_date
    del ini_eci_date
    return False


#初始化
def my_init():
    global url,is_debug,xl,end_row
    global begin_row,target_col,result_col
    global line_no_list
    line_no_list = {}
    begin_row = 4
    target_col = 6
    result_col = 5 
    getpwd()
    end_row = lib_excel.get_table_rows()
    xl = win32com.client.Dispatch("Excel.Application")
    is_debug = bool(lib_help.get_is_debug())
    print 'is_debug:%s\n'%(str(is_debug))
    #初始化当天，如果程序奔溃，再跑一次检测
    lib_help.set_curdays(0)
    #获取提交数据url
    if(is_debug == False):
        url = lib_help.get_post_data_url()
    logging.basicConfig(level=logging.DEBUG,
                format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',
                datefmt='%a, %d %b %Y %H:%M:%S',
                filename='./log/'+time.strftime('%Y-%m-%d',time.localtime(time.time()))+'_myapp'+str(random.uniform(1, 10))+'.log',
                 filemode='w')

    tables = lib_excel.excel_table_byindex_init_basicdata()
    country_l = {}
    rank_l = {}
    title_l = {}
    unit_l = {}
    opt_l = {}
    for row in tables:
        eci = row[0]
        eci = eci[0:-4]
        unit_l[eci] = row[7]
        country_l[eci] = row[8] #国家
        title_l[eci] = row[9]   #标题
        rank_l[eci]  = row[10]  #星级
        opt_l[eci]= row[12]     #换算量
    lib_help.set_eci_basic(country_l, rank_l, title_l, unit_l, opt_l)

def main_mul():
    #获取提交数据url
    url = lib_help.get_post_data_url()
    #tables = lib_excel.excel_table_byindex()
    #tables = lib_excel.excel_table_byindex_dynamic()
    threads = []
    rows = lib_excel.get_table_rows()
    step =200
    start_row = 4
    end_row = start_row+step
    q = Queue()  
    p_count = int(math.ceil((rows*1.0)/step))
    # p = Process(target=lib_excel.offer, args=(q, start_row, end_row))  
    # p.start() 
    for i in range(p_count):
        p = Process(target=lib_excel.offer, args=(q, start_row, end_row))  
        p.start() 
        #print "start:%d,end:%d\n"%(start_row, end_row)
        start_row = end_row +1 
        end_row = start_row + step
        if(end_row>=rows):
          end_row = rows

    
    while True:
        my_data = q.get()
        #my_do(my_data)

def main_simple():
  url = lib_help.get_post_data_url()
  pass

def main():
  global xl,end_row
  global begin_row,target_col,result_col

  #预处理
  row_list = pretreatment_only_day(xl, begin_row, end_row, target_col, result_col)
  #row_list = pretreatment_v1(xl, begin_row, end_row, target_col, result_col)
  for k in row_list:
      my_do(row_list[k])

  del row_list
  #预跑十遍
  # run_times = 10
  # for i in range(run_times):
  #     for k in row_list:
  #        data = lib_excel.excel_table_row_byindex_dynamic(xl, k)
  #        my_do(data)

#version v1:
#改善程序模式
#1.在指定时间获取当天公布的经济指标;
#2.提前指定时间开始获取公布时间;
#3.在指定时间获取预期和前值;
#4.在准点没获取到公布值后,延迟指定时间段获取公布值,还是获取不到就放弃;
def main_v1():
    
    pass

#改变当前执行路径
def getpwd():
    ddir = sys.path[0]
    if os.path.isfile(ddir):
        ddir,filen = os.path.split(ddir)
    os.chdir(ddir)

if __name__ == '__main__':
    # begin = time.clock()
    # main()
    # #main_simple()
    # end = time.clock()
    # print "time:%.03f\n"%(end-begin)

    # begin = time.clock()
    # #beginDate = "2016-2-25"
    # #beginDate = "1900-1-1"
    # basic_number = 42425
    # endDate = time.strftime('%Y-%m-%d',time.localtime(time.time()))
    # cur_days =  lib_help.datediff(beginDate,endDate)
    # print "cur_days:%d\n"%(cur_days)
    # end = time.clock()
    # print "time:%.03f\n"%(end-begin)

    # xl = win32com.client.Dispatch("Excel.Application")
    # print xl.Range("F9").value
     my_init()


     # begin = time.clock()
     # main()
     # end = time.clock()
     # print "time:%.03f\n"%(end-begin)
     # print "end"

            
     while True:
         begin = time.clock()
         main()
         end = time.clock()
         print "time:%.03f\n"%(end-begin)
         del begin
         del end
         gc.collect()
         print "end"
        
         #logging.info('run time:%.03f\n'%(end-begin))