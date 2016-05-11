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

#eursl
reload(sys) 
sys.setdefaultencoding("utf-8") 
logging.basicConfig(level=logging.DEBUG,
                format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',
                datefmt='%a, %d %b %Y %H:%M:%S',
                filename='./log/'+time.strftime('%Y-%m-%d',time.localtime(time.time()))+'_myapp'+str(random.uniform(1, 10))+'.log',
                 filemode='w')
#获取提交数据url
url = lib_help.get_post_data_url()
is_debug = False
# whnd = ctypes.windll.kernel32.GetConsoleWindow()  
# if whnd != 0:  
#     ctypes.windll.user32.ShowWindow(whnd, 0)  
#     ctypes.windll.kernel32.CloseHandle(whnd)

def my_do(row):
    csv_list =[]
    today = datetime.date.today()
    
    #读取经济指标
    try:
        eci = row[0]
    except Exception, e:
        return
    eci = eci[0:-4]
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
        if('' != before):
            before = round(float(before)*unit_fmt,2)
        if('' != result):
            result = round(float(result)*unit_fmt,2)
        if('' != prediction):
            prediction = round(float(prediction)*unit_fmt,2)
     
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
       print lib_help.eci_data_post(url, data)
    print 'send '+eci+"\n"
    if(is_debug == False):
      time.sleep(1)

    #加入数据到csv
    if(is_debug == True):
      csv_list.append([title,before,prediction,result,country,rank])
     
    if(is_debug == True):
      lib_help.write_csv("my.csv", csv_list)

#预处理,获取当天需要处理的行数
def pretreatment(xl, begin_row, end_row, target_col, result_col):
  #计算从1900-1-1到当前的天数
  beginDate = "1900-1-1"
  endDate = time.strftime('%Y-%m-%d',time.localtime(time.time()))
  cur_days =  lib_help.datediff_ex(beginDate,endDate)
  cur_days += 2

  #查询所有和要求天数相同的所有经济指标的行号
  line_no_list = {}

  for i in range(begin_row, end_row+1):
    try:
        days = str(xl.Cells(i, target_col).value).strip()
    except Exception,e:
        # print e
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
        line_no_list[i] = lib_excel.excel_table_row_byindex_dynamic(xl, i)
  return line_no_list
  
#初始化
def my_init():
    tables = lib_excel.excel_table_byindex_init_basicdata()
    country_l = {}
    rank_l = {}
    title_l = {}
    unit_l = {}
    for row in tables:
        eci = row[0]
        eci = eci[0:4]
        unit_l[eci] = row[7]
        country_l[eci] = row[8]
        title_l[eci] = row[9]
        rank_l[eci]  = row[10]
    lib_help.set_eci_basic(country_l, rank_l, title_l, unit_l)

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
  begin_row = 4
  end_row = lib_excel.get_table_rows()
  target_col = 6
  result_col = 5
  xl = win32com.client.Dispatch("Excel.Application")
  #预处理
  row_list = pretreatment(xl, begin_row, end_row, target_col, result_col)
  for k in row_list:
      my_do(row_list[k])
  #预跑十遍
  # run_times = 10
  # for i in range(run_times):
  #     for k in row_list:
  #        data = lib_excel.excel_table_row_byindex_dynamic(xl, k)
  #        my_do(data)


if __name__ == '__main__':
    # begin = time.clock()
    # main()
    # #main_simple()
    # end = time.clock()
    # print "time:%.03f\n"%(end-begin)
     my_init()
     while True:
         begin = time.clock()
         main()
         end = time.clock()
         print "time:%.03f\n"%(end-begin)
         print "end"
         logging.info('run time:%.03f\n'%(end-begin))