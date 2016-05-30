# -*- coding: utf-8 -*-
import  xdrlib ,sys
import xlrd
import win32com.client
import datetime
import thread
import threading
import time
import math
import lib_help
#from Queue import Queue
from multiprocessing import Process, Queue 
reload(sys) 
sys.setdefaultencoding("utf-8")

def open_excel(file= 'file.xls'):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception,e:
        print str(e)

#根据索引获取Excel表格中的数据   参数:file：Excel文件路径     colnameindex：表头列名所在行的所以  ，by_index：表的索引
def excel_table_byindex(file= 'get_rt_data.xls',colnameindex=0,by_index=0):
    data = open_excel(file)
    table = data.sheets()[by_index]
    nrows = table.nrows #行数
    ncols = table.ncols #列数
    colnames =  table.row_values(colnameindex) #某一行数据
    #print colnames[0]
    #print colnames
    list =[]
    for rownum in range(3,nrows):
         row = table.row_values(rownum)
         index = 0
         if row:
             app = {}
             for i in range(0,len(colnames)):
                 app[i] = str(row[i]).strip()
             list.append(app)
             index=index+1
    return list

#获取指定行的数据
def excel_table_get_unit(t_row, t_col, file='get_rt_data.xls'):
    data = open_excel(file)
    table = data.sheets()[0]
    row = table.row_values(t_row)
    return str(row[t_col]).strip()



#初始化基本数据
def excel_table_byindex_init_basicdata(file= 'get_rt_data.xls',colnameindex=0,by_index=0):
    data = open_excel(file)
    table = data.sheets()[by_index]
    nrows = table.nrows #行数
    ncols = table.ncols #列数
    colnames =  table.row_values(colnameindex) #某一行数据
    #print colnames[0]
    #print colnames
    list =[]
    for rownum in range(3,nrows):
         row = table.row_values(rownum)
         index = 0
         if row:
             app = {}
             for i in range(0,len(colnames)):
                 app[i] = str(row[i]).strip()
             list.append(app)
             index=index+1
    return list

#获取表的行列
def get_table_rows(file= 'get_rt_data.xls'):
    rows = 0
    data = open_excel(file)
    table = data.sheets()[0]
    rows = table.nrows #行数
    return rows

def excel_table_byindex_dynamic(file= 'get_rt_data.xls', colnameindex=0,by_index=0):
    col_title = ('A','B','C','D','E','F','G','H','I','J','K','K')
    xl = win32com.client.Dispatch("Excel.Application")
    #work_book = xl.sheets("Sheet1")
    work_book = xl
    data = open_excel(file)
    table = data.sheets()[by_index]
    nrows = table.nrows #行数
    ncols = table.ncols #列数
    colnames =  table.row_values(colnameindex) #某一行数据
    #print colnames[0]
    #print colnames
    # print work_book.Cells(3,1).value
    # sys.exit()

    list =[]
    for rownum in range(3,nrows):
         row = table.row_values(rownum)
         index = 1
         if row:
             app = {}
             #starttime = datetime.datetime.now()
             for i in range(0,len(colnames)):
                 try:
                    tag_title = col_title[i]+str(index)
                 except Exception, e:
                    print e
                 #print "i:%d\n"%i
                 #app[i] = str(row[i]).strip()
                 #print "rownum:%d,i:%d\n"%(rownum+1, i+1)
                 #app[i] = work_book.Cells(rownum+1, i+1).Value
                 
                 try:
                    app[i] = str(work_book.Cells(rownum+1, i+1).Value).strip()
                 except Exception, e:
                    app[i] = ''
             list.append(app)
             index=index+1
             #print "%s\n"%(app[0])
             #endtime = datetime.datetime.now()
             #print endtime - starttime         
    return list

def excel_table_row_byindex_dynamic(xl, row_index):
    col_title = ('A','B','C','D','E','F','G','H','I','J','K')
    #print colnames[0]
    #print colnames
    # print work_book.Cells(3,1).value
    # sys.exit()

    list =[]
    
    index = row_index
    eci = ''
    country = ''
    unit = ''
    title = ''
    rank = ''
    tag = 'A'+str(index)
    try:
        eci = str(xl.Range(tag).value).strip()
    except Exception,e:
        print e
        return {}
    eci = eci[0:-4]
    app = {}
    for i in range(0,len(col_title)):
        if('A' == col_title[i]):#eci
            app[i] = eci
            continue
        elif('H' == col_title[i]):#unit
            app[i] = lib_help.get_eci_unit(eci)
            continue
        elif('I' == col_title[i]):#country
            app[i] = lib_help.get_eci_country(eci)
            continue
        elif('J' == col_title[i]):#title
            app[i] = lib_help.get_eci_title(eci)
            continue
        elif('K' == col_title[i]):#rank
            app[i] = lib_help.get_eci_rank(eci)
            continue

        try:
           tag_title = col_title[i]+str(index)
        except Exception, e:
           print e
           return {}
        try:
            app[i] = str(xl.Range(tag_title).value).strip()
        except Exception, e:
            print e
            return {}
        if 'None' == app[i]:
            app[i] = ''
    return app
    
    #print "%s\n"%(app[0])
    #endtime = datetime.datetime.now()

#get value of public 
def excel_table_row_public_value(xl, row_index):
    result = ''
    try:
        tag_title = 'E'+str(row_index)
    except Exception, e:
        print e
    try:
        result = str(xl.Range(tag_title).value).strip()
    except Exception, e:
        print e

    if 'None' == result:
        result = ''
    return result

def excel_table_check_today(xl):
    pass

#根据名称获取Excel表格中的数据   参数:file：Excel文件路径     colnameindex：表头列名所在行的所以  ，by_name：Sheet1名称
def excel_table_byname(file= 'file.xls',colnameindex=0,by_name=u'Sheet1'):
    data = open_excel(file)
    table = data.sheet_by_name(by_name)
    nrows = table.nrows #行数
    colnames =  table.row_values(colnameindex) #某一行数据
    list =[]
    for rownum in range(1,nrows):
         row = table.row_values(rownum)
         if row:
             app = {}
             for i in range(len(colnames)):
                app[colnames[i]] = row[i]
             list.append(app)
    return list

def main():
#   tables = excel_table_byindex()
#   for row in tables:
#       print row
#       break
   # r=xlrd.xldate_as_tuple(41828.0, 0)
   # print r
  # reload(sys) 
  # sys.setdefaultencoding("utf-8") 

  # tables = excel_table_byindex_dynamic()
  # for row in tables:
  #     print row
  #     break
  pass

class Product(threading.Thread): #The timer class is derived from the class threading.Thread  
    def __init__(self, xl, start_row, end_row, queue):  
        threading.Thread.__init__(self) 
        self.thread_stop = False  
        self.xl = xl
        self.start_row = start_row
        self.end_row = end_row
        self.data = queue
   
    def run(self): #Overwrite run() method, put what you want the thread do here  
        while not self.thread_stop:
            for i in range(self.start_row,self.end_row):
                self.data.put(excel_table_row_byindex_dynamic(self.xl, i))
            self.thread_stop = True

    def stop(self):  
        self.thread_stop = True

class Consumer(threading.Thread):
    def __init__(self, t_name, queue):
        threading.Thread.__init__(self, name=t_name)
        self.data=queue  
  
    def run(self):  
        for i in range(5):  
            val = self.data.get()  
            print "%s: %s is consuming. %d in the queue is consumed!\n" %(time.ctime(), self.getName(), val)  
            time.sleep(random.randrange(10))  
        print "%s: %s finished!" %(time.ctime(), self.getName())  

def offer(queue, start_row, end_row):  
    #begin = time.clock()
    #queue.put("Hello World")
    xl = win32com.client.Dispatch("Excel.Application")
    for i in range(start_row, end_row):
        queue.put(excel_table_row_byindex_dynamic(xl, i))
        #queue.put(excel_table_row_byindex_dynamic(xl, i))
    #end = time.clock()
    #print "time:%.03f\n"%(end-begin)


if __name__=="__main__":
    print excel_table_get_unit(3,0)
    pass

    # begin = time.clock()
    # threads = []
    # rows = get_table_rows()
    # xl = win32com.client.Dispatch("Excel.Application")
    # work_book = xl
    
    #queue = Queue()
    # step = 200
    # start_row = 4
    # end_row = start_row+step
    # threads.append(Product(xl,start_row, end_row, queue)) 

    # while True:
    #     if(end_row >= rows):
    #         end_row = rows
    #         threads.append(Product(xl,start_row, end_row, queue)) 
    #         break
    #     threads.append(Product(xl, start_row, end_row, queue))
    #     start_row = end_row+1
    #     end_row = start_row+step

    # #启动多线程
    # for i in range(0, len(threads)):
    #     threads[i].start()    

    # for i in range(0,len(threads)):
    #     threads[i].join()

    # list = []
    # for i in range(queue.qsize()):
    #     list.append(queue.get())

    # step = 50
    # start_row = 4
    # end_row = start_row+step
    # q = Queue()
    # p_count = int(math.ceil((rows-3)/step))
    # print p_count
    # for i in range(p_count):
    #     p = Process(target=offer, args=(q, start_row, end_row))
    #     p.start()
    #
    #
    #
    # while True:
    #     print q.get()

    # is_run  = True
    # 
    # while(is_run):
    #     try:
    #         list.append(queue.get())
    #     except Exception, e:
    #         is_run = False

    # print list

    # for i in range(4,rows):
    #     excel_table_row_byindex_dynamic(xl, i)
    # end = time.clock()
    # print "time:%.03f\n"%(end-begin)
    # while True:
    #     main()
    #     time.sleep(3)

    # end = time.clock()
    # print "time:%.03f\n"%(end-begin)