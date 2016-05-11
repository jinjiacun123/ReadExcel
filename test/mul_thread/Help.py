# -*- coding: utf-8 -*-
import threading
from time import ctime

class Setting:
    
    #获取数据条目数量
    def get_data_amount(self):
        return 706
        
class Excel:
    
    #读取excel行数
    def get_rows(self):
        pass
    
    #读取excel列数
    def get_cols(self):
        pass
        
    #读取excel指定(行,列)的值
    def get_value(self, i_row, i_col):
        pass
    
class Ini:
    
    #获取值
    def get(self):
        pass
    
    #写入值
    def set(self):
        pass

#任务
class Task(threading.Thread):
    
    def __init__(self, func, args, name=''):
        threading.Thread.__init__(self)
        self.name = name
        self.func = func
        self.args = args
        
    def getResult(self):
        return self.res
        
    def run(self):
        ctime()
        self.res = apply(self.func, self.args)
        ctime()
            
    
    