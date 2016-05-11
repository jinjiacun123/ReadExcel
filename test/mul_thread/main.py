# -*- coding: utf-8 -*-
import const
import Help

def main():
    #初始化任务
   setting = Help.Setting()
   ##获取需要读取的数据总条数
   data_amount = setting.get_data_amount()
   print data_amount
    
if __name__ == "__main__":
#    #任务数组
#    read_task = []
#    
#    main()
    import win32com.client
    xl = win32com.client.Dispatch("Excel.Application")
    print 'test'