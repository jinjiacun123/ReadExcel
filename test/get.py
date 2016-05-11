# -*- coding: utf-8 -*-
import win32api
import sys
#eursl
reload(sys) 
sys.setdefaultencoding("utf-8") 

e_msg = win32api.FormatMessage(-2147418111)
print e_msg.decode('CP1251')
print e_msg.decode('gbk')