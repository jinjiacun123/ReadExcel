# -*- coding: utf-8 -*-
import sys
import os
import ConfigParser
class Db_Connector:
    cf = None
    fd = None
    
    def __init__(self, config_file_path):
        self.fd = config_file_path
        self.cf = ConfigParser.ConfigParser()
        self.cf.read(config_file_path)
        s = self.cf.sections()
        #print 'section:', s
        #o = cf.options("baseconf")
        #print 'options:', o
        v = self.cf.items("baseconf")
        #print 'db:', v
        #print self.cf.get('baseconf','data_url')
    #    db_host = cf.get("baseconf", "host")
    #    db_port = cf.getint("baseconf", "port")
    #    db_user = cf.get("baseconf", "user")
    #    db_pwd = cf.get("baseconf", "password")
    #    print db_host, db_port, db_user, db_pwd
    #    cf.set("baseconf", "db_pass", "123456")
    #    cf.write(open("config_file_path", "w"))
    
    #write
    def write(self, pre, opt, data):
        self.cf.set(pre, opt, data)
        self.cf.write(open(self.fd, "w"))
 
     #read
    def read(self, pre, opt):
        if self.cf.has_option(pre, opt):
            return self.cf.get(pre, opt)
        else:
            return ''
    
if __name__ == "__main__":
  f = Db_Connector("my.ini")
  #read data_url
  print f.read('baseconf', 'data_url')
  f.write('baseconf', 'host', 'localhost')

