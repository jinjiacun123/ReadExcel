# -*- coding: utf-8 -*-
import urllib2, urllib

def my_post(url, data):
    f = urllib2.urlopen(
            url     = url,
            data    = urllib.urlencode(data)
      )
    return f.read()  

def main():
    url = 'http://192.168.1.248:86/ReutersHandler.ashx'
    data = {'op':'addfinancedata',
        		'content':'澳大利亚11月出口(月率)',
        		'before':'-3%',
        		'prediction':'-3%',
        		'result':'1%',
        		'country':'澳大利亚',
        		'rank':'5',                       
          }
    print my_post(url, data)

if __name__=="__main__":
    main()