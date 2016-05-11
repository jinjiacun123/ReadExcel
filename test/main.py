# -*- coding: utf-8 -*-
import time
import win32com.client

tag_list = ("A","B","C","D","E","F","G","H","I","J","K")
def main():
	xl = win32com.client.Dispatch("Excel.Application")
	index = 4

	item = {
		"eci":xl.Range(tag_list[0]+str(index)).value,
		"ctbtr_1ll":xl.Range(tag_list[1]+str(index)).value,
		"ctbtr_1":xl.Range(tag_list[2]+str(index)).value,
		"offc_code2":xl.Range(tag_list[3]+str(index)).value,
		"gn_txt16_4":xl.Range(tag_list[4]+str(index)).value,
		"cf_date":xl.Range(tag_list[5]+str(index)).value,
		"cf_time":xl.Range(tag_list[6]+str(index)).value,
		"gv3_text":xl.Range(tag_list[7]+str(index)).value,
		"country":xl.Range(tag_list[8]+str(index)).value,
		"title":xl.Range(tag_list[9]+str(index)).value,
		"rank":xl.Range(tag_list[10]+str(index)).value,
	}
	print item

if __name__ == "__main__":
	#while True:
    begin = time.clock()
    main()
