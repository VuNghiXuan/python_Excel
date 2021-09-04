import os
import pathlib
import socket
from tkinter import messagebox
from typing import Dict
import pandas as pd
from Package_VuNghiXuan_Excel.useFunc import *

# if 'key1' in dict.keys():
#   print "blah"
# else:
#   print "boo"


# Add key, value vào dict
def createDict_from_list_assign_False(lis):
	dict={}
	for i in lis:
		dict[i] = False
		# print("Trước:",i, dict[i])
	return dict

def isTrueorFalse_inDict(*dicts):
	for dict in dicts:
		for i in dict.keys():
			if dict[i]==False:
				dict[i]=True
			else:
				dict[i]=False
			# print("Sau:",i, dict[i])
	return dicts

class Excell_NghiXuan:
	def __init__(xlsx, name_Exten, link, shts):
		xlsx.name_Exten = name_Exten
		xlsx.link = link
		xlsx.shts = shts

class Sheet_NghiXuan():
	def __init__(sht, name, data, rows, cols):#
		sht.name= name
		sht.data= data
		sht.rows= rows
		sht.cols= cols

def read_File_Excell(excel_file):
	xlsx = pd.ExcelFile(excel_file)
	name_Exten = get_FilenameExtention(excel_file)
	link = excel_file
	# FilenameExtention = os.path.basename(file)
	
	exl = Excell_NghiXuan(name_Exten, link, shts)

def readFile_Excell(excel_file):
	data=[]	
	xlsx = pd.ExcelFile(excel_file)
	
	for i in range(len(xlsx.sheet_names)):		
		name=xlsx.sheet_names[i]
		data.append(xlsx.parse(xlsx.sheet_names[i]))		
		rows=data[i].shape[0]
		cols=data[i].shape[1]
		sh=Sheet_NghiXuan(name, data, rows, cols)
		
		# print (data[i].shape)

	# for sheet in xlsx.sheet_names:		
	# 	data.append(xlsx.parse(sheet))
	# 	# read_sheet_into_class_sheet_NghiXuan(sheet)
	# 	df = xlsx.parse(xlsx.sheet_names[i for i in range[len(xlsx.sheet_names)]])
	# 	df.shape	

	return xlsx

# def read_sheet_into_class_sheet_NghiXuan(sheet):
# 	sh = Sheet_NghiXuan(sheet)
# 	return sh
	
def main():
	excel_file = "I:/Code/Python/pyExell/python_excel/Data_20210509/File_Test_Excel/movies.xls"
	t = readFile_Excell(excel_file)
	# print (t)
	# sh=Sheet_NghiXuan.name
	# print("sheetname", sh)
	
	# f={"tui":True}
	# a=['e',"f","g","h","j"]
	# b=['a',"b","c","d","e"]
	# t1=createDict_from_list_assign_False(a)
	# t2=createDict_from_list_assign_False(b)
	# print(t1)
	# print(t2)
	# t=isTrueorFalse_inDict(t1,f)
	# print(t)

if __name__=='__main__':   #__name__=='__main__': khi nào là hàm main() nằm trong bảng code náy mới dc chạy
    main()	
	