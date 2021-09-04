import os
from os.path import isfile
import pathlib
import socket
from tkinter import messagebox
from typing import Dict
import pandas as pd
from Package_VuNghiXuan_Excel.useFunc import *

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

#  Học 
# def read_Files_Excell(wb):
#     excels = []	
#     excel = read_File_Excell_on_ThisComputer(wb)
#     excels.append(excel)  
#     return  excels
     
class List_App():
    def __init__(self, title, playlists, click):
        self.title = title
        self.playlists = playlists
        self.click = True

    def __delattr__(self):
        del self
        
          
class Playlist:
    def __init__(self, title, excels, click):
        self.title = title
        self.excels = excels
        self.click = False

    def __len__(self):
        return len(self.excels)
        
    def __delattr__(self):
        del self     
        

class Excel:
    
    def __init__(self, name_Exten, link, sheets, click):
        self.name_Exten = name_Exten 
        self.link = link
        self.sheets = sheets
        self.click = False

    def __delattr__(self):
        del self
        
    #     self.click = False


    # def open(self):
	# 	webbrowser.open(self.link)
	# 	self.click = True
	# t.shts[0].data: cách lấy tên sheet trong class

class Sheet():
    def __init__(self, name , data, num_Rows, num_Cols, click):#, exist
        self.name = name 		
        self.data= data
        self.num_Rows= num_Rows
        self.num_Cols= num_Cols 
        self.click = False
        # self.exist = False

    def __delattr__(self):
        del self
          
class tree():
    def __init__(self, space, branch, tee, last):
        self.space =  '    '
        self.branch = '│   '
        self.tee =    '├── '
        self.last =   '└── '
        
def create_playlists_Excels(): # playlists_name = 
    titles = title_playlists()
    title = titles[0]
    playlists = read_playlists()
    playlists_Excels = List_App(title, playlists, False)
    return playlists_Excels

def read_playlists():
    playlists = []
    list_files_from_ThisCom, list_files_from_OtherCom = read_file_txt_return_lists_for_ThisCom_OtherCom()
    
    # add playlists ThisComputer
    playlist = read_playlist(list_files_from_ThisCom)
    playlists.append(playlist)    

    # add playlists OtherComputer
    playlist = read_playlist(list_files_from_OtherCom)
    playlists.append(playlist)
    return playlists    

def read_file_txt_return_lists_for_ThisCom_OtherCom():
    file = os.getcwd() + "\Package_VuNghiXuan_Excel\data.txt" 
    list_files_from_ThisCom=[]
    list_files_from_OtherCom=[]
    with open(file, "r") as file:		
        total_files = int(file.readline())        
        for i in range(total_files):			
            wb = file.readline()
            wb = wb.replace("\n","")            
            # Tạo danh sách các file on ThisComputer
            if file_Exists(wb)==True:
                list_files_from_ThisCom.append(wb)               
            # Tạo danh sách các file on OtherComputer
            else :
                list_files_from_OtherCom.append(wb)                
    return [list_files_from_ThisCom, list_files_from_OtherCom]

def read_playlist(wbs):
    # title, excels
    total_file = len(wbs)
    if total_file>0:
        if isfile(wbs[0]): # check tồn tại file đầu tiên tức là ==>ThicComputer
            tit_playlist = title_playlists()
            title = tit_playlist[1]
            excels = read_files_Excels(wbs)
            playlist = Playlist(title, excels, False)
        else:
            tit_playlist = title_playlists()
            title = tit_playlist[2]
            excels = read_files_Excels(wbs)
            playlist = Playlist(title, excels, False)
    else:
        playlist= None
        # print(playlist) 
    return playlist

def title_playlists():
    title_playlists = "Playlists: "
    title_ThisCom = f"File on This Computer ({nameComputer()}): "
    title_OtherCom = f"File on Other Computer: "
    titles = [title_playlists, title_ThisCom, title_OtherCom]
    return titles

def read_files_Excels(wbs):
    excels = []
    total_file = len(wbs)
    if total_file>0:
        for wb in wbs:
            excel = read_file_Excel(wb)
            excels.append(excel)
    return excels

def read_file_Excel(wb):
    name_Exten = get_FilenameExtention(wb)
    link = wb
    sheets = get_sheets(wb)
    excel = Excel(name_Exten, link, sheets, False)
    return excel

def get_sheets(wb):    
    sheets=[]
    if isfile(wb):  # kiểm tra sự  tồn tại file: isfile(wb)
        xlsx = pd.ExcelFile(wb) # movie=pd.read_excel(wb) # đọc dữ liệu mặc định sheet đầu tiên
    
        for i in range(len(xlsx.sheet_names)):
            sheet_name= xlsx.sheet_names[i]        
            sheet = get_sheet(xlsx, sheet_name)
            sheets.append(sheet)      
            # sheets.append(xlsx.parse(sheet)) # Sử dụng phương thức để đọc toàn bộ dữ liệu 1 sheet       
    else:
        sheets.append(None) 
    return sheets

def get_sheet(xlsx, sheet_name):    
    sheetName = sheet_name
    dataSheet = xlsx.parse(sheet_name)
    num_Rows = dataSheet.shape[0]
    num_Cols = dataSheet.shape[1]
    sheet = Sheet(sheetName, dataSheet, num_Rows, num_Cols, False)
    return sheet
    
def define_list_App():
    list_App = create_playlists_Excels()
    playlists = list_App.playlists
    playlist_ThisCom = playlists[0]
    if playlist_ThisCom != None:
        excels_ThisCom = playlist_ThisCom.excels
    else:
        excels_ThisCom = None
    playlist_OtherCom = playlists[1]
    if playlist_OtherCom != None:
        excels_OtherCom = playlist_OtherCom.excels
    else: 
        excels_OtherCom = None
    # excels = playlist.excels
    # sheets = excels.sheets
    return list_App, playlists, playlist_ThisCom, excels_ThisCom, playlist_OtherCom, excels_OtherCom

def main():	
    list_App, playlists, playlist_ThisCom, excels_ThisCom, playlist_OtherCom, excels_OtherCom = define_list_App()
    # list_App, excels, sheets = define_list_App()
    # t = len(list_App.playlist)

    print(excels_ThisCom[0].name_Exten)
    # print("Trước xóa", playlists_Excel.playlists[0].excels[1])
    # del playlists_Excel.playlists[0].excels[1]
    # print("Sau xóa", playlists_Excel.playlists[0].excels[1])

if __name__=='__main__':   #__name__=='__main__': khi nào là hàm main() nằm trong bảng code náy mới dc chạy
    main()	
	