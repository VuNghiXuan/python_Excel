import pandas as pd
import os
# excel_file="E:/code/python/py_exell/Data_20210424/File_Test_Excel/movies.xls"
# xlsx = pd.ExcelFile(excel_file)
# movies_sheets = []
# for sheet in xlsx.sheet_names:
#     print("ten sheet: ",sheet)#liệt kê tên sheet
#     movies_sheets.append(xlsx.parse(sheet))
#     movies = pd.concat(movies_sheets)
#     print()
# print(movies_sheets.head ())
# # print("dgjgjdfgjdfgd", movies_sheets)
def get_sheetNames_inPdExl(excel_file):
    xlsx = pd.ExcelFile(excel_file)
    sheetNames = xlsx.sheet_names
    return sheetNames
# def get_SheetWithData_inPd_Exl(excel_file, *args): #*args, **kwargs
# 	# excel_file="E:/code/python/py_exell/Data_20210424/File_Test_Excel/movies.xls"
#     xlsx = pd.ExcelFile(excel_file)
#     sheetNames = []
#     data_in_sheets = []
#     for sheet in xlsx.sheet_names:
#         sheetNames.append(sheet)
#         data_in_sheets.append(xlsx.parse(sheet))
#         data_conect = pd.concat(data_in_sheets)
#     	# print("ten sheet: ",sheet)#liệt kê tên sheet
# 		# data_in_sheets.append(xlsx.parse(sheet))

		
# 		# data_in_sheets = pd.concat(movies_sheets)
    
#     print("fhgfdhghdfgh",sheetNames, xlsx.sheet_names)#, data_in_sheets
# # print("dgjgjdfgjdfgd", movies_sheets)
#     return sheetNames, data_in_sheets



def get_SheetWithData_inPd_Exl(excel_file,All=None, SheetNames=None, Data=None, conect_data=None): #*args, **kwargs
	# excel_file="E:/code/python/py_exell/Data_20210424/File_Test_Excel/movies.xls"
    xlsx = pd.ExcelFile(excel_file)
    sheetNames = xlsx.sheet_names
    data_in_sheets = []
    for sheet in xlsx.sheet_names:
        # sheetNames.append(sheet)
        data_in_sheets.append(xlsx.parse(sheet))
        # data_conect = pd.concat(data_in_sheets) #Kết nối các sheet có cùng dữ liệu thành 1 bảng
    	# print("ten sheet: ",sheet)#liệt kê tên sheet
    
    if All==True:# show toàn bộ dữ liệu và tên sheet, ko connect các sheet
        return sheetNames, data_in_sheets
    elif SheetNames==True:#only show tên sheet 
        return sheetNames
    elif Data==True:# Show data
        if conect_data==True: # Show data và connect
            data_conect = pd.concat(data_in_sheets) #Kết nối các sheet có cùng dữ liệu thành 1 bảng
            return data_in_sheets
        return data_in_sheets#Only Show data 
    else: 
        if conect_data==True:
            data_conect = pd.concat(data_in_sheets) #Kết nối các sheet có cùng dữ liệu thành 1 bảng
            return sheetNames, data_in_sheets
        return sheetNames, data_in_sheets

def printFile_wb(excel_file, *datas):
	# print(f"Excel_file: {excel_file}")
	#, SheetNames=None, Data=None, conect_data=None
    datas = get_SheetWithData_inPd_Exl(excel_file,SheetNames=True)#, Data=True)#, Data=True
    print("dfgjkdfjgkdfjk",datas)
    i=1
    for sheet in datas:#sheets.items()
        print(f"Sheet{i}: {sheet}")
        i=i+1
        # print(f"dat: {datas[1]}")
	# printPetNames(excel_file, sheets= sheets[0])

#Lấy tên tệp
def get_FilenameExtention(file):
	FilenameExtention = os.path.basename(file)#os.path.splitext(pathfile)[0]
	return FilenameExtention

def main():
    # excel_file="E:/code/python/py_exell/Data_20210424/File_Test_Excel/movies.xls"
    # t=get_SheetWithData_inPd_Exl(excel_file, SheetNames=True)#"sn",SheetNames=True
    # t = get_sheetNames_inPdExl(excel_file)
    # print(t)

    excel_file = "I:/Code/Python/pyExell/python_excel/Data_20210426/File_Test_Excel/movies.xls"
    t= get_FilenameExtention(excel_file)
    print(t)

if __name__== '__main__':
    main()