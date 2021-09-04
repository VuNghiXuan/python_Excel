import os
import xlwings as xw
from Package_VuNghiXuan_Exell import File_Exell
print(File_Exell.__name__)
# from Package_VuNghiXuan_Exell.File_Exell import *#Cách gọi tất cả các hàm trong gói Package_VuNghiXuan_Exell
# from Package_VuNghiXuan_Exell import * 

# Infor_workbook
file_path = "D:\\ThanhVu\\code\\python\\pyExell" ####### Chú ý thay dg dan khi vào fiel khác
file_name = "test.xlsx"
# pathfull = os.path.join(file_path,file_name)
# wb1 = xw.Book(pathfull)
obj_workbook = File_Exell.Workbook(file_path, file_name)
print("[[[[[[[obj_workbook.pathfull", obj_workbook.pathfull)
wb2 = obj_workbook.open()

# wb1 = xw.Book(obj_workbook.pathfull)
# check thư viện
# wb2 = File_Exell.Workbook.wb.open()
print("Mở file từ class OK!!"+ "\n")
print("FileName qua class OK!!: "+ File_Exell.Workbook.name+ "\n")
print("FilePath qua class OK!!: "+ File_Exell.Workbook.path+ "\n")


# print("Số sheet theo cách gán trực tiếp:"+ str(num_sheet1) + "\n")

# # num_sheet2 = wb2.api.Sheets.Count
# print("Số sheet theo cách gán qua Class:"+ str(num_sheet2) + "\n")

# # Note
# print(dir(File_Exell)) # Kết quả ['Workbook', '__builtins__', '__cached__', '__doc__', '__file__', '__loader__', '__name__', '__package__', '__spec__', 'xw']
# # __hàm__: là hàm python. Trong đó '__file__':
# print(File_Exell.__name__)# cách thực hiện kiếm đường dẫn đến hàm từ package

# num_sheets = wb.api.Sheets.Count
# print(num_sheets)
