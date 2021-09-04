import xlwings as xl
# path_file_open = "I:\Code\Python\pyExell\learnPyExell1.xlsx"
# app = xl.App(visible = True, add_book= False)
# sh1 = xl.Range("B1:C2").value # Range("B2:C12")
# sh2 = xl.Range((1,2),(2,3)).value # Range("B2:C12")
sh = xl.sheets.active
sh.range("A1").value = [1,2,3]#cách gán theo cột, tức là 1 dòng
rgn1 = sh.range("A1:A3").value

sh.range("A1").value = [[1],[2],[3]]#cách gán theo dòng, tức là 1 cột
rgn2 = sh.range("A1:A3").value
print(rgn1, rgn2)
# print(sh2)