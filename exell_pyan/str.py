import xlwings as xl
import string

sht = xl.sheets.active
# rgn= sht.range("A1:A3").value
# print(rgn)

# sht.range("A1:A3").options(transpose = True).value = [str.lower() for str in rgn] # in thuong
# sht.range("A1:A3").options(transpose = True).value = [str.upper() for str in rgn] # in hoa
# sht.range("A1:A3").options(transpose = True).value = [str.capitalize() for str in rgn]# in hoa chữ cái đầu
rng = sht.range("A1").value
sht.range("A2").options(transpose = True).value = rng.title()# in hoa chữ họ và tên 
# sht.range("A1:A3").options(transpose = True).value = [str.replace(" ","") for str in rgn]# bỏ ký tự trống
# sht.range("A1:A3").options(transpose = True).value = [str.strip() for str in rgn]# xóa dấu cách 2 đầu trái phải
#str.lstrip() xóa bên trái, str.rstrip() xóa bên phải
# sht.range("A1:A3").options(transpose = True).value = [string.capwords(str) for str in rgn]# xóa dấu cách 2 đầu trái phải

# # Sử dụng splitlines cho các cell xuống dòng
# sht.range("A1").value = ["dang" + "\n" + "thanh" + "\n" +"vu"]
# print(sht.range("A1").value)
# # đưa 1 dòng thành 3 dòng
# range_split = sht.range("A1").value
# print("range_split", range_split)
# sht.range("B1").options(transpose=True).value = range_split.splitlines()

# # Hàm str.center(width[, fillchar]), width: len(str); fillchar: loại ký tự
# sht.range("c1").value = range_split.center(14, "*") 
# sht.range("c2").value = range_split.center(15, "*") 
# print(len(range_split.center(14, "*")))#. Trường hợp ký chiều dài + thêm 1"*"=14 thì thêm vào sau chuỗi trước
# sosao = str(range_split.center(14, "*").count("*"))
# print("Số '*' trong chuỗi: ", sosao)
