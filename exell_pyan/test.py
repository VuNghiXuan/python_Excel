import xlwings as xl

sh = xl.sheets.active
sh.range("A1").value = [["r1c1", "r1c2","r1c3"], ["r2c1", "r2c2", "r2c3"], ["r3c1", "r3c2","r3c3"]]
# sh.range("D1").value = [1,2,3,4,5]

range_cell = sh.range("A1:C3").value
print("===", range_cell)
for i in range(len(range_cell)):
    # for j  in range(len(range_cell)):
    print(range_cell[i])
sh.range("E1").value= range_cell

# >>> import xlwings as xw
    # >>> sht = xw.Book().sheets[0]
    # >>> sht.range('A1').value = [[1, 2], [3, 4]]
    # >>> sht.range('A1').value
    # 1.0
    # >>> sht.range('A1').options(ndim=1).value
    # [1.0]
    # >>> sht.range('A1').options(ndim=2).value
    # [[1.0]]
    # >>> sht.range('A1:A2').value
    # [1.0 3.0]
    # >>> sht.range('A1:A2').options(ndim=2).value
    # [[1.0], [3.0]]

#2 Converts the strings  in cell to uppercase
# sh.range("A1").value = range_cell.upper()

#3 First character capitalization
# sh.range("A1").value = range_cell.capitalize()

#4 Remove all spaces
# sh.range("A1").value = range_cell.replace(" ", "")

# #5 Delete the space on the left
# sh.range("A1").value = range_cell.lstrip()

# #6 Delete the space on the right
# sh.range("A1").value = range_cell.rstrip()

# #7 Delete the space on the left and right
# sh.range("A1").value = range_cell.strip()

# #8 Capitalize the first letter of each letter
# sh.range("A1").value = string.capwords(range_cell)