import xlwings as xl

sh = xl.sheets.active
sh.range("A1:A3").value = [[1, 2], [3, 4]]

range_cell = sh.range("A1:A3").value
#2 Converts the strings  in cell to uppercase
# sh.range("A1").value = range_cell.upper()

#3 First character capitalization
# sh.range("A1").value = range_cell.capitalize()

#4 Remove all spaces
# sh.range("A1").value = range_cell.replace(" ", "")

#5 Delete the space on the left
sh.range("A1").value = range_cell.lstrip()

#6 Delete the space on the right
sh.range("A1").value = range_cell.rstrip()

#7 Delete the space on the left and right
sh.range("A1").value = range_cell.strip()

#8 Capitalize the first letter of each letter
sh.range("A1").value = string.capwords(range_cell)