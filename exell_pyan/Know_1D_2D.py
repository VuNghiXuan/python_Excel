import xlwings as xl

sht = xl.sheets.active
rgn= sht.range("A1:A3").value
print(rgn)
rng.options(transpose = True).value = [str.lower for str in range(rgn)]