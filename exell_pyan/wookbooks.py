import xlwings as xl
path_file_open = "I:\Code\Python\pyExell\learnPyExell1.xlsx"
app = xl.App(visible = True, add_book= False)
#  ====>Step 1: open wookbook, nhược điểm ko thể mở khi dc 1 chương trình nào đó đã gọi trước
# xl.Book("I:\Code\Python\pyExell\learnPyExell.xlsx")
    #xl.Book(): Thực hiện tạo Wookbook mới
    #xl.Book("I:\Code\Python\pyExell\Sofe_DieuChinhDutoan_Q3Q4(Bailey).xlsm"): Mở wookbook có path

#  ====>Step 2: open wookbook, tạo ra 02 book
# app = xl.App() #creat 1 book from method xl.App()
# app.books.add() ##creat 1 book from method app.books.add()

#  ====>Step 3: open wookbook, tạo ra 01 book với (add_book= False, nếu ko có mặc nhiện là True sẽ xuật hiện 02 book1 and book2
# app = xl.App(add_book= True) #creat 1 book from method xl.App()
# app.books.add() ##creat 1 book from method app.books.add()

#  ====>Step 4: open wookbook có path app = xl.App(add_book= True) and book1
# app = xl.App(add_book= True) #creat 1 book from method xl.App()
# app.books.open(path_file_open) ##creat 1 book from method app.books.add()

#  ====>Step 4: open wookbook có path app = xl.App(add_book= False) not and book1
# app = xl.App(add_book= False) #creat 1 book from method xl.App()

# app.books.open(path_file_open) ##creat 1 book from method app.books.add()

# Note: 02 method open book
    # 1:
# wb = xl.Book(path_file_open)
wb = app.books.open(path_file_open) #This method cho phép save, close, còn Books ko đóng dc
sh1 = wb.sheets("Vu")
sh1.range("B1:C12").value = "Tui"
    # 2:
# wb = app.books.open(path_file_open)
wb.save()
wb.close()
app.quit()

# wb1 = app.books.open(path_file_open)
# sh1 = wb.sheets("Vu")
# print(sh1.range("B1:C12").value)



