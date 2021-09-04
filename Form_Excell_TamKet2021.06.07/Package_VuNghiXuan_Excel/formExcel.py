from tkinter import *
from tkinter.filedialog import askopenfilename
from Package_VuNghiXuan_Excel.useFunc import read_file_from_txt#from Package_VuNghiXuan_Excel.check 
from Package_VuNghiXuan_Excel.useFunc import write_files_wb_to_txt
# from Package_VuNghiXuan_Excel.useFunc import Update_lisbox
# import tkMessageBox
import os
import ctypes

# import tkinter
def remove_out_list():
    value= listbox.get(ANCHOR)
    if value == "":                
        # Display a progress bar thực hiện khi doawload haoc xử lý %
        # bar = EasyDialogs.ProgressBar(maxval=100)
        # for i in range(100):
        #     bar.inc()
        # del bar

        returnMessage = ctypes.windll.user32.MessageBoxW(0, "Nhấn 'OK' để thực hiện xóa!!!", "Bạn có muốn xóa toàn bộ file ra khỏi danh sách không?",1)#"Nhấn 'No' để hủy",
        # print(returnMessage)
        if returnMessage==1 :
            listbox.delete(0,'end')
                   
    else:
        # Remove elment in listbox            
        idx = listbox.get(0, END).index(value)
        listbox.delete(idx)

def addfile_in_listbox():
    filename = askopenfilename() # open seach file
    listbox.insert(END,filename) # add file vào lisbox
    # Update_lisbox(listbox)
    # index_listbox = listbox.get(0, END).index(filename)#Số phần tử trong listbox
    
    # files_wbs = []
    # for i in range(index_listbox + 1):
    #     # Xóa emptyrows       
    #     if listbox.get(i)!="":      
    #         files_wbs.append(listbox.get(i))
    #     # print ("gdfgdfgdfg---listbox.get(i):", listbox.get(i))
    #     # files_wbs.append(listbox[i])
    
    # # print ("gdfgdfgdfg---files_wbs:", files_wbs)
    # write_files_wb_to_txt(file_txt, files_wbs)
    # list_file = []
    # print()
    # print("gdfgdfgdfg---idx_add:",idx_add, file_txt)
    # return filename

windown = Tk()
windown.title("Sofe of VuNghiXuan")
windown.geometry("400x300")
# windown.config(bg=#'446644')

file_txt = os.getcwd() + "\Package_VuNghiXuan_Excel\data.txt"
wbs = read_file_from_txt(file_txt)

#tao listbox
listbox = Listbox(windown, w=300)#, w=50
total = len(wbs)
for i in range(total):
    listbox.insert(END,wbs[i])
listbox.pack(anchor='w')#

# Create Button remove trong listbox
btn_remove = Button(windown, text="Remove file", command = remove_out_list).pack()

# Create Button remove trong listbox
btn_add = Button(windown, text="Open and add file ... ", command = addfile_in_listbox).pack()

#, pandy=20 command = remove_out_list
# # btn_remove.grid()#.grid(row=1, column=0)
# show=Label(windown)

# listbox.pack()
# btn_remove.grid()
# show.pack()
# frm1= Tk(windown)
windown.mainloop()

# # E:\code\python\py_exell\test1.xlsm
# def remove_file(listbox_wbs):#  tkMessageBox.showinfo( "Hello Python", "Hello World")
#     # print("Value list box:" + str(listbox_wbs.get()))
#     # print("Value list box:" + str(listbox_wbs.get(ANCHOR)))
#     # t = StringVar(listbox_wbs.get(ANCHOR))
#     # t= listbox_wbs.delete(0,'end')
#     # print(t)
#     pass
#     # pass
# def CurSelet(evt):
#     values = [listbox_wbs.get(idx) for idx in listbox_wbs.curselection()]
#     print (', '.join(values))

# def show_form():
#     window = Tk()
#     window.title("Sofe Exell VuNghiXuan")
#     window.geometry("800x600")

#     # create list file quản lý: gồm danh sách file đã lưu 
#     frm1=Frame(window)
#     frm1.pack(fill=BOTH)

#     # Show danh sách đã lưu
#     lb_listbox = Label(frm1, text="Danh sách file đã lưu ", fg = "blue", font = ("Arial", 10))
    
#     listbox_wbs = Listbox(frm1)#, width=50
#     # # +str(listbox.get(ANCHOR), có thể lấy từ vị trí con trỏ chăng?    
#     file_txt = os.getcwd() + "\Package_VuNghiXuan_Excel\data.txt"
#     # wbs = read_file_from_txt(file_txt)
#     wbs = ['mot', "hai"]
    
#     # for i in range(len(wbs)):
#     #     listbox_wbs.insert(END, wbs[i])
#     for items in wbs:
#         listbox_wbs.insert(END,items)
    
#     lb_listbox.grid(row=0, column=0)
#     for i in range(listbox_wbs):
#         print(listbox_wbs[i])



    
#     # listbox_wbs.bind('<<ListboxSelect>>',CurSelet)#'<<ListboxSelect>>'
    
#     # listbox_wbs.bind('<<ListboxSelect>>',CurSelet)
#     # listbox_wbs.place(x=32,y=90)
    
#     # button for remove file



#     #đỂ CODE SASUSASUAUSUAS
    
#     # bnRemove = Button(frm1, text ="Remove", command = remove_file(listbox_wbs))
    
#     # bnRemove.grid(row=1,column=1)






#     # listbox_wbs.insert(1, "Python")
#     # listbox_wbs.insert(2, "Perl")
#     # listbox_wbs.insert(3, "C")
#     listbox_wbs.grid(row=1, column=0)
#     # create select file 
#     # Tk().withdraw()
#     # filename = askopenfilename() # show an "Open" dialog box and return the path to the selected file
#     # #ile = filedialog.askopenfilename(filetypes = (("Text files","*.txt"),("all files","*.*")))
#     # print(filename)

#     # Add label
#     # lb_addFile = Label(form, text="Thêm file: ", fg = "blue", font = ("Arial", 20))
#     # lb_addFile.grid(row=0, column=0)#(column=0,row=0)

#     # lb_filename = Label(form, text="Chọn file: ", fg = "blue", font = ("Arial", 20))#, fg="blue", font=("Arial", 20)
#     # lb_filename.grid(row=1, column=0)#(column=0,row=0)
#     window.mainloop()
def main():
    show_form()
if __name__=='__main__':   #__name__=='__main__': khi nào là hàm main() nằm trong bảng code náy mới dc chạy
    main()