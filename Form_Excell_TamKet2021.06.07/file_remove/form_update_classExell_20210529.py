from genericpath import isdir
from tkinter import *
import tkinter.ttk as version
import os
# import ctypes
from tkinter import messagebox
from tkinter.filedialog import askopenfilename

# from Package_VuNghiXuan_Excel.useFunc import read_file_from_txt  
# from Package_VuNghiXuan_Excel.useFunc import write_files_wb_to_txt
from Package_VuNghiXuan_Excel.useFunc import *
from Package_VuNghiXuan_Excel.excelPandas import *


def show_form():
    class formExcel(Frame):
        
        def placeGui(obj, event):

            obj.update()       
            objW = obj.winfo_width()
            objH = obj.winfo_height()           
            obj.master.title("App's Excel VuNghiXuan") 

            # Configure obj.changeX, obj.changeY: Dùng dự phòng trường hợp di chuyển toàn bộ Frame xuống dưới
            obj.tab_X = 10
            obj.tab_Y = 10
            obj.changeX = 0
            obj.changeY = 0


            # 1.Tiêu đề Listbox+File            
            obj.H_lbTieudelistbox = 25
            obj.W_lbTieudelistbox = objW*0.2
            obj.posX_lbTieudelistbox = 10
            obj.posY_lbTieudelistbox = 5*obj.tab_Y
            obj.pos_ABCD_lbTieudelistbox = pos_Rectangular(obj.H_lbTieudelistbox, obj.W_lbTieudelistbox, obj.posX_lbTieudelistbox, obj.posY_lbTieudelistbox) 
            
            obj.lbTieudelistbox.place(height = obj.H_lbTieudelistbox, width =obj.W_lbTieudelistbox, x=obj.posX_lbTieudelistbox, y=obj.posY_lbTieudelistbox)
            
            # 2.Listbox+File        
            obj.H_listboxFile = 800
            obj.W_listboxFile = objW*0.2
            obj.posX_listboxFile = obj.posX_lbTieudelistbox
            obj.posY_listboxFile = obj.pos_ABCD_lbTieudelistbox[3][1]+obj.tab_Y # lấy yD của lbTieudelistbox + obj.tab_Y =10
            obj.pos_ABCD_listboxFile = pos_Rectangular(obj.H_listboxFile, obj.W_listboxFile, obj.posX_listboxFile, obj.posY_listboxFile) 
            
            obj.listboxFile.place(height = obj.H_listboxFile, width =obj.W_listboxFile, x=obj.posX_listboxFile, y=obj.posY_listboxFile)
            obj.listboxFile.bind("<<ListboxSelect>>", obj.onSelect_listboxFile)
           
            #3a.scrollbar listbox
            obj.H_scroll_Y_listboxFile = obj.H_listboxFile
            obj.W_scroll_Y_listboxFile = 20
            obj.posX_scroll_Y_listboxFile = obj.pos_ABCD_listboxFile[1][0]#xB_Listbox
            obj.posY_scroll_Y_listboxFile = obj.pos_ABCD_listboxFile[1][1]

            # Tọa độ scroll_listboxFile_Y tính từ obj.posXY_liH_listboxFile
            obj.pos_ABCD_scroll_Y_listboxFile = pos_Rectangular(obj.H_scroll_Y_listboxFile, obj.W_scroll_Y_listboxFile, obj.posX_scroll_Y_listboxFile, obj.posY_scroll_Y_listboxFile)
            obj.scroll_Y_listboxFile.place(height = obj.H_scroll_Y_listboxFile, width =obj.W_scroll_Y_listboxFile, x=obj.posX_scroll_Y_listboxFile, y=obj.posY_scroll_Y_listboxFile)            
            
            #3b.scrollbar_X listbox File
            obj.H_scroll_X_listboxFile = 20
            obj.W_scroll_X_listboxFile = obj.W_listboxFile
            obj.posX_scroll_X_listboxFile = obj.pos_ABCD_listboxFile[3][0]#xB_Listbox
            obj.posY_scroll_X_listboxFile = obj.pos_ABCD_listboxFile[3][1]

            # Tọa độ scroll_listboxFile_X tính từ obj.posXY_liH_listboxFile
            obj.pos_ABCD_scroll_X_listboxFile = pos_Rectangular(obj.H_scroll_X_listboxFile, obj.W_scroll_X_listboxFile, obj.posX_scroll_X_listboxFile, obj.posY_scroll_X_listboxFile)
            obj.scroll_X_listboxFile.place(height = obj.H_scroll_X_listboxFile, width =obj.W_scroll_X_listboxFile, x=obj.posX_scroll_X_listboxFile, y=obj.posY_scroll_X_listboxFile)
            
            # 4.Button Xóa
            obj.H_bntXoa = 30
            obj.W_bntXoa = 70
            obj.posX_bntXoa = obj.pos_ABCD_scroll_Y_listboxFile[1][0]+obj.tab_X#xB_Listbox
            obj.posY_bntXoa = obj.pos_ABCD_scroll_Y_listboxFile[1][1]+3*obj.tab_Y

            # Tọa độ bntXoa_Y tính từ obj.posXY_liH_listboxFile
            obj.pos_ABCD_bntXoa = pos_Rectangular(obj.H_bntXoa, obj.W_bntXoa, obj.posX_bntXoa, obj.posY_bntXoa) 
            obj.bntXoa.place(height = obj.H_bntXoa, width =obj.W_bntXoa, x=obj.posX_bntXoa, y=obj.posY_bntXoa)  

            # 5.Button Thêm file
            obj.H_bntThem = obj.H_bntXoa
            obj.W_bntThem = obj.W_bntXoa
            obj.posX_bntThem = obj.pos_ABCD_bntXoa[3][0]
            obj.posY_bntThem = obj.pos_ABCD_bntXoa[3][1]+2*obj.tab_Y

            obj.pos_ABCD_bntThem = pos_Rectangular(obj.H_bntThem, obj.W_bntThem, obj.posX_bntThem, obj.posY_bntThem)
            obj.bntThem.place(height = obj.H_bntThem, width =obj.W_bntThem, x=obj.posX_bntThem, y=obj.posY_bntThem)

            # 6.Button Luu file
            obj.H_bntLuu = obj.H_bntXoa
            obj.W_bntLuu = obj.W_bntXoa
            obj.posX_bntLuu = obj.pos_ABCD_bntThem[3][0]
            obj.posY_bntLuu = obj.pos_ABCD_bntThem[3][1]+2*obj.tab_Y
            
            
            obj.pos_ABCD_bntLuu = pos_Rectangular(obj.H_bntLuu, obj.W_bntLuu, obj.posX_bntLuu, obj.posY_bntLuu)

            obj.bntLuu.place(height = obj.H_bntLuu, width =obj.W_bntLuu, x=obj.posX_bntLuu, y=obj.posY_bntLuu)
            
            # 7.Tiêu đề label Seach
            obj.H_lbTieudeSeach = 100
            obj.W_lbTieudeSeach = 200
            obj.posX_lbTieudeSeach = obj.posX_bntXoa + obj.W_bntXoa +obj.tab_X 
            obj.posY_lbTieudeSeach = obj.tab_Y
            obj.pos_ABCD_lbTieudeSeach = pos_Rectangular(obj.H_lbTieudeSeach, obj.W_lbTieudeSeach, obj.posX_lbTieudeSeach, obj.posY_lbTieudeSeach) 
            
            obj.lbTieudeSeach.place(height = obj.H_lbTieudeSeach, width =obj.W_lbTieudeSeach, x=obj.posX_lbTieudeSeach, y=obj.posY_lbTieudeSeach)

            # 7.Entry Seach
            obj.H_entrySeach = 30
            obj.W_entrySeach = 600
            obj.posX_entrySeach = obj.posX_lbTieudeSeach + obj.W_lbTieudeSeach + obj.tab_X
            obj.posY_entrySeach = (obj.pos_ABCD_lbTieudeSeach[2][1])/4+2*obj.tab_X#obj.posY_lbTieudeSeach+obj.H_lbTieudeSeach
            obj.pos_ABCD_entrySeach = pos_Rectangular(obj.H_entrySeach, obj.W_entrySeach, obj.posX_entrySeach, obj.posY_entrySeach) 
            
            obj.entry_Seach.place(height = obj.H_entrySeach, width = obj.W_entrySeach, x=obj.posX_entrySeach, y=obj.posY_entrySeach)
            obj.entry_Seach.bind("<KeyRelease>", obj.seach_string_trView) 
            #"<FocusIn>": con chuột; <Button-2>, <Double 1><Enter><Leave>
            
            # 8. trView
            obj.H_trView = obj.H_listboxFile
            obj.W_trView = 1360
            obj.posX_trView = obj.posX_bntXoa+obj.W_bntXoa+3*obj.tab_X
            obj.posY_trView = obj.posY_listboxFile
            obj.pos_ABCD_trView = pos_Rectangular(obj.H_trView, obj.W_trView, obj.posX_trView, obj.posY_trView) 
            
            obj.trView.place(height = obj.H_trView, width = obj.W_trView, x=obj.posX_trView, y=obj.posY_trView)
            
            #8a.scrollbar trView
            obj.H_scroll_Y_trView = obj.H_listboxFile
            obj.W_scroll_Y_trView = 20
            obj.posX_scroll_Y_trView = obj.pos_ABCD_trView[1][0]#xB_Listbox
            obj.posY_scroll_Y_trView = obj.pos_ABCD_trView[1][1]

            # Tọa độ scroll_TrView
            obj.pos_ABCD_scroll_Y_trView = pos_Rectangular(obj.H_scroll_Y_trView, obj.W_scroll_Y_trView, obj.posX_scroll_Y_trView, obj.posY_scroll_Y_trView)
            obj.scroll_Y_trView.place(height = obj.H_scroll_Y_trView, width =obj.W_scroll_Y_trView, x=obj.posX_scroll_Y_trView, y=obj.posY_scroll_Y_trView)            
            
            #8b.scrollbar_X trView
            obj.H_scroll_X_trView = 20
            obj.W_scroll_X_trView = obj.W_trView
            obj.posX_scroll_X_trView = obj.pos_ABCD_trView[3][0]#xB_Listbox
            obj.posY_scroll_X_trView = obj.pos_ABCD_trView[3][1]

            # Tọa độ scroll_listboxFile_X tính từ obj.posXY_liH_listboxFile
            obj.pos_ABCD_scroll_X_trView = pos_Rectangular(obj.H_scroll_X_trView, obj.W_scroll_X_trView, obj.posX_scroll_X_trView, obj.posY_scroll_X_trView)
            obj.scroll_X_trView.place(height = obj.H_scroll_X_trView, width =obj.W_scroll_X_trView, x=obj.posX_scroll_X_trView, y=obj.posY_scroll_X_trView)
            
            # obj.trView.bind("<<TreeviewSelect>>", obj.abc)
            
            #  # 7.Tiêu đề ListboxSheet
            # obj.H_lbTieudelistboxSheet = 25
            # obj.W_lbTieudelistboxSheet = objW*0.3
            # obj.posX_lbTieudelistboxSheet = obj.posX_bntXoa + obj.tab_X
            # obj.posY_lbTieudelistboxSheet = obj.posY_lbTieudelistbox
            # obj.pos_ABCD_lbTieudelistboxSheet = pos_Rectangular(obj.H_lbTieudelistboxSheet, obj.W_lbTieudelistboxSheet, obj.posX_lbTieudelistboxSheet, obj.posY_lbTieudelistboxSheet) 
            
            # obj.lbTieudelistboxSheet.place(height = obj.H_lbTieudelistboxSheet, width =obj.W_lbTieudelistboxSheet, x=obj.posX_lbTieudelistboxSheet, y=obj.posY_lbTieudelistboxSheet)

            # 7.ListboxSheet
            # obj.H_listboxSheets = obj.H_listboxFile
            # obj.W_listboxSheets = objW*0.3
            # obj.posX_listboxSheets = obj.pos_ABCD_bntXoa[1][0] + obj.tab_X
            # obj.posY_listboxSheets = obj.posY_listboxFile
            # obj.pos_ABCD_listboxSheets = pos_Rectangular(obj.H_listboxSheets, obj.W_listboxSheets, obj.posX_listboxSheets, obj.posY_listboxSheets) 
            
            # obj.listboxSheets.place(height = obj.H_listboxSheets, width =obj.W_listboxSheets, x=obj.posX_listboxSheets, y=obj.posY_listboxSheets)      
            
            # obj.listboxSheets.bind("<<ListboxSelect>>") #, obj.onSelect_listboxSheets
           
            # # 8a.Scrollbar Y_listboxSheet
            # obj.H_scroll_listboxSheets = obj.H_listboxFile
            # obj.W_scroll_listboxSheets = 20
            # obj.posX_scroll_listboxSheets = obj.pos_ABCD_listboxSheets[1][0]#xB_Listbox
            # obj.posY_scroll_listboxSheets = obj.pos_ABCD_listboxSheets[1][1]

            # # Tọa độ scroll_listboxSheets_Y tính từ obj.posXY_liH_listboxFile
            # obj.pos_ABCD_scroll_listboxSheets = pos_Rectangular(obj.H_scroll_listboxSheets, obj.W_scroll_listboxSheets, obj.posX_scroll_listboxSheets, obj.posY_scroll_listboxSheets) 
        
            # obj.scroll_listboxSheets.place(height = obj.H_scroll_listboxSheets, width =obj.W_scroll_listboxSheets, x=obj.posX_scroll_listboxSheets, y=obj.posY_scroll_listboxSheets)  
            
            # 9 Checkbox for listboxFile
            obj.H_chkbox_mutiSelect_listboxFile = 20
            obj.W_chkbox_mutiSelect_listboxFile = 135
            obj.posX_chkbox_mutiSelect_listboxFile = obj.pos_ABCD_scroll_X_listboxFile[3][0]
            obj.posY_chkbox_mutiSelect_listboxFile = obj.pos_ABCD_scroll_X_listboxFile[3][1]+obj.tab_Y            
            
            obj.pos_ABCD_chkbox_mutiSelect_listboxFile = pos_Rectangular(obj.H_chkbox_mutiSelect_listboxFile, obj.W_chkbox_mutiSelect_listboxFile, obj.posX_chkbox_mutiSelect_listboxFile, obj.posY_chkbox_mutiSelect_listboxFile)

            obj.chkbox_mutiSelect_listboxFile.place(height = obj.H_chkbox_mutiSelect_listboxFile, width =obj.W_chkbox_mutiSelect_listboxFile, x=obj.posX_chkbox_mutiSelect_listboxFile, y=obj.posY_chkbox_mutiSelect_listboxFile)
            
            
            # , text='Python',variable=var1, onvalue=1, offvalue=0, command=print_selection

            # # label kết quả (pause date 1.5.2021)
            # obj.label_ketqua.place(height = 25, width =220, x=10+obj.changeX, y=260+obj.changeY)
            
            # # combobox show sheets(pause date 1.5.2021)
            # obj.cbb_sheet.place(height = 25, width =100, x=230+obj.changeX, y=260+obj.changeY)
            # obj.cbb_sheet.bind("<<ComboboxSelected>>", obj.onSelectcomboSheet)

            # obj.label_wbs.place(height = 25, width =1000, x=0+obj.changeX, y=280+obj.changeY)

        def __init__(obj, master):
            super().__init__(master)#Frame().__init__(parent)#super().__init__(master)
            
            obj.master = master

            # 1.Khởi tạo listboxFile
            obj.lbTieudelistbox = Label(obj, text="Danh mục Playlits files!!!", fg = "blue",font=("Times New Roman", 15))              
            obj.listboxFile = Listbox(obj, bg="white", fg="green", font=("Times New Roman", 12))#, selectmode=MULTIPLE , exportselection=False, selectmode=MULTIPLE, exportselection=False giữ in event click, yscrollcommand = obj.scroll_listboxFile.set,
                        
            #Khởi tạo scrollBar cho listboxFile
            obj.scroll_Y_listboxFile = Scrollbar(obj, orient="vertical")#"orient": định hướng, "vertical": chiều dọc
            #Tạo phục thuộc listbox, thức hiện sau khi khởi tạo scollBar
            obj.scroll_Y_listboxFile.configure(command = obj.listboxFile.yview)     
            obj.listboxFile.configure(yscrollcommand = obj.scroll_Y_listboxFile.set)
                    
            
            # ----------------------
            #Khởi tạo X_scrollBar cho listboxFile
            obj.scroll_X_listboxFile = Scrollbar(obj, orient="horizontal")#"orient": định hướng, "horizontal": chiều ngang
            #Tạo phục thuộc listbox, thức hiện sau khi khởi tạo scollBar             
            obj.scroll_X_listboxFile.configure(command = obj.listboxFile.xview)
            obj.listboxFile.configure(xscrollcommand = obj.scroll_X_listboxFile.set) 
            # -------------------
            obj.load_data_from_excelPandas()

            # 2.Khởi tạo các nút button
            
            obj.bntXoa = version.Button(obj, text="Xóa file", command=obj.onClick_remove)#variable=obj.strlabel_ketqua,
            obj.bntThem = version.Button(obj, text="Thêm file", command=obj.addfile_in_listbox)
            obj.bntLuu = version.Button(obj, text="Lưu File", command=obj.write_from_listbox_to_txt)#, command=obj.addfile_in_listbox
            
            # 3.Khởi tạo label_Seach
            obj.lbTieudeSeach = Label(obj, text="Tìm kiếm:", fg = "blue",font=("Times New Roman", 25))              
            # obj.listboxSheets = Listbox(obj,bg="white",  fg="green", font=("Times New Roman", 20))#, selectmode=MULTIPLE , exportselection=False, selectmode=MULTIPLE, exportselection=False giữ in event click, yscrollcommand = obj.scroll_listboxSheets.set,

            # 3a.Entry Seach
            
            obj.entry_Seach = Entry(obj, fg = "green", font = ("Times New Roman", 15))#text="",  textvariable=obj.str_entry, validate="focusout"textvariable=obj.str_entry, 
                            
            # 3b.TrView
            obj.trView = version.Treeview(obj, selectmode = 'browse')# selectmode ='browse', show="headings",  columns="columns",
            obj.scroll_X_trView = Scrollbar(obj, orient = "horizontal", command = obj.trView.xview)#
            
            obj.trView.configure(xscrollcommand = obj.scroll_X_trView.set) 
            obj.scroll_X_trView.configure(command = obj.trView.xview)

            obj.scroll_Y_trView = Scrollbar(obj, orient = "vertical", command = obj.trView.yview)            
            obj.trView.configure(yscrollcommand = obj.scroll_Y_trView.set)
            obj.scroll_Y_trView.configure(command = obj.trView.yview)          
            
            # Khởi tạo Checkbox for listboxFile
            obj.var1 = IntVar()
            obj.chkbox_mutiSelect_listboxFile = Checkbutton(obj,text='Chọn nhiều files',
                            variable=obj.var1, onvalue=1, offvalue=0, 
                            command=obj.chk_mutiSelect, compound='left',
                            font=("Arial Bold", 10), fg="blue")#, command=chk_mutiSelect

            #label kết quả ( Tham Khảo # self.var = BooleanVar())
            # obj.strlabel_wbs = StringVar() #kết quả chọn listbox
            # obj.label_wbs = Label(obj, text=0, textvariable=obj.strlabel_wbs, fg="blue" ) #dấu kết quả wbs           
            
            # # label kết quả (pause date 1.5.2021)
            # obj.strlabel_ketqua = StringVar()
            # obj.label_ketqua = Label(obj, text=0, textvariable=obj.strlabel_ketqua, fg="blue" )
            
            # # ComboSheet
            # obj.value_comboSheet = StringVar()
            # obj.cbb_sheet = version.Combobox(obj, textvariable= obj.value_comboSheet)#, values="", valuevariable=onSelect_listboxFile)
            # # obj.cbb_sheet.current(0) # Lấy index hiện hành
        
            master.bind("<Configure>", obj.placeGui)
        
        # Hàm seach cho Entry_Seach
        def onValidate(obj, d, i, P, s, S, v, V, W):
            obj.text.delete("1.0", "end")
            obj.text.insert("end","OnValidate:\n")
            obj.text.insert("end","d='%s'\n" % d)
            obj.text.insert("end","i='%s'\n" % i)
            obj.text.insert("end","P='%s'\n" % P)
            obj.text.insert("end","s='%s'\n" % s)
            obj.text.insert("end","S='%s'\n" % S)
            obj.text.insert("end","v='%s'\n" % v)
            obj.text.insert("end","V='%s'\n" % V)
            obj.text.insert("end","W='%s'\n" % W)
            print(obj.text)
            # Disallow anything but lowercase letters
            if S == S.lower():
                return True
            else:
                obj.bell()
                return False

        # Khóa func này 30.4.2021
        def onSelectcomboSheet(obj, val):            
            sender = val.widget #gán sender: trả về vị trí x, y trong widget (ở đây là lisbox)
            obj.value_click_comboSheet = sender.get() # index tại vị trí click
            obj.value_comboSheet = obj.value_click_comboSheet      
            # obj.cbb_sheet.select() # Lấy index hiện hành

        # 1. Đọc từ file data.txt from: create_playlists_Excels()
        def load_data_from_excelPandas(obj):        
            # obj.list_App = create_playlists_Excels() 
            obj.list_App, obj.playlists, obj.playlist_ThisCom, obj.excels_ThisCom, obj.playlist_OtherCom, obj.excels_OtherCom = define_list_App()
        
            obj.update_listboxFile() #load data lên lisboxFile 
        
        def update_listboxFile(obj):           
            obj.listboxFile.delete(0,'end') #Xóa sạch listbox trước khi nạp tránh lỗi ko cần thiết            
            # obj.set_name_class_Excel() # Đặt tên lại cho dễ nhớ
            obj.show_All_data_List_App() # Show tiêu đề cho list_App

        # -----------Begin load file lên lisxbox-------------------------------        
        def show_All_data_List_App(obj):
            if obj.list_App != None:
                # show title List_App
                margin_listApp = ["--------------- <<", ">> ---------------"]
                # obj.playlists = obj.list_App.playlists # đặt lại tên dễ nhớ            
                obj.total_playlists = len(obj.playlists) #Tổng số playlist 
                # show title playlists:"----------Select all file <<On click>>----------"	 
                obj.listboxFile.insert(END, f"{margin_listApp[0]}{obj.list_App.title} {obj.total_playlists} playlist {margin_listApp[1]}")   
                obj.row_title_playlists = obj.listboxFile.index(END)-1 # Gán dòng tiêu đề list_App            
                
                # Show title playlists      
                if obj.total_playlists>0:                
                    for i_pl in range(obj.total_playlists):                     
                        obj.show_title_playlists(i_pl)

        def show_title_playlists(obj, i_pl):
            
            # if obj.playlists[i_pl].click==True:
            margin_playlist = ["I./ ", "II./ "]
            if obj.playlists[i_pl]== None:
                if i_pl==0:
                    # name_Computer=nameComputer()
                    obj.listboxFile.insert(END, f"{margin_playlist[i_pl]} File on This Computer ({nameComputer()}): 0 file") 
                else:
                    # name_Computer=nameComputer()
                    obj.listboxFile.insert(END, f"{margin_playlist[i_pl]} File on Other Computer ({nameComputer()}): 0 file")
            else:
                obj.total_excels = len(obj.playlists[i_pl].excels)
                obj.listboxFile.insert(END, f"{margin_playlist[i_pl]}{obj.playlists[i_pl].title} {obj.total_excels} file") 
                # show title playlist ThisCom & OtherCom
                
                # 3. Show file excels
                if obj.total_excels>0:
                    for j_excel in range(obj.total_excels):#playlists_Excel.playlists[i].excels:
                        excel = obj.playlists[i_pl].excels[j_excel] # đặt tên lại 
                        # excel.click = True
                        obj.show_file_excels(excel, j_excel)

        def show_file_excels(obj, excel, j_excel):        
            # if excel.click ==True:
                margin_empty_line = "   " # 3 khoảng trống
                margin_middle = margin_empty_line + "├──"
                margin_conti = margin_empty_line +"|   "
                margin_last =margin_empty_line + "└──"            
                chks_endFile=False
                if j_excel != obj.total_excels-1: # truong hop file ko phải dòng cuối"├──"                
                    obj.listboxFile.insert(END, f"{margin_middle} [File{j_excel+1}]: //{excel.name_Exten}")
                    # excel = obj.playlists[i_pl].excels[j_excel]
                    
                    # Show sheet, rows, cols                   
                    total_sh = len(excel.sheets)            
                    for i_sh in range(total_sh):
                        sheet = excel.sheets[i_sh]  
                        obj.show_sheet(sheet, i_sh, total_sh, chks_endFile)
                    
                else:# truong hop file dòng cuối "└──"   
                    obj.listboxFile.insert(END, f"{margin_last} [File{j_excel+1}]: //{excel.name_Exten}")
                    
                    # Show sheets, rows, cols
                    chks_endFile = True                    
                    total_sh = len(excel.sheets)            
                    for i_sh in range(total_sh): 
                        sheet =  excel.sheets[i_sh]
                        # Show sheet
                        obj.show_sheet(sheet, i_sh, total_sh, chks_endFile)
        # -----------End load file lên lisxbox-------------------------------
       
        def show_sheet(obj, sheet, i_sh, total_sh, chks_endFile):
            if sheet!= None : #and sheet.click == True      and sheet.click == True           
                margin_empty_line = "   " # 3 khoảng trống
                margin_middle = 2* margin_empty_line + "├──"
                margin_conti_file = margin_empty_line +" |   "
                margin_conti_Endfile = margin_empty_line +"     "
                margin_last = 2*margin_empty_line + "└──"  
                chks_endSheet = False
                if chks_endFile !=True: # ko phải file cuối
                    if i_sh != total_sh-1:
                        obj.listboxFile.insert(END, f"{margin_conti_file}{margin_middle} [Sheet{i_sh+1}]: {sheet.name}!")                
                        obj.show_data(sheet, i_sh, total_sh, chks_endFile, chks_endSheet)  #Show Data, rows, columns               
                            
                    else: #file giữa, sheet cuối
                        obj.listboxFile.insert(END, f"{margin_conti_file}{margin_last} [Sheet{i_sh+1}]: {sheet.name}!")           
                        chks_endSheet = True
                        obj.show_data(sheet, i_sh, total_sh, chks_endFile, chks_endSheet) #Show Data, rows, columns           
                else: #File cuối
                    if i_sh != total_sh-1:
                        obj.listboxFile.insert(END, f"{margin_conti_Endfile}{margin_middle} [Sheet{i_sh+1}]: {sheet.name}!")                
                        obj.show_data(sheet, i_sh, total_sh, chks_endFile, chks_endSheet)  #Show Data, rows, columns               
                            
                    else:# File cuối, sheet cuối
                        obj.listboxFile.insert(END, f"{margin_conti_Endfile}{margin_last} [Sheet{i_sh+1}]: {sheet.name}!")           
                        chks_endSheet = True
                        obj.show_data(sheet, i_sh, total_sh, chks_endFile, chks_endSheet) #Show Data, rows, columns    

        def show_data(obj, sheet, i_sh, total_sh, chks_endFile, chks_endSheet):
            if sheet.click ==True:
                margin_empty_line = "   " # 3 khoảng trống
                margin_middle = 2* margin_empty_line + "├──"
                margin_conti_file = margin_empty_line +" |   "
                margin_conti_sheet = 2*margin_empty_line +" |   "
                margin_conti_Endsheet = 2*margin_empty_line +"     "
                margin_conti_Endfile = margin_empty_line +"     "
                margin_last = 2*margin_empty_line + "└──"  
                if chks_endFile!=True and chks_endSheet!=True: #sheet và file ko phải là dòng cuối
                    obj.listboxFile.insert(END, f"{margin_conti_file}{margin_conti_sheet}{margin_middle} [Data]: {sheet.data}")
                    obj.listboxFile.insert(END, f"{margin_conti_file}{margin_conti_sheet}{margin_middle} [Rows]: {sheet.num_Rows} (row)")
                    obj.listboxFile.insert(END, f"{margin_conti_file}{margin_conti_sheet}{margin_last} [Cols]: {sheet.num_Cols} (column)")
                elif chks_endFile !=True and chks_endSheet==True:#file <> cuối, sheet cuối
                    obj.listboxFile.insert(END, f"{margin_conti_file}{margin_conti_Endsheet}{margin_middle} [Data]: {sheet.data}")
                    obj.listboxFile.insert(END, f"{margin_conti_file}{margin_conti_Endsheet}{margin_middle} [Rows]: {sheet.num_Rows} (row)")
                    obj.listboxFile.insert(END, f"{margin_conti_file}{margin_conti_Endsheet}{margin_last} [Cols]: {sheet.num_Cols} (column)")
                elif chks_endFile ==True and chks_endSheet!=True:#file cuối, sheet <> cuối
                    obj.listboxFile.insert(END, f"{margin_conti_Endfile}{margin_conti_sheet}{margin_middle} [Data]: {sheet.data}")
                    obj.listboxFile.insert(END, f"{margin_conti_Endfile}{margin_conti_sheet}{margin_middle} [Rows]: {sheet.num_Rows} (row)")
                    obj.listboxFile.insert(END, f"{margin_conti_Endfile}{margin_conti_sheet}{margin_last} [Cols]: {sheet.num_Cols} (column)")    
                elif chks_endFile ==True and chks_endSheet==True:#file cuối, sheet cuối
                    obj.listboxFile.insert(END, f"{margin_conti_Endfile}{margin_conti_Endsheet}{margin_middle} [Data]: {sheet.data}")
                    obj.listboxFile.insert(END, f"{margin_conti_Endfile}{margin_conti_Endsheet}{margin_middle} [Rows]: {sheet.num_Rows} (row)")
                    obj.listboxFile.insert(END, f"{margin_conti_Endfile}{margin_conti_Endsheet}{margin_last} [Cols]: {sheet.num_Cols} (column)") 
                   
        def read_data_on_lixtboxFile(obj):
            obj.row_title_playlists = -1
            obj.row_title_ThisCom = -1
            obj.row_title_OtherCom = -1
            obj.file_on_listbox=[]

            total_id = obj.listboxFile.index(END)
            # value = obj.listboxFile.get()
            for i_lbox in range(total_id):
                value = obj.listboxFile.get(i_lbox)
                if value.find("---------------")!=-1:
                    obj.row_title_playlists = i_lbox
                elif value.find("File on This Computer")!=-1:
                    obj.row_title_ThisCom = i_lbox
                elif value.find("File on Other Computer")!=-1:
                    obj.row_title_OtherCom = i_lbox                
                
                elif value.find("[File")!=-1:                    
                    str = Find_Str_mid_two_Char(value, "[", "]")
                    id_file = int(str[4:len(str)])-1
                    
                    # Tồn tại thisCom, OtherCom
                    if obj.row_title_ThisCom!=-1 and obj.row_title_OtherCom!=-1:
                        if i_lbox<obj.row_title_OtherCom:#file trong thisCom
                            obj.file_on_listbox.append(obj.playlists[0].excels[id_file].link)
                        elif i_lbox>obj.row_title_OtherCom:#file trong OtherCom: 
                            obj.file_on_listbox.append(obj.playlists[1].excels[id_file].link)
                    # ko tồn tại thisCom
                    elif obj.row_title_ThisCom==-1 and obj.row_title_OtherCom!=-1:
                        obj.file_on_listbox.append(obj.playlists[1].excels[id_file].link)
                    elif obj.row_title_ThisCom!=-1 and obj.row_title_OtherCom==-1:
                        obj.file_on_listbox.append(obj.playlists[0].excels[id_file].link)
             
            print (f"đây là file listbox tổng file là: {len(obj.file_on_listbox)}, gồm: {obj.file_on_listbox}")

            return obj.row_title_playlists, obj.row_title_ThisCom, obj.row_title_OtherCom, obj.file_on_listbox 
        
        def onSelect_listboxFile(obj, event_click): #trả về giá trị tại vị trí click                      
            # Nhận id từ listboxFile
            obj.row_title_playlists, obj.row_title_ThisCom, obj.row_title_OtherCom, obj.file_on_listbox = obj.read_data_on_lixtboxFile()
                        
            # Lấy chỉ số dòng người dùng đã click trên lisboxFile
            obj.idx_ClickOnListboxFile = obj.listboxFile.curselection()

            # Kiểm tra các điều kiện người dùng đã click vào (title playlist hay file, sheets....)
            obj.chk_Used_has_click_On_listboxFile()            

        def chk_Used_has_click_On_listboxFile(obj): # làm sau chk_range_conditions_On_listboxFile
            obj.total_click_listboxFile = len(obj.idx_ClickOnListboxFile)
            if obj.total_click_listboxFile >0:
                for i_click in range(obj.total_click_listboxFile):                    
                    # show kết quả sau click
                    obj.return_result_after_Click(i_click)

        def return_result_after_Click(obj, i_click):
            # Trường hợp click vào tiêu đề title_listApp
            if obj.idx_ClickOnListboxFile[i_click] == obj.row_title_playlists: #khi người dùng chọn vào tiêu đề playlists                        
                obj.showOrhide_playlist()
                obj.update_listboxFile()
                obj.listboxFile.select_set(obj.row_title_playlists, END) # chọn toàn bộ  rows trên listboxFile                        
                # obj.listboxFile.focus_set()

            # Trường hợp click vào tiêu đề ThisCom         
            elif obj.idx_ClickOnListboxFile[i_click]==obj.row_title_ThisCom:
                #gán dòng tiêu đề cho ThisCom
                if obj.playlist_ThisCom!=None:
                    if obj.playlist_ThisCom.click==True:# for excel_ThisCom in obj.playlist_ThisCom:  
                        obj.playlist_ThisCom.click=False
                        obj.return_default_IsFalse_files()
                    else: 
                        obj.playlist_ThisCom.click=True
                        obj.return_IsTrue_files()
                        # obj.change_Atribute_click_for_all_file_ThisCom(obj.playlist_ThisCom) # Thay đổi False ==True
                    obj.update_listboxFile()
                if obj.row_title_OtherCom!=-1:
                    obj.listboxFile.select_set(obj.row_title_ThisCom, obj.row_title_OtherCom-1)
                    # obj.listboxFile.focus_set()
                else:
                    obj.listboxFile.select_set(obj.row_title_ThisCom, END)
                    
                # #  Trả về giá trị các dòng đã chọn sau khi click tiêu đề trên listboxFile
                # for i_afterClick in range(obj.row_title_ThisCom+1, obj.row_title_OtherCom): 
                #     # obj.id_afterClick_lisboxFile =[obj.row_title_ThisCom+1, obj.row_title_OtherCom-1]
                #     print("obj.id_afterClick_lisboxFile", i_afterClick)            
            
            # Trường hợp click vào tiêu đề OtherCom 
            elif obj.listboxFile.get(obj.idx_ClickOnListboxFile[i_click]).find("File on Other Computer")!=-1:
                #gán dòng tiêu đề cho ThisCom
                obj.row_title_OtherCom = obj.idx_ClickOnListboxFile[i_click]                      
                if obj.playlist_OtherCom!=None:
                    if obj.playlist_OtherCom.click==True:# for excel_ThisCom in obj.playlist_OtherCom:                            
                        obj.playlist_OtherCom.click=False
                        obj.return_default_IsFalse_files()
                    else: 
                        obj.playlist_OtherCom.click=True
                        obj.return_IsTrue_files()
                    # obj.change_Atribute_click_for_all_file_ThisCom(obj.playlist_ThisCom) # Thay đổi False ==True
                obj.update_listboxFile()                
                obj.listboxFile.select_set(obj.row_title_OtherCom, END)
                # obj.listboxFile.focus_set()
            #click ngoài các tiêu đề App, ThisCom, OtherCom
            else: 
                if obj.idx_ClickOnListboxFile[i_click] < obj.row_title_OtherCom: # phạm vi ThisCom
                    id_click=obj.idx_ClickOnListboxFile[i_click]                   
                    obj.add_otheritem_from_Thiscomputer(id_click)
                    # print("Click in range ThisCom", obj.idx_ClickOnListboxFile[i_click])
                else: # phạm vi OtherCom
                    obj.find_file_After_Clicked(obj.idx_ClickOnListboxFile[i_click])
                    # print("Click in range OtherCom", obj.idx_ClickOnListboxFile[i_click])

        def showOrhide_playlist(obj):
            for playlist in obj.playlists:
                if playlist!= None and playlist.click==True:
                    playlist.click=False           

        def add_otheritem_from_Thiscomputer(obj, id_click):
            In_str = obj.listboxFile.get(id_click)
            proces_Str = obj.find_file_sheet_data_After_Clicked(In_str, id_click)
            
        def find_file_sheet_data_After_Clicked(obj, In_str, id_click):
            sheetname =""
            # Find_Str_mid_two_Char(into_Str, begin_chr, end_chr)
            str = Find_Str_mid_two_Char(In_str, "[", "]")
            if str!="":                
                if str == "Data":
                    id_sheet = id_click-1
                    obj.id_sheet_classExell = obj.find_sheet_After_Clicked(id_sheet)                    
                    obj.id_file_classExell = obj.find_file_After_Clicked(id_sheet-1)#-1: Thực hiện đếm lùi trên listbox để tìm dòng file
                    
                    """ ********* Đoạn này sau này xuất data qua treeView  *************************"""
                    
                    obj.show_data_trView()# print("fsdfsdF")                
                
                # sheetname[5:len(sheetname)]
                elif str[0:5] == "Sheet":                    
                    obj.id_sheet_classExell = obj.find_sheet_After_Clicked(id_click)                    
                    obj.id_file_classExell = obj.find_file_After_Clicked(id_click-1)#-1: Thực hiện đếm lùi trên listbox để tìm dòng file
                    obj.change_Atribute_click_for_sheet_ThisCom()#obj.id_file_classExell, obj.id_sheet_classExell
                    obj.update_listboxFile()
                    if obj.excels_ThisCom[obj.id_file_classExell].sheets[obj.id_sheet_classExell].click == True:
                        obj.listboxFile.select_set(id_click, id_click+3)
                        # obj.listboxFile.focus_set()
                    else:
                        obj.listboxFile.select_set(id_click)
                        # obj.listboxFile.focus_set() 
                elif str[0:4] == "File":
                    obj.id_sheet_classExell = -1                  
                    obj.id_file_classExell = obj.find_file_After_Clicked(id_click)#-1: Thực hiện đếm lùi trên listbox để tìm dòng file
                    id_file = int(str[4:len(str)])-1  #'File1: lấy số 1
                    obj.change_Atribute_click_for_File()
                    obj.update_listboxFile()
                    Rows_select = int(id_click)
                    if obj.excels_ThisCom[id_file].click == True:
                        for sh in obj.excels_ThisCom[id_file].sheets:
                            if sh.click==True:
                                Rows_select = Rows_select+4
                            else: 
                                Rows_select = Rows_select+1
                        obj.listboxFile.select_set(id_click, Rows_select)
                        # obj.listboxFile.focus_set()
                    else:
                        obj.listboxFile.select_set(id_click)
            print (f"file: {obj.id_file_classExell+1}, id_file_classExell: {obj.id_file_classExell}; sheetname: {obj.id_sheet_classExell+1}, obj.id_sheet_classExell: {obj.id_sheet_classExell}")
            return obj.id_file_classExell, obj.id_sheet_classExell

        def find_sheet_After_Clicked(obj, id_sheet):
            In_str_sheet = obj.listboxFile.get(id_sheet)
            sheetname = Find_Str_mid_two_Char(In_str_sheet, "[", "]")
            obj.id_sheet_classExell= int(sheetname[5:len(sheetname)])-1
            return obj.id_sheet_classExell

        def find_file_After_Clicked(obj, id_begin):
            fi_name =""                     
            for id_file in range(id_begin, 0, -1):#-1: Thực hiện đếm lùi trên listbox để tìm dòng file
                In_str = obj.listboxFile.get(id_file)
                fi_name = Find_Str_mid_two_Char(In_str, "[", "]")
                len_fi_name = len(fi_name)
                if len_fi_name >= 5 and fi_name[0:4] == "File": # len("File")
                    id_file = int(fi_name[4:len(fi_name)])-1  #'File1: lấy số 1
                    break    
            print (f"fi_name: {fi_name}, id_file: {id_file}")            
            return id_file

        def return_default_IsFalse_files(obj):
            for file in obj.playlist_ThisCom.excels:
                file.click=False

        def return_IsTrue_files(obj):
            if obj.playlist_ThisCom != None:
                for file in obj.playlist_ThisCom.excels:
                    file.click=True

        def change_Atribute_click_for_all_sheets_ThisCom(obj, excel):
            # Show or hide file
            for sh in excel.sheets:
                if sh.click==True:
                    sh.click=False            
                else: 
                    sh.click=True
            
        def change_Atribute_click_for_File(obj):
            
            if obj.excels_ThisCom[obj.id_file_classExell].click == True:
                obj.excels_ThisCom[obj.id_file_classExell].click = False
            else: obj.excels_ThisCom[obj.id_file_classExell].click = True    

        def change_Atribute_click_for_sheet_ThisCom(obj):
            if obj.excels_ThisCom[obj.id_file_classExell].sheets[obj.id_sheet_classExell].click == True:
                obj.excels_ThisCom[obj.id_file_classExell].sheets[obj.id_sheet_classExell].click = False
            else: obj.excels_ThisCom[obj.id_file_classExell].sheets[obj.id_sheet_classExell].click = True    
       
        def chk_mutiSelect(obj): 
            if obj.var1.get() == 1:
                obj.chkbox_mutiSelect_listboxFile.config(text='Đã chọn nhiều file')
                obj.listboxFile.config(selectmode=MULTIPLE)
            elif obj.var1.get() == 0:
                obj.chkbox_mutiSelect_listboxFile.config(text='Chọn nhiều files')
                obj.listboxFile.config(selectmode=EXTENDED) #selectmode=False
        
        def get_file_after_click_lixtboxFile(obj):
            obj.file_select=[]
            for i_select in obj.listboxFile.curselection():
                value = obj.listboxFile.get(i_select)
                if value.find("[File")!=-1:
                    if obj.row_title_ThisCom!=-1 and obj.row_title_OtherCom!=-1:
                        str = Find_Str_mid_two_Char(value, "[", "]")
                        id_file = int(str[4:len(str)])-1
                        if i_select<obj.row_title_OtherCom:#file trong thisCom
                            obj.file_select.append(obj.playlists[0].excels[id_file].link)
                        elif i_select>obj.row_title_OtherCom:#file trong OtherCom: 
                            obj.file_select.append(obj.playlists[1].excels[id_file].link)
                        # ko tồn tại thisCom
                        elif obj.row_title_ThisCom==-1 and obj.row_title_OtherCom!=-1:
                            obj.file_on_listbox.append(obj.playlists[1].excels[id_file].link)
                        elif obj.row_title_ThisCom!=-1 and obj.row_title_OtherCom==-1:
                            obj.file_on_listbox.append(obj.playlists[0].excels[id_file].link)
            return obj.file_select
                             
        def onClick_remove(obj):
            obj.get_file_after_click_lixtboxFile()            
            # obj.file_on_listbox = obj.file_on_listbox
            total_f = len(obj.file_on_listbox)
            i_f = 0
            
            # for i_f in range(total_f-1,-1,-1):
            while i_f <total_f:                
                for f_sl in obj.file_select:
                    if obj.file_on_listbox[i_f] == f_sl:
                        del obj.file_on_listbox[i_f]
                        total_f = total_f-1
                        i_f = i_f-1 # lùi lại vị trí bị xóa 
                        break
                i_f = i_f+1
            obj.write_from_listbox_to_txt()
            obj.load_data_from_excelPandas()                          
            # print (len(obj.file_on_listbox))
                   
        def addfile_in_listbox(obj):            
            # id_Endrows = obj.endRows_listbox()
            newFile = askopenfilename(filetypes = (("Excel","*.xl*"),("all files","*.*")))#askopenfilename() # open seach file
            obj.read_data_on_lixtboxFile()
            #filedialog.askopenfilename(filetypes = (("Text files","*.txt"),("all files","*.*")))

            #Add file vào
            obj.file_on_listbox.append(newFile)
            obj.write_from_listbox_to_txt()
            obj.load_data_from_excelPandas() 
                      
        def write_from_listbox_to_txt(obj):#file = os.getcwd() + "\Package_VuNghiXuan_Excel\data.txt" 
            
            obj.file_txt= os.getcwd() + "\Package_VuNghiXuan_Excel\data.txt" 
            write_files_from_lists_to_txt(obj.file_txt, obj.file_on_listbox)
            # write_files_wb_to_txt(obj.file_txt, obj.wbs)            

        # Viết cho trView
        def show_data_trView(obj):
            obj.clear_data_trView()
            
            id_f = obj.id_file_classExell
            id_sh = obj.id_sheet_classExell
            df = obj.excels_ThisCom[id_f].sheets[id_sh].data # gán lại cho giống Pandas
            
            obj.df_cols = list(df.columns)
            cols = obj.df_cols
            obj.trView["column"] = cols
            obj.inser_Cols_into_trView(cols)

            obj.df_rows = df.to_numpy().tolist()
            rows = obj.df_rows
            obj.inser_Rows_into_trView(rows)
        
        def inser_Cols_into_trView(obj, cols):    
            obj.trView["show"]="headings"
            for column in cols:#obj.trView["columns"]
                obj.trView.heading(column, text=column)
                # pass
        def inser_Rows_into_trView(obj, rows):            
            for row in rows:
                obj.inser_Row_into_trView(row)

        def inser_Row_into_trView(obj, row):
            obj.trView.insert("", "end", values=row)

            # pass
        def clear_data_trView(obj):
            obj.trView.delete(*obj.trView.get_children())

        def seach_string_trView(obj, e):            
            in_str = obj.entry_Seach.get()            
            if in_str != "":
                obj.clear_data_trView()

                for row in obj.df_rows:
                    for cell in row:
                        in_str = str(in_str)
                        cell = str(cell)
                        if str(in_str.lower()) in str(cell.lower()):
                                obj.inser_Row_into_trView(row)
                                break

                        # # if not in_str.isdigit():
                        # if type(in_str)==type(cell) and not cell.isdigit():
                        #     if str(in_str.lower()) in str(cell.lower()):
                        #         obj.inser_Row_into_trView(row)
                        #         break
                        # elif type(in_str)==type(cell) and cell.isdigit():#and in_str.isdigit()
                        #     if str(in_str) in str(cell):                            
                        #         obj.inser_Row_into_trView(row)
                        #         break
            else: 
                obj.show_data_trView()        
                   
            pass

    window = Tk()
    # path= os.getcwd()
    # vunghixuan=Image('photo', file=path+"\\10 xanh.png")
    # window.wm_iconphoto(True, vunghixuan)
    window.geometry("500x300+700+500") #"700+500": hiển thị vị trí khi show ra creenWindow
    showApp = formExcel(window)
    showApp.place(relwidth=1, relheight=1)
    window.mainloop()


def main():
    show_form()
if __name__=='__main__':   #__name__=='__main__': khi nào là hàm main() nằm trong bảng code náy mới dc chạy
    main()