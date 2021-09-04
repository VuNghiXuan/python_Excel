from tkinter import *
import tkinter.ttk as doimau
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
        
        def placeGui(obj, e):
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
            obj.W_lbTieudelistbox = objW*0.3
            obj.posX_lbTieudelistbox = 10
            obj.posY_lbTieudelistbox = 10
            obj.pos_ABCD_lbTieudelistbox = pos_Rectangular(obj.H_lbTieudelistbox, obj.W_lbTieudelistbox, obj.posX_lbTieudelistbox, obj.posY_lbTieudelistbox) 
            
            obj.lbTieudelistbox.place(height = obj.H_lbTieudelistbox, width =obj.W_lbTieudelistbox, x=obj.posX_lbTieudelistbox, y=obj.posY_lbTieudelistbox)
            
            # 2.Listbox+File        
            obj.H_listboxFile = 200
            obj.W_listboxFile = objW*0.3
            obj.posX_listboxFile = obj.posX_lbTieudelistbox
            obj.posY_listboxFile = obj.pos_ABCD_lbTieudelistbox[3][1]+obj.tab_Y # lấy yD của lbTieudelistbox + obj.tab_Y =10
            obj.pos_ABCD_listboxFile = pos_Rectangular(obj.H_listboxFile, obj.W_listboxFile, obj.posX_listboxFile, obj.posY_listboxFile) 

            
            obj.listboxFile.place(height = obj.H_listboxFile, width =obj.W_listboxFile, x=obj.posX_listboxFile, y=obj.posY_listboxFile)
            obj.listboxFile.bind("<<ListboxSelect>>", obj.onSelect_listboxFile)
           
            #3.scrollbar listbox
            obj.H_scroll_listboxFile = obj.H_listboxFile
            obj.W_scroll_listboxFile = 20
            obj.posX_scroll_listboxFile = obj.pos_ABCD_listboxFile[1][0]#xB_Listbox
            obj.posY_scroll_listboxFile = obj.pos_ABCD_listboxFile[1][1]

            # Tọa độ scroll_listboxFile_Y tính từ obj.posXY_liH_listboxFile
            obj.pos_ABCD_scroll_listboxFile = pos_Rectangular(obj.H_scroll_listboxFile, obj.W_scroll_listboxFile, obj.posX_scroll_listboxFile, obj.posY_scroll_listboxFile) 

            obj.scroll_listboxFile.place(height = obj.H_scroll_listboxFile, width =obj.W_scroll_listboxFile, x=obj.posX_scroll_listboxFile, y=obj.posY_scroll_listboxFile)            
            
            # 4.Button Xóa
            obj.H_bntXoa = 30
            obj.W_bntXoa = 70
            obj.posX_bntXoa = obj.pos_ABCD_scroll_listboxFile[1][0]+obj.tab_X#xB_Listbox
            obj.posY_bntXoa = obj.pos_ABCD_scroll_listboxFile[1][1]+3*obj.tab_Y

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
            
            # 7.Tiêu đề ListboxSheet
            obj.H_lbTieudelistboxSheet = 25
            obj.W_lbTieudelistboxSheet = objW*0.3
            obj.posX_lbTieudelistboxSheet = obj.posX_bntXoa + obj.tab_X
            obj.posY_lbTieudelistboxSheet = obj.posY_lbTieudelistbox
            obj.pos_ABCD_lbTieudelistboxSheet = pos_Rectangular(obj.H_lbTieudelistboxSheet, obj.W_lbTieudelistboxSheet, obj.posX_lbTieudelistboxSheet, obj.posY_lbTieudelistboxSheet) 
            
            obj.lbTieudelistboxSheet.place(height = obj.H_lbTieudelistboxSheet, width =obj.W_lbTieudelistboxSheet, x=obj.posX_lbTieudelistboxSheet, y=obj.posY_lbTieudelistboxSheet)

            # 7.ListboxSheet
            obj.H_listboxSheets = obj.H_listboxFile
            obj.W_listboxSheets = objW*0.3
            obj.posX_listboxSheets = obj.pos_ABCD_bntXoa[1][0] + obj.tab_X
            obj.posY_listboxSheets = obj.posY_listboxFile
            obj.pos_ABCD_listboxSheets = pos_Rectangular(obj.H_listboxSheets, obj.W_listboxSheets, obj.posX_listboxSheets, obj.posY_listboxSheets) 
            
            obj.listboxSheets.place(height = obj.H_listboxSheets, width =obj.W_listboxSheets, x=obj.posX_listboxSheets, y=obj.posY_listboxSheets)      
            
            obj.listboxSheets.bind("<<ListboxSelect>>") #, obj.onSelect_listboxSheets
           
            # 8.Scrollbar listboxSheet
            obj.H_scroll_listboxSheets = obj.H_listboxFile
            obj.W_scroll_listboxSheets = 20
            obj.posX_scroll_listboxSheets = obj.pos_ABCD_listboxSheets[1][0]#xB_Listbox
            obj.posY_scroll_listboxSheets = obj.pos_ABCD_listboxSheets[1][1]

            # Tọa độ scroll_listboxSheets_Y tính từ obj.posXY_liH_listboxFile
            obj.pos_ABCD_scroll_listboxSheets = pos_Rectangular(obj.H_scroll_listboxSheets, obj.W_scroll_listboxSheets, obj.posX_scroll_listboxSheets, obj.posY_scroll_listboxSheets) 

            obj.scroll_listboxSheets.place(height = obj.H_scroll_listboxSheets, width =obj.W_scroll_listboxSheets, x=obj.posX_scroll_listboxSheets, y=obj.posY_scroll_listboxSheets)  
            
            # 9 Checkbox for listboxFile
            obj.H_chkbox_mutiSelect_listboxFile = 20
            obj.W_chkbox_mutiSelect_listboxFile = 135
            obj.posX_chkbox_mutiSelect_listboxFile = obj.pos_ABCD_listboxFile[3][0]
            obj.posY_chkbox_mutiSelect_listboxFile = obj.pos_ABCD_listboxFile[3][1]+2*obj.tab_Y            
            
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
            obj.lbTieudelistbox = Label(obj, text="Danh mục các file đã lưu!!!", fg = "blue",font=("Times New Roman", 15))              
            obj.listboxFile = Listbox(obj, bg="white", fg="green", font=("Times New Roman", 10))#, selectmode=MULTIPLE , exportselection=False, selectmode=MULTIPLE, exportselection=False giữ in event click, yscrollcommand = obj.scroll_listboxFile.set,
                        
            #Khởi tạo scrollBar cho listboxFile
            obj.scroll_listboxFile = Scrollbar(obj, orient="vertical")#"orient": định hướng, "vertical": chiều dọc
            #Tạo phục thuộc listbox, thức hiện sau khi khởi tạo scollBar 
            obj.listboxFile.configure(yscrollcommand = obj.scroll_listboxFile.set)
            obj.scroll_listboxFile.configure(command = obj.listboxFile.yview)            
            obj.read_File_From_txt_saved()

            # 2.Khởi tạo các nút button
            
            obj.bntXoa = doimau.Button(obj, text="Xóa file", command=obj.onClick_remove)#variable=obj.strlabel_ketqua,
            obj.bntThem = doimau.Button(obj, text="Thêm file", command=obj.addfile_in_listbox)
            obj.bntLuu = doimau.Button(obj, text="Lưu File", command=obj.write_from_listbox_to_txt)#, command=obj.addfile_in_listbox
            
            # 3.Khởi tạo listboxSheet
            obj.lbTieudelistboxSheet = Label(obj, text="Sheets File đã chọn!!!", fg = "blue",font=("Times New Roman", 15))              
            obj.listboxSheets = Listbox(obj,bg="white",  fg="green", font=("Times New Roman", 10))#, selectmode=MULTIPLE , exportselection=False, selectmode=MULTIPLE, exportselection=False giữ in event click, yscrollcommand = obj.scroll_listboxSheets.set,
                        
            #Khởi tạo scrollBar cho listboxSheets
            obj.scroll_listboxSheets = Scrollbar(obj, orient="vertical")#"orient": định hướng, "vertical": chiều dọc
            #Tạo phục thuộc listbox, thức hiện sau khi khởi tạo scollBar 
            obj.listboxSheets.configure(yscrollcommand = obj.scroll_listboxSheets.set)
            obj.scroll_listboxSheets.configure(command = obj.listboxSheets.yview)            
            # obj.read_File_From_txt_saved()



            # Khởi tạo Checkbox for listboxFile
            obj.var1 = IntVar()
            obj.chkbox_mutiSelect_listboxFile =Checkbutton(obj,text='Chọn nhiều files',variable=obj.var1, onvalue=1, offvalue=0, command=obj.chk_mutiSelect, compound='left',font=("Arial Bold", 10), fg="blue")#, command=chk_mutiSelect

            #label kết quả ( Tham Khảo # self.var = BooleanVar())
            # obj.strlabel_wbs = StringVar() #kết quả chọn listbox
            # obj.label_wbs = Label(obj, text=0, textvariable=obj.strlabel_wbs, fg="blue" ) #dấu kết quả wbs           
            
            # # label kết quả (pause date 1.5.2021)
            # obj.strlabel_ketqua = StringVar()
            # obj.label_ketqua = Label(obj, text=0, textvariable=obj.strlabel_ketqua, fg="blue" )
            
            # # ComboSheet
            # obj.value_comboSheet = StringVar()
            # obj.cbb_sheet = doimau.Combobox(obj, textvariable= obj.value_comboSheet)#, values="", valuevariable=onSelect_listboxFile)
            # # obj.cbb_sheet.current(0) # Lấy index hiện hành
                  

            master.bind("<Configure>", obj.placeGui)

        # Khóa func này 30.4.2021
        def onSelectcomboSheet(obj, val):            
            sender = val.widget #gán sender: trả về vị trí x, y trong widget (ở đây là lisbox)
            obj.value_click_comboSheet = sender.get() # index tại vị trí click
            obj.value_comboSheet = obj.value_click_comboSheet      
            # obj.cbb_sheet.select() # Lấy index hiện hành

        

        
        # 1. Đọc từ file data.txt
        def read_File_From_txt_saved(obj):        
            obj.file_txt = os.getcwd() + "\Package_VuNghiXuan_Excel\data.txt"            
            obj.wbs_ThisComputer,  obj.wbs_OtherComputer = read_file_from_txt(obj.file_txt) #tao listbox  
            obj.update_listboxFile()  
        
        def update_listboxFile(obj):
            # Xóa sạch listbox
            obj.listboxFile.delete(0,'end') #Xóa sạch listbox trước khi nạp tránh lỗi ko cần thiết
            
            # covert fullpath --->fileName
            obj.wbs_ThisComputer_Exten = convertFullpath_to_fileName(obj.wbs_ThisComputer)
            obj.wbs_OtherComputer_Exten = convertFullpath_to_fileName(obj.wbs_OtherComputer)
            
            obj.total_wbs_ThisComputer = len(obj.wbs_ThisComputer)
            obj.total_wbs_OtherComputer = len(obj.wbs_OtherComputer)

            # Tiêu đề listbox
            obj.listMucLuc_Wbs = title_listbox()            
            obj.listMucLuc_LisxboxFile = [] #add các value trên listbox_File để khi thực hiện click ko làm gì

            obj.listboxFile.insert(END, obj.listMucLuc_Wbs[0])

            # Row tiêu đề đầu tiên
            obj.title_allFile_onlistboxFile=obj.endRows_listbox()

            obj.listboxFile.insert(END, f"I./ {obj.listMucLuc_Wbs[1]} {obj.total_wbs_ThisComputer} files")
            
            # Row tiêu đề wbs_ThisComputer
            obj.title_allFile_ThisCom_onlistboxFile=obj.endRows_listbox()

            # append vào obj.listMucLuc_LisxboxFile
            obj.listMucLuc_LisxboxFile.append(obj.listMucLuc_Wbs[0])
            obj.listMucLuc_LisxboxFile.append(f"I./ {obj.listMucLuc_Wbs[1]} {obj.total_wbs_ThisComputer} files")  
            obj.upFile_OnlistboxFile()#Up file xlsm lên listbox

        def upFile_OnlistboxFile(obj):
            if obj.total_wbs_OtherComputer>0: #Có sự tồn tại File trên Other Computer
                for i in range(obj.total_wbs_ThisComputer):
                    obj.listboxFile.insert(END, f"      {i+1}. {obj.wbs_ThisComputer_Exten[i]}")
                obj.listboxFile.insert(END, f"II./ {obj.listMucLuc_Wbs[2]} {obj.total_wbs_OtherComputer} files")

                # Row tiêu đề wbs_ThisOtherputer
                obj.title_allFile_OtherCom_onlistboxFile=obj.endRows_listbox()

                # tiếp tục append vào obj.listMucLuc_LisxboxFile
                obj.listMucLuc_LisxboxFile.append(f"II./ {obj.listMucLuc_Wbs[2]} {obj.total_wbs_OtherComputer} files")
                for i in range(obj.total_wbs_OtherComputer):
                    obj.listboxFile.insert(END, f"      {i+1}. {obj.wbs_OtherComputer_Exten[i]}")
            else:
                for i in range(obj.total_wbs_ThisComputer):
                    obj.listboxFile.insert(END, f"      {i+1}. {obj.wbs_ThisComputer_Exten[i]}")

        def onSelect_listboxFile(obj, event_click): #trả về giá trị tại vị trí click                      
            # obj.listboxFile.delete(0,END)
            obj.idx_ClickOnListboxFile=obj.listboxFile.curselection()
            obj.return_listItems_click()
            obj.showSheets_afterClickFile()

        def return_listItems_click(obj): 
            obj.listItemSelect_ThisComputer=[]
            obj.listItemSelect_OtherComputer=[]            
            # xử lý trường hợp ko click trên listboxFile
            try:
                obj.idx_ClickOnListboxFile=obj.listboxFile.curselection()
            except:                
                msgbox("Thông báo!!!", "Nhấn <<Ok!>> và chọn file trước khi thao tác", 0)
            else:
                # xử lý trường hợp có click trên listboxFile
                for obj.rowSelect_listboxFile in obj.idx_ClickOnListboxFile:
                    
                    #tiêu dề ------<<select file>> 
                    if obj.rowSelect_listboxFile == obj.title_allFile_onlistboxFile:                      
                        if obj.total_wbs_ThisComputer>0 and obj.total_wbs_OtherComputer>0:
                            obj.listboxFile.select_set(obj.title_allFile_onlistboxFile+1, END) #Chọn tô xanh toàn bộ listboxFile
                            obj.getAll_item_from_This_and_Other_wbs()
                        elif obj.total_wbs_ThisComputer>0:
                            obj.listboxFile.select_set(obj.title_allFile_ThisCom_onlistboxFile+1, obj.total_wbs_ThisComputer+1)#set(2, obj.total_wbs_ThisComputer+1): là toàn bộ dòng ThisComputer
                            obj.getAll_item_from_This_wbs()
                        elif obj.total_wbs_OtherComputer>0:
                            obj.listboxFile.select_set(obj.title_allFile_OtherCom_onlistboxFile+1, END)
                            obj.getAll_item_from_Other_wbs()

                    #Click tiêu đề danh sách All file ThisCompueter
                    elif obj.rowSelect_listboxFile == obj.title_allFile_ThisCom_onlistboxFile: 
                        if obj.total_wbs_ThisComputer>0:
                            obj.listboxFile.select_set(obj.title_allFile_ThisCom_onlistboxFile+1, obj.total_wbs_ThisComputer+1)#set(2, obj.total_wbs_ThisComputer+1): là toàn bộ dòng ThisComputer
                            obj.getAll_item_from_This_wbs()

                    # Click tiêu đề danh sách Allfile OtherCompueter: #tiêu đề danh sách file ThisCompueter
                    elif obj.rowSelect_listboxFile == obj.title_allFile_OtherCom_onlistboxFile:
                        if obj.total_wbs_OtherComputer>0:
                            obj.listboxFile.select_set(obj.title_allFile_OtherCom_onlistboxFile+1, END)
                            obj.getAll_item_from_Other_wbs()

                    # phần chọn không thuộc tiêu đề và thuộc wbs_ThisComputer
                    elif obj.rowSelect_listboxFile>obj.title_allFile_ThisCom_onlistboxFile and obj.rowSelect_listboxFile<(obj.title_allFile_OtherCom_onlistboxFile):
                        if obj.total_wbs_ThisComputer>0:
                            obj.add_otheritem_from_Thiscomputer(obj.rowSelect_listboxFile)
                        # # proces_Str = findStr_From_chrSpecial(obj.listboxFile.get(item), ".", "", left=True)
                    
                    # phần chọn không thuộc tiêu đề và thuộc wbs_OtherComputer
                    elif obj.rowSelect_listboxFile > (obj.title_allFile_OtherCom_onlistboxFile):#and obj.rowSelect_listboxFile < obj.listboxFile.index(END)
                        if obj.total_wbs_OtherComputer>0:
                            obj.add_otheritem_from_Othercomputer(obj.rowSelect_listboxFile)
                        
        def getAll_item_from_This_and_Other_wbs(obj):            
            for i in range(obj.total_wbs_ThisComputer):
                obj.listItemSelect_ThisComputer.append(i)            
            for i in range(obj.total_wbs_OtherComputer):
                obj.listItemSelect_OtherComputer.append(i)
            print(f"Index for this computer: {obj.listItemSelect_ThisComputer}; Index for Other computer: {obj.listItemSelect_OtherComputer}")
           
        def getAll_item_from_This_wbs(obj):            
            for i in range(obj.total_wbs_ThisComputer):
                obj.listItemSelect_ThisComputer.append(i)            
            print(f"Index for this computer: {obj.listItemSelect_ThisComputer}; Index for Other computer: {obj.listItemSelect_OtherComputer}")

        def getAll_item_from_Other_wbs(obj):
            for i in range(obj.total_wbs_OtherComputer):
                obj.listItemSelect_OtherComputer.append(i)
            print(f"Index for this computer: {obj.listItemSelect_OtherComputer}; Index for Other computer: {obj.listItemSelect_OtherComputer}")


        def add_otheritem_from_Thiscomputer(obj,item):
            proces_Str = findStr_From_chrSpecial(obj.listboxFile.get(item), ".", "", left=True)
            for i in range(len(obj.wbs_ThisComputer_Exten)):
                if obj.wbs_ThisComputer_Exten[i]== proces_Str:
                    # Kiểm tra sự tồn tại của proces_Str trong thisComputer với hàm isValueInList(value, listValue)
                    isFile = isValueInList(i, obj.listItemSelect_ThisComputer)
                    if isFile==False:                                    
                        obj.listItemSelect_ThisComputer.append(i)
                        break 
            print(f"Index for this computer: {obj.listItemSelect_ThisComputer}; Index for Other computer: {obj.listItemSelect_OtherComputer}")
            

        def add_otheritem_from_Othercomputer(obj,item):
            proces_Str = findStr_From_chrSpecial(obj.listboxFile.get(item), ".", "", left=True)
            for i in range(len(obj.wbs_OtherComputer_Exten)):
                if obj.wbs_OtherComputer_Exten[i]== proces_Str:
                        # Kiểm tra sự tồn tại của proces_Str trong OtherComputer với hàm isValueInList(value, listValue)
                        # Nếu trong obj.listItemSelect_OtherComputer ko có thì mới append vào
                        isFile = isValueInList(i, obj.listItemSelect_OtherComputer)
                        if isFile==False: 
                            # print("First:", obj.listItemSelect_OtherComputer)                                   
                            obj.listItemSelect_OtherComputer.append(i) 
                            # print("After:", obj.listItemSelect_OtherComputer)
                            break
            print(f"Index for this computer: {obj.listItemSelect_ThisComputer}; Index for Other computer: {obj.listItemSelect_OtherComputer}")

        def showSheets_afterClickFile(obj):

            total_file_ThisCom =len(obj.listItemSelect_ThisComputer)
            if obj.rowSelect_listboxFile == obj.title_allFile_onlistboxFile or obj.rowSelect_listboxFile == obj.title_allFile_ThisCom_onlistboxFile:
                pass

            else:
                # sheetShow_trueOrfalse=ooo

                for i in range(total_file_ThisCom):
                    file_getSheet=obj.wbs_ThisComputer[obj.listItemSelect_ThisComputer[i]]
                    if file_Exists(file_getSheet):
                        sheetName=get_SheetWithData_inPd_Exl(file_getSheet, SheetNames=True)
                        total_sheet = len(sheetName)

                        if total_sheet>0:

                            for j in range(len(sheetName)):
                                obj.listboxFile.insert(obj.rowSelect_listboxFile+j+1,f"            Sheet{j+1}:   |__ '{sheetName[j]}'" )
                        # Update lại vị trí của tiêu đề OtherCom
                        obj.update_title_allFile_OtherCom_onlistboxFile( total_sheet)
                        # if file_Exists:
                        #     sheetName=get_SheetWithData_inPd_Exl(excel_file,All=None, SheetNames=None, Data=None, conect_data=None)
        
        def update_title_allFile_OtherCom_onlistboxFile(obj, total_sheet):
            obj.title_allFile_OtherCom_onlistboxFile+=  total_sheet
            return obj.title_allFile_OtherCom_onlistboxFile

        def chk_mutiSelect(obj): 
            if obj.var1.get() == 1:
                obj.chkbox_mutiSelect_listboxFile.config(text='Đã chọn nhiều file')
                obj.listboxFile.config(selectmode=MULTIPLE)
            elif obj.var1.get() == 0:
                obj.chkbox_mutiSelect_listboxFile.config(text='Chọn nhiều files')
                obj.listboxFile.config(selectmode=EXTENDED) #selectmode=False
                
        def endVal_listbox(obj):
            val = obj.listboxFile.get(END)
            return val

        def endRows_listbox(obj):
            val = obj.endVal_listbox()
            if val != '':
                id_Endrows = obj.listboxFile.get(0,END).index(val)#bỏ cộng +1 và xử lý lại file writer txt                
            else: 
                id_Endrows = 0
            return id_Endrows                  

        def onClick_remove(obj):            
            # try:
            total_file_ThisCom =len(obj.listItemSelect_ThisComputer)
            total_file_OtherCom =len(obj.listItemSelect_OtherComputer) 
            listRemove_ThisCom=[]
            listRemove_OtherCom=[]

            if total_file_ThisCom>0:                
                for i in range(total_file_ThisCom):
                    listRemove_ThisCom.append(obj.wbs_ThisComputer[obj.listItemSelect_ThisComputer[i]])
                obj.wbs_ThisComputer=remove_duplicates_for_listmain(obj.wbs_ThisComputer, listRemove_ThisCom)
                obj.listboxFile.delete(0,END)           
                obj.update_listboxFile()

            if total_file_OtherCom>0:                
                for i in range(total_file_OtherCom):
                    listRemove_OtherCom.append(obj.wbs_OtherComputer[obj.listItemSelect_OtherComputer[i]])
                    
                obj.wbs_OtherComputer=remove_duplicates_for_listmain(obj.wbs_OtherComputer, listRemove_OtherCom)
                obj.listboxFile.delete(0,END)
                obj.update_listboxFile()
                print (obj.wbs_ThisComputer, obj.wbs_OtherComputer)
            # except: #Lỗi do nhấn Xóa mà ko chọn file                               
            #     msgbox("Thông báo!!!", "Nhấn <<OK>> và chọn file trước khi xóa file!!!", 0)
       
        def addfile_in_listbox(obj):           
            
            id_Endrows = obj.endRows_listbox()
            obj.newFile = askopenfilename(filetypes = (("Excel","*.xl*"),("all files","*.*")))#askopenfilename() # open seach file
            #filedialog.askopenfilename(filetypes = (("Text files","*.txt"),("all files","*.*")))

            #Add file vảo wbs
            obj.wbs_ThisComputer.append(obj.newFile) #.insert(index, val)

            # show tệp lên listbox
            obj.listboxFile.delete(0,END)
            obj.update_listboxFile()
            
                     
        def write_from_listbox_to_txt(obj):
            write_files_from_lists_to_txt(obj.file_txt, obj.wbs_ThisComputer, obj.wbs_OtherComputer)
            # write_files_wb_to_txt(obj.file_txt, obj.wbs)            
            
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