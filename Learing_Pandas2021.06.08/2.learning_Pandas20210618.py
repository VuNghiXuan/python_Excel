import pandas as pd
# import xlsxwriter
import os
import numpy as np
from pandas.core.frame import DataFrame
# from Package_VuNghiXuan_Excel.useFunc import *
from Package_Pandas.func_method_class_ExcelFile_Pandas import *
from Package_Pandas.func_method_read_excel_Pandas import *

def main():
    pth= os.getcwd()
    file_1 = pth + "/In_Excel/TToanGopQ3.4.xlsm"
    file_2 = pth + "/In_Excel/In.xlsx"
    out_file_1 = pth+ "/Out_Excel/out_1.xlsx"
    out_file_2 = pth+ "/Out_Excel/out_KLPS.xlsx"
    
    
    '''In đậm tiêu đề
    # header_fmt = workbook.add_format({'bold': True})
    # worksheet.set_row(0, None, header_fmt)'''
    
    # 1. Lấy nhanh tên sheeets qua class ExcelFile print printSceen
    list_sheet_1 = get_list_sheetname_from_Filepath(file_1)
    # print(list_sheet_1)
    
    # 2. read file method read_excel Lấy sheet TheoDoi_HopDong làm việc
    df = pd.read_excel(file_1, sheet_name = 'TheoDoi_BDTX', 
                            skiprows= 9, skipfooter= 15, 
                            )#, None) index_col=1,# chú ý có None, index_col=1 lấy cột thứ 2 làm index
    
    # 3. sceenPrint cột
    # print(df.columns)

    # 4. Dict Value 'HangMuc_ChinhSua': Lọc danh sách duy nhất
    # dict_HangMuc= df["HangMuc_ChinhSua"].unique()#unique: độc nhất
    
    # 5. Chọn cột sử dụng
    """Cách 1"""
    df1 = df[['TT_Tuyen','STT', 'HangMuc', 'DVT', 'KLSGT',       
       'TongDieuChinh_KhoiluongKeHoach_Q3,4', 'DG_So']]
    
    """Cách 2"""
    # filtered_columns = ['MHM', 'HangMuc_ChinhSua', 'DVT', 'KL_DuThau', 'DG_DuThau','TT_DuThau']
    # df=df.reindex(columns = filtered_columns)
    """ Check 
    type(df["KL_DuThau"]) # có dạng pandas.core.series.Series
    df.["KL_DuThau"] >>> cho số row 
    df.iloc[9:25, 2:5]>>>Lọc hàng 10 đến 25 và cột 3 đến 5.   """
    
    #  6 Thêm cột : 
    df1['Kl(thuchien-Kehoach)']=df1["TongDieuChinh_KhoiluongKeHoach_Q3,4"]- df["KLSGT"]
    
    # Gán biến dieukien_KL
    dieukien_KL = df1["TongDieuChinh_KhoiluongKeHoach_Q3,4"]- df1["KLSGT"]
    
    # 6a. Tạo cột KL_PS
    """Cách 1: Gán trực tiếp bằng thư viện np với np là mảng lọc giá trị số.
    Trong đó: 
        + df['KL_PS']: Tạo thêm cột mới 
        + Điều kiện: df["TongDieuChinh_KhoiluongKeHoach_Q3,4"]- df["KLSGT"] >0 : thì KL-Thuchien - KL_KeHoach
        + None : ngược lại """
    df1['KL_PS: Cách 1'] = np.where(dieukien_KL>0,dieukien_KL,None)
    
    # # Tạm thời comment cách 2: 
    # """Cách 2: df._get_numeric_data(): Lấy giá trị là số. 
    #     Tuy nhiên: num_df[num_df < 0] = None sẽ thay luôn cột df['Kl(thuchien-Kehoach)'] ko còn số âm"""
    # df['KL_PS Cách 2'] = dieukien_KL
    # num_df= df._get_numeric_data() # giữ lại nguyên bản df, 
    # num_df[num_df < 0] = None
   
    
    # 7. Lấy value >1000 cột Kl_Duthau
    # Kl_DThau_Above_100 = df[df["KL_DuThau"]>100]
    # print(Kl_DThau_Above_100)   
    # df = Kl_DThau_Above_100.replace(np.nan,'', regex=False)
    # df=Kl_DThau_Above_100.to_excel(out_file_2, sheet_name="TheoDoi_HopDongxxx")
    
    # 8. Lọc cột theo diều kiện với:     
    # df_Col_condition = df.loc[(df["MT"]>0)|(df["HangMuc_ChinhSua"] =="Đào hốt đất sụt bằng thủ công") 
    #     & (df["KL_DuThau"] >120)]

    # """Ghi chú lỗi: df["MT"]>0 phải đặt trong dấu ngoặc 
    #  """
    
    # 9. Loại bỏ số âm trong df    
    """ - Cách 1: này bỏ giá trị âm--------     """
    ## num=df.select_dtypes(include=[np.number]) # chọn số và bỏ luôn các String 
      
    """ Nhớ là kết quả: print df, chứ ko phải là num 
    *Chưa thử : for k, v in df.iteritems():    v[v < 0] = 0 """

    
    # df.loc[(df['Salary_in_1000']>=100) & (df['Age']< 60) & 
    # (df['FT_Team'].str.startswith('S')),['Name','FT_Team']] #https://kanoki.org/2020/01/21/pandas-dataframe-filter-with-multiple-conditions/
    
    df = df.replace(np.nan,'', regex=False) #regex: regexs khớp với to_replace sẽ được thay thế bằng giá trị
    df.to_excel(out_file_1, sheet_name="PS_Tang")
    
    # Lưu file nhiều sheets
    list_dfs =[]
    list_dfs.append(df)
    list_dfs.append(df1)
    save_xlsm(list_dfs, out_file_2)
    # print(df)

    # ???. Remove Nan type(df["Age"])
    # df = df.replace(np.nan,'', regex=False)
    # df=df.to_excel(out_file_2, sheet_name="TheoDoi_HopDongxxx")
    print ("20210611")
    

""" Remember: 
    Khi chọn tập hợp con dữ liệu, dấu ngoặc vuông [] được sử dụng.

    Bên trong các dấu ngoặc này, bạn có thể sử dụng một nhãn cột / hàng, danh sách các nhãn cột / hàng, một phần nhãn, một biểu thức điều kiện hoặc dấu hai chấm.

    Chọn các hàng và / hoặc cột cụ thể bằng cách sử dụng loc khi sử dụng tên hàng và cột

    Chọn các hàng và / hoặc cột cụ thể bằng iloc khi sử dụng các vị trí trong bảng
    
            " loc: biết cụ thể là dòng cột nào
            ===>iloc giống range: Phạm vi"

    Bạn có thể gán các giá trị mới cho một lựa chọn dựa trên loc / iloc.
"""
"""CHưa test ---------------------
str(df['fyear'])
    for row in df.iterrows():
        if df["fyear"] != 2009 | df["fyear"] !=2019 | df["fyear"] !=2020:
        df.drop(row)
        --------------------------
"""

if __name__=='__main__':   #__name__=='__main__': khi nào là hàm main() nằm trong bảng code náy mới dc chạy
    main()	