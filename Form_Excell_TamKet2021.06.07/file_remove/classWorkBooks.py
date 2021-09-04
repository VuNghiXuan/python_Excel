import pandas as pd
class Workbook:
    def __init__(self, sheetName):#, numRows, numCols
        self.sheetName = sheetName
        # self.numRows = numRows
        # self.numCols = numCols

def main():
    excel_file ="D:/ThanhVu/code/python/pyExell/Data_20210425/File_Test_Excel/movies (1).xls"
    # excel_file = "I:/Code/Python/pyExell/python_excel/Data_20210426/File_Test_Excel/movies.xls"

    # *********Mai sử dụng dc
    xlsx = pd.ExcelFile(excel_file)
    # print("tên sheet:", xlsx.sheet_names) 
    # print("Số sheet:", len(xlsx.sheet_names))
    df = xlsx.parse(xlsx.sheet_names[0])    
    print(df.shape[0])#df, Column and row
    a = Workbook(xlsx.sheet_names[0])
    print (a.sheetName)

if __name__=='__main__':   #__name__=='__main__': khi nào là hàm main() nằm trong bảng code náy mới dc chạy
    main()