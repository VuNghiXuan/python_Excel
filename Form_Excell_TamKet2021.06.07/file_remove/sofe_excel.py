from Package_VuNghiXuan_Excel import ioData 
from Package_VuNghiXuan_Excel import formExcel

# print("danh sách các hàm trong modul check", check.__name__)#cách xem danh sách các hàm trong modul check
# print(check.__name__)#cách xem danh sách các hàm trong modul check

def main():
    # # Bước 1 (hoàn thiện lấy thông tin người dùng và save vào file txt)
    # #Show ra màn hình các option   
    # ioData.option_excel()
    # #Check số từ người dùng và nhận về giá trị nhập từ số
    # ioData.choice_option_menu()
    # # Bước 1 <----------------------------------------------------------------
    # Bước 2: Sử dụng form điều khiển
    formExcel.show_form() 

if __name__=='__main__':   #__name__=='__main__': khi nào là hàm main() nằm trong bảng code náy mới dc chạy
    main()
