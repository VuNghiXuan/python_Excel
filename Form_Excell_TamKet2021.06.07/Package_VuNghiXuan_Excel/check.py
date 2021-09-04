# import show_sceen as show
# print(dir(check))

def select_range_number(choice, min_num, max_num):
	# choice = input(promt)	

	while not choice.isdigit() or int(choice) < min_num or int(choice) > max_num:
		choice = input(f"Hãy chọn số từ ({min_num}:{max_num}): ") #print(f"Số nhập vào từ ({min_num}:{max_num}):")		
		
	choice = int(choice)
	return choice



def print_screen():
	select_range_number(3, 1, 3)

def main():
    print_screen()
    
if __name__=='__main__':   #__name__=='__main__': khi nào là hàm main() nằm trong bảng code náy mới dc chạy
    main()	
