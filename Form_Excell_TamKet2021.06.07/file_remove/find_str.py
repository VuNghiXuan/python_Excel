# Tìm kiếm dòng và cột. Ex: a='["DVT  "]=m2,["HangMuc"]="Cắt cỏ"'
# kết quả:>>> ['DVT  ', 'HangMuc'] ['m2', 'Cắt cỏ']
def find_valueStr_OfRowCol_on_trView(into_str):
	
    list_value_Cols=[]
    list_value_Rows=[]
    chr_first_Col = "["
    chr_after_Col = "]"
    chr_first_Row = "="
    chr_after_Row = ","

    char_special ='"'
    begin_valCol = ""
    after_valCol = ""
    begin_valRow = ""
    after_valRow = ""
    
    if into_str.find("[") !=-1:
		
        for i in range(len(into_str)):
            if into_str[i] == chr_first_Col:
                begin_valCol = i
                
            elif into_str[i] == chr_after_Col:
                after_valCol = i
                value_Col = into_str[begin_valCol+1:after_valCol]
                value_Col = value_Col.replace(char_special,"")
                value_Col=value_Col.strip()
                list_value_Cols.append(value_Col)

            elif into_str[i] == chr_first_Row:
                begin_valRow = i
                
            elif into_str[i] == chr_after_Row:
                after_valRow = i
                value_Row = into_str[begin_valRow+1:after_valRow]
                value_Row = value_Row.replace(char_special,"")
                value_Col=value_Col.strip()
                list_value_Rows.append(value_Row)
            elif i == len(into_str)-1:
                after_valRow = i
                value_Row = into_str[begin_valRow+1:after_valRow+1]
                value_Row = value_Row.replace(char_special,"")
                value_Col=value_Col.strip()
                list_value_Rows.append(value_Row)
            		
		# out_Str = into_str.replace(list_value_Cols[0], replace_str)
		# out_Str = out_Str.lstrip()
        return list_value_Cols, list_value_Rows
        
a= '["title  "]="Boss",["Genres"]="Roman"'
# a='["DVT  "]=m2,["HangMuc"]="Cắt cỏ"'
list_value_Cols, list_value_Rows = find_valueStr_OfRowCol_on_trView(a)
print (list_value_Cols, list_value_Rows)