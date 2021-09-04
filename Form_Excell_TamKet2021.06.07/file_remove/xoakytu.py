# Add key, value vào dict
def createDict_from_list_assign_False(lis):
	dict={}
	for i in lis:
		dict[i] = False
		# print("Trước:",i, dict[i])
	return dict

def isTrueorFalse_inDict(*dicts):
	for dict in dicts:
		for i in dict.keys():
			if dict[i]==False:
				dict[i]=True
			else:
				dict[i]=False
			# print("Sau:",i, dict[i])
	return dicts
