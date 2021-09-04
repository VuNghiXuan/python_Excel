import xlwings as xw
import os
class Workbook():
	def __init__(self, path, name):
		self.path = str(path)
		self.name = str(name)
		# self.sheets = sheets
		self.active = False

		self.pathfull = os.path.join(self.path,self.name)

		print("self.pathfull------ " + self.pathfull +"\n")

	def num_sheets(self):
		
		# count_sh = wb.api.Sheets.Count
		num_sheet = self.wb.api.Sheets.Count
		print(num_sheet)
		return num_sheet

	def open(self):
		# print("begin open(self)======    self.wb =========", self.pathfull +"\n" )
		self.wb = xw.Book(self.pathfull)   #  = self.wb = 
		self.active = True

		print("thử cho print ==== ", self.wb )
		return (self.wb )