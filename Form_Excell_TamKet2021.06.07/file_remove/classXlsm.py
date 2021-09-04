
class Wookbook: #Video
	def __init__(self, title_Exten, link ):
		self.title_Exten = title_Exten
		self.link = link
		self.seen = False

	def open(self):
		# webbrowser.open(self.link)#su dung cách mo file xlsm sau
		self.seen = True
		
class ListFile: #playlist
	def __init__(self, name, description, wookbooks):#videos = wookbooks
		self.name = name # Tên máy This Computer		
		self.wookbooks = wookbooks
