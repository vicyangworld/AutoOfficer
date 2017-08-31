import win32com
from win32com.client import Dispatch, constants
import os
import shutil
import gc
import re

class easyWord(object):
	"""docstring for ClassName"""
	def __init__(self, FileName):
		self.WordAPP = win32com.client.DispatchEx('Word.Application')
		self.WordAPP.Visible=False
		self.WordAPP.DisplayAlerts = False
		if FileName:
			self.FileName = FileName
			self.Doc = self.WordAPP.Documents.Open(FileName)
		else:
			raise "No input file!"
			return
	def __del__(self):
	 	self.Doc.Close()
	 	self.WordAPP.Quit()
	def PageCount(self):
		return self.WordAPP.ActiveWindow.ActivePane.Pages.Count
	def SetCell(self,R,C,Value):
		self.Doc.Tables[0].Rows[R].Cells[C].Range.Text = Value
	def ReadPersonNumbers(self):
		return self.Doc.Tables[1].Rows.Count-1
	def ReadCunZhang(self):
		return self.Doc.Tables[0].Rows[2].Cells[1].Range.Text

class easyExcel(object):
	"""docstring for easyExcel"""
	def __init__(self, FileName):
		self.ExcelApp = win32com.client.DispatchEx('Excel.Application')
		self.ExcelApp.Visible=False
		self.ExcelApp.DisplayAlerts = False
		if FileName:
			self.FileName = FileName
			self.Xls = self.ExcelApp.Workbooks.Open(self.FileName)
		else:
			raise "No input file!"
			return
	def __del__(self):
		self.Xls.Close()
		self.ExcelApp.Quit()
	def PageCount(self):
		"""目前只能计算只有一个工作表sheet的文档"""
		return (self.ExcelApp.ActiveSheet.VPageBreaks.Count)*(self.ExcelApp.ActiveSheet.HPageBreaks.Count)
	def read_areacode_time(self):
		tempList=[]
		TimesAdict = {}
		for x in range(1,self.ExcelApp.Worksheets.Count):
			Activesheet = self.ExcelApp.Worksheets(x)
			for i in range(1,Activesheet.UsedRange.Rows.Count):
				for j in range(0,Activesheet.UsedRange.Columns.Count):
					string = Activesheet.Rows[i].Cells[j].Value
					if string!=None:
						if j==0:
							key = re.sub("\D", "", string)
						else:
							# translate 2017.1.2  to  20170102
							time = (Activesheet.Rows[i].Cells[j].Value).split('.')
							if len(time[1])==1:
								time[1] = '0'+time[1]
							if len(time[2])==1:
								time[2] = '0'+time[2]
							tempList.append("".join(time))
				if not tempList:
					continue
				tempList2= tempList.copy()
				TimesAdict[key]=tempList2
				tempList.clear()
		return TimesAdict
class Job(object):
	"""docstring for Job"""
	def __init__(self, RootPath):
		#一些提示
		print("\n")
		print("**************************  特供赟哥软件  ******************************")
		print("****                                                       *************")
		print("****       将本程序放在根目录，运行之前请确保根目录下具有  *************")
		print("****   (1) *包含每个村民的个人目录                         *************")
		print("****   (2) *必须具有\"卷内文件目录\"文件模板(.doc或.docx)    *************")
		print("****   (3) 可以添加\"地区代码及时间表\"文件模板(.xls或.xlsx) *************")
		print("****   (4) 可以添加\"软卷皮封面\"文件模板(.xls或.xlsx)       *************")
		print("****                                                       *************")
		print("************************************************************************")
		self.RootPath = RootPath+"\\"
		self.Pages = 0
		self.adict = {}
		self.WithTime = True
		try:
			filesOrDirs = os.listdir(self.RootPath)
		except Exception:
			print("No such path: "+self.RootPath)
		else:
			print("扫描 "+ self.RootPath)
			if os.path.exists(self.RootPath+"卷内文件目录.doc"):
				pass
			else:
				print("在 "+ self.RootPath + " 没有找到\"卷内文件目录.doc\"")
				print("程序中断，请完善相应资料！")
				quit = input("按任意键退出...")
				return
			# 询问是否需要将“软卷皮封面.doc”考入个人目录
			while True:
				content = input("是否需要将\"软卷皮封面.doc\"一并考入个人目录? 请输入y/Y或者n/N:")
				if content=="y" or content=="Y" or content=="n" or content=="N":
					break
			if content=="y" or content=="Y":
				self.CopyFengmiam = True
			else:
				self.CopyFengmiam = False

			if "地区代码及时间表.xlsx" in filesOrDirs:
				Excel = easyExcel(self.RootPath+"地区代码及时间表.xlsx")
				self.TimesAdict = Excel.read_areacode_time()
				del Excel
			elif "地区代码及时间表.xls" in filesOrDirs:
				Excel = easyExcel(self.RootPath+"地区代码及时间表.xls")
				Excel.read_areacode_time()
				self.TimesAdict = Excel.read_areacode_time()
				del Excel
			else:
				print("在 "+self.RootPath+" 中缺失\"地区代码及时间表.xlsx\"或\"地区代码及时间表.xls\",无法完成目录表中时间自动填充!\"")

				while True:
					content = input("是否继续? 请输入y/Y或者n/N:")
					if content=="y" or content=="Y" or content=="n" or content=="N":
						break
				if content=="y" or content=="Y":
					self.WithTime = False
					pass
				else:
					print("程序中断，请完善相应资料！")
					quit = input("按任意键退出...")
					return
			nNumFile = 0;
			for x in filesOrDirs:
				if os.path.isdir(x) and x.startswith(('1','2','3','4','5','6','7','8','9','0')):
					nNumFile = nNumFile + 1

			print("共有 "+str(nNumFile) + " 户的资料需要统计.")
			if nNumFile==0:
				quit = input("按任意键退出...")
				return

			f1 = open('全部操作.txt','w')
			f2 = open('操作成功.txt','w')
			f3 = open('操作失败.txt','w')
			dirsCount = 0
			print("------------------------------------------")
			print("序号      户主编号与名字         操作状态")
			print("------------------------------------------")
			for fileOrDir in filesOrDirs:
				if os.path.isdir(fileOrDir) and fileOrDir.startswith(('1','2','3','4','5','6','7','8','9','0')):
					dirsCount = dirsCount + 1
					# 得到户主名字
					self.HuZhu = fileOrDir.split('_')[1]
					# 获取户主所在村庄的编号，前12位
					self.HuZhuVillageCode = (fileOrDir.split('_')[0])[0:12]

					self.FilesPath = self.RootPath + fileOrDir + "\\"
					subFilesOrDirs = os.listdir(fileOrDir)
					filesCount=0
					if os.path.exists(self.RootPath+"卷内文件目录.doc"):
						if os.path.exists(self.FilesPath +"卷内文件目录.doc"):
							os.remove(self.FilesPath +"卷内文件目录.doc")
						shutil.copyfile(self.RootPath+"卷内文件目录.doc", self.FilesPath +"卷内文件目录.doc")
					else:
						print("在 "+ self.RootPath + " 没有找到\"卷内文件目录.doc\"")
						print("程序中断，请完善相应资料！")
						quit = input("按任意键退出...")
						return
					if self.CopyFengmiam:
						if os.path.exists(self.RootPath+"软卷皮封面.doc"):
							if os.path.exists(self.FilesPath +"软卷皮封面.doc"):
								os.remove(self.FilesPath +"软卷皮封面.doc")
							shutil.copyfile(self.RootPath+"软卷皮封面.doc", self.FilesPath +"软卷皮封面.doc")
						else:
							print("在 "+ self.RootPath + " 没有找到\"软卷皮封面.doc\"")
							print("程序中断，请完善相应资料！")
							quit = input("按任意键退出...")
							return

					# 获取所需数据
					for files in subFilesOrDirs:
						if files.startswith(('1','2','3','4','5','6','7','8','9','0')):
							filesCount = filesCount + 1
							(filepath,tempfilename) = os.path.split(files)
							(filename,extension) = os.path.splitext(tempfilename)
							if extension==r".docx" or extension==r".doc":
								Word = easyWord(self.FilesPath+files)
								self.adict[filename] = Word.PageCount()
								self.Pages = Word.PageCount() + self.Pages
								if '登记簿' in files:
									self.CunZhang=Word.ReadCunZhang()[:-2] #去掉最后两个字符：一个是BEL,一个是换行
									self.PersonNumber = Word.ReadPersonNumbers()
								del Word
							elif extension==r".xlsx" or extension==r".xls":
								Excel = easyExcel(self.FilesPath+files)
								self.adict[filename] = Excel.PageCount()
								self.Pages = Excel.PageCount() + self.Pages
								del Excel

					if filesCount!=9 or len(files)==0:
						print("Files in \"" + os.getcwd() + "\" occur error!")
						print("Please Check the files")
						f3.write(str(dirsCount) + "    " + fileOrDir + "\n")
						f1.write(str(dirsCount) + "    " + fileOrDir + "   操作失败" + "\n")
						print(" "+str(dirsCount) + "   " +fileOrDir + "   操作失败")
						continue
					# 更新“卷内文件目录.doc”
					Word2 = easyWord(self.FilesPath+"卷内文件目录.doc")
					nTotal = 1
					Word2.SetCell(1,2,self.HuZhu) # "责任者"
					if self.WithTime:
						Word2.SetCell(1,4,(self.TimesAdict[self.HuZhuVillageCode])[0])  # "日期"
					Word2.SetCell(1,5,nTotal)     # "页号"
					for x in self.adict.keys():
						if '登记簿' in x:
							nTotal = self.adict[x]+nTotal
					Word2.SetCell(2,2,self.HuZhu)
					if self.WithTime:
						Word2.SetCell(2,4,(self.TimesAdict[self.HuZhuVillageCode])[1])  # "日期"
					Word2.SetCell(2,5,nTotal)
					for x in self.adict.keys():
						if '承包方调查表' in x:
							nTotal = self.adict[x]+nTotal
					Word2.SetCell(3,2,self.HuZhu)
					if self.WithTime:
						Word2.SetCell(3,4,(self.TimesAdict[self.HuZhuVillageCode])[2])  # "日期"
					Word2.SetCell(3,5,nTotal)
					for x in self.adict.keys():
						if '地块调查表' in x:
							nTotal = self.adict[x]+nTotal
					Word2.SetCell(4,2,self.HuZhu)
					if self.WithTime:
						Word2.SetCell(4,4,(self.TimesAdict[self.HuZhuVillageCode])[3])  # "日期"
					Word2.SetCell(4,5,nTotal)
					for x in self.adict.keys():
						if '公示结果归户表' in x:
							nTotal = self.adict[x]+nTotal
							#承包方推荐证明
					Word2.SetCell(5,2,self.HuZhu)
					if self.WithTime:
						Word2.SetCell(5,4,(self.TimesAdict[self.HuZhuVillageCode])[4])  # "日期"
					Word2.SetCell(5,5,nTotal)
					nTotal = nTotal+1
					Word2.SetCell(6,2,self.CunZhang+self.HuZhu)
					if self.WithTime:
						Word2.SetCell(6,4,(self.TimesAdict[self.HuZhuVillageCode])[5])  # "日期"
					Word2.SetCell(6,5,nTotal)
					for x in self.adict.keys():
						if '承包合同' in x:
							nTotal = self.adict[x]+nTotal
					#户主户口本及身份证复印件
					Word2.SetCell(7,2,self.HuZhu)
					if self.WithTime:
						Word2.SetCell(7,4,(self.TimesAdict[self.HuZhuVillageCode])[6])  # "日期"
					Word2.SetCell(7,5,str(nTotal)+"-"+str(nTotal+self.PersonNumber))
					self.adict.clear()
					del Word2
					gc.collect()
					f2.write(str(dirsCount) + "    " +fileOrDir + "\n")
					f1.write(str(dirsCount) + "    " +fileOrDir + "   操作成功" + "\n")
					print(" "+str(dirsCount) + "   " +fileOrDir + "   操作成功")
			print("------------------------------------------")
			f1.close()
			f2.close()
			f3.close()

if __name__ == '__main__':
	rootpath = os.getcwd()
	Jobb = Job(rootpath)