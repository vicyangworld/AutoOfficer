import win32com
from win32com.client import Dispatch, constants
import os
import shutil
import gc

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
		return self.Doc.Tables.Item(1).Rows-1
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

class Job(object):
	"""docstring for Job"""
	def __init__(self, RootPath):
		self.RootPath = RootPath
		self.Pages = 0
		self.adict = {}
		try:
			filesOrDirs = os.listdir(self.RootPath)
		except Exception:
			print("No such path: "+self.RootPath)
		else:
			f1 = open('log1.txt','w')
			f2 = open('log2.txt','w')
			f3 = open('log3.txt','w')
			dirsCount = 0
			for fileOrDir in filesOrDirs:
				if os.path.isdir(fileOrDir):
					# 得到户主名字
					self.HuZhu = fileOrDir.split('_')[1]
					self.FilesPath = self.RootPath + fileOrDir + "\\"
					subFilesOrDirs = os.listdir(fileOrDir)
					filesCount=0
					dirsCount = dirsCount + 1

					if os.path.exists(self.RootPath+"卷内文件目录.doc"):
						if os.path.exists(self.FilesPath +"卷内文件目录.doc"):
							os.remove(self.FilesPath +"卷内文件目录.doc")
						shutil.copyfile(self.RootPath+"卷内文件目录.doc", self.FilesPath +"卷内文件目录.doc")

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
								del Word
							elif extension==r".xlsx" or extension==r".xls":
								Excel = easyExcel(self.FilesPath+files)
								self.adict[filename] = Excel.PageCount()
								self.Pages = Excel.PageCount() + self.Pages
								del Excel

					if filesCount!=9 or len(files)==0:
						print("Files in \"" + os.getcwd() + "\" occur error!")
						print("Please Check the files")
						f3.write(str(dirsCount) + "  " + fileOrDir + "\n")
						f1.write(str(dirsCount) + "  " + fileOrDir + ": 1" + "\n")
						continue

					# 更新“卷内文件目录.doc”
					Word2 = easyWord(self.FilesPath+"卷内文件目录.doc")
					nTotal = 1
					Word2.SetCell(1,2,self.HuZhu) # "责任者"
					Word2.SetCell(1,5,nTotal)     # "页号"
					for x in self.adict.keys():
						if '登记簿' in x:
							nTotal = self.adict[x]+nTotal
					Word2.SetCell(2,2,self.HuZhu)
					Word2.SetCell(2,5,nTotal)
					for x in self.adict.keys():
						if '承包方调查表' in x:
							nTotal = self.adict[x]+nTotal
					Word2.SetCell(3,2,self.HuZhu)
					Word2.SetCell(3,5,nTotal)
					for x in self.adict.keys():
						if '地块调查表' in x:
							nTotal = self.adict[x]+nTotal
					Word2.SetCell(4,2,self.HuZhu)
					Word2.SetCell(4,5,nTotal)
					for x in self.adict.keys():
						if '公示结果归户表' in x:
							nTotal = self.adict[x]+nTotal
							#承包方推荐证明
					Word2.SetCell(5,2,self.HuZhu)
					Word2.SetCell(5,5,nTotal)
					nTotal = nTotal+1
					Word2.SetCell(6,2,self.CunZhang+self.HuZhu)
					Word2.SetCell(6,5,nTotal)
					for x in self.adict.keys():
						if '承包合同' in x:
							nTotal = self.adict[x]+nTotal
					#户主户口本及身份证复印件
					Word2.SetCell(7,2,self.HuZhu)
					Word2.SetCell(7,5,nTotal)
					del Word2
					gc.collect()
					f2.write(str(dirsCount) + "  " +fileOrDir + "\n")
					f1.write(str(dirsCount) + "  " +fileOrDir + ": 0" + "\n")
			f1.close()
			f2.close()
			f3.close()

if __name__ == '__main__':
	rootpath = "E:\\PythonProjects\\AutoPagesCounter\\"
	Jobb = Job(rootpath)
