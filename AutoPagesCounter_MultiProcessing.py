import win32com
import cmd_format
from win32com.client import Dispatch, constants
import shutil
import gc
import re
import copy
import os
from multiprocessing import Pool
from ProgressBar import ProgressBar

CDMF = cmd_format.CmdFormat("特供赟哥")

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
			raise IOError("文件 "+FileName+" 没找到！")
			return
	def PageCount(self):
		return self.WordAPP.ActiveWindow.ActivePane.Pages.Count
	def SetCell(self,R,C,Value,TableIndex=0,FontSize=-1):
		#Range是一个非常重要的概念，可以设置字体，行间距，文本！！
		self.Doc.Tables[TableIndex].Rows[R].Cells[C].Range.Text = Value
		if FontSize!=-1:
			self.Doc.Tables[TableIndex].Rows[R].Cells[C].Range.Font.Size = FontSize
	def GetCell(self,R,C,TableIndex=0):
		return self.Doc.Tables[TableIndex].Rows[R].Cells[C].Range.Text
	def ReadPersonNumbers(self):
		return self.Doc.Tables[1].Rows.Count-1
	def ReadCunZhang(self):
		return self.Doc.Tables[0].Rows[2].Cells[1].Range.Text
	def Close(self):
		self.Doc.Save()
		self.Doc.Close()
		self.WordAPP.Quit()
		del self.WordAPP

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
			raise IOError("文件 "+FileName+" 没找到！")
			return
	def Close(self):
		self.Xls.Close(SaveChanges=0)
		self.ExcelApp.Quit()
		del self.ExcelApp
		#self.ExcelApp.Quit()
	def PageCount(self):
		pages = 0
		for x in range(1,self.Xls.Worksheets.Count+1):
			Activesheet = self.Xls.Worksheets(x)
			#执行到Sheet1.HPageBreaks.Count的时候，它才强制分页了，所以先运行一次
			# Activesheet.VPageBreaks.Count
			# Activesheet.HPageBreaks.Count
			pages = pages + (Activesheet.VPageBreaks.Count)*(Activesheet.HPageBreaks.Count)
			# self.ExcelApp.Volatile
		return pages
	def read_areacode_time(self):
		AreaTimeAdict = {}
		for x in range(1,self.Xls.Worksheets.Count+1):
			nValidRows = 0
			Activesheet = self.Xls.Worksheets(x)
			CDMF.print_blue_text("提取 "+Activesheet.Name +" 信息...",endd='')
			# UsedRange 从1开始
			nTempCode = Activesheet.Cells(1,1).Value
			nRows = Activesheet.UsedRange.Rows.Count
			nColumns = Activesheet.UsedRange.Columns.Count
			for i in range(2,nRows+1):
				tempList=[]
				for j in range(1,nColumns+1):
					string = Activesheet.Cells(i,j).Value
					if string!=None:
						if j==1:
							key = re.sub("\D", "", string)
							strRe = (re.split(r'[县镇乡村]',string)[1:3])
							strRe.append(str(nTempCode)) #加入乡镇编号，用于填充“软卷皮封面.doc”中的分类号
							tempList.append(strRe)
						elif j<=8 and j>1:
							# translate 2017.1.2  to  20170102
							time = string.split('.')
							if len(time)!=3:
								continue
							if len(time[1])==1:
								time[1] = '0'+time[1]
							if len(time[2])==1:
								time[2] = '0'+time[2]
							tempList.append("".join(time))
						else:
							tempList.append(string) #村的编号，用于填充“软卷皮封面.doc”中的分类号
				if not tempList:
					continue
				nValidRows = nValidRows + 1
				#这是一个疑点,为什么要加一个.copy()，没有弄清楚还
				AreaTimeAdict[key]=tempList.copy()
				tempList.clear()
			CDMF.print_blue_text("成功, 共有 "+str(nValidRows)+" 个村庄.")
		CDMF.print_blue_text("有效村庄共有 "+str(len(AreaTimeAdict))+" 个.")
		return AreaTimeAdict


# 核心函数--------------------------------------------------------------------------------------------
def Tasks(fileOrDir,RootPath,AreaTimeAdict,bRegenerateContent,bWithTime,bFillCover):
	# 得到户主名字
	HuZhu = fileOrDir.split('_')[1]
	# 获取户主所在村庄的编号，前12位
	HuZhuVillageCode = (fileOrDir.split('_')[0])[0:12]   #个人所在的村的Code
	HuzhuPersonalCode= (fileOrDir.split('_')[0])[-3:]  #个人的Code，可以填在“软卷皮封面.doc”中的案卷号中

	PersionalDir = RootPath + fileOrDir + "\\"
	FileOrDirInPersionalDir_list = os.listdir(fileOrDir)
	filesCount=0
	#到这行，卷内文件目录.doc肯定存在，无需判断
	if bRegenerateContent:
		if os.path.exists(PersionalDir +"卷内文件目录.doc"):
			os.remove(PersionalDir +"卷内文件目录.doc")
		shutil.copyfile(RootPath+"卷内文件目录.doc", PersionalDir +"卷内文件目录.doc")
	else:
		if os.path.exists(PersionalDir +"卷内文件目录.doc"):
			return
		else:
			shutil.copyfile(RootPath+"卷内文件目录.doc", PersionalDir +"卷内文件目录.doc")
	#到这行，软卷皮封面.doc肯定存在或者self.CopyFengmian为false
	if bFillCover:
		if os.path.exists(RootPath+"软卷皮封面.doc"):
			if os.path.exists(PersionalDir +"软卷皮封面.doc"):
				os.remove(PersionalDir +"软卷皮封面.doc")
			shutil.copyfile(RootPath+"软卷皮封面.doc", PersionalDir +"软卷皮封面.doc")
	# 获取所需数据
	Pages_adict = {}
	try:
		#对于一个特定的村民文件夹
		for file in FileOrDirInPersionalDir_list:
			if not os.path.isdir(file) and file.startswith(('1','2','3','4','5','6','7','8','9','0')):
				filesCount = filesCount + 1
				(filepath,tempfilename) = os.path.split(file)
				(filename,extension) = os.path.splitext(tempfilename)
				if extension==r".docx" or extension==r".doc":
					Word = easyWord(PersionalDir+file)
					Pages_adict[filename] = Word.PageCount()
					if '登记簿' in file:
						CunZhang=Word.ReadCunZhang()[:-2] #去掉最后两个字符：一个是BEL,一个是换行
						PersonNumber = Word.ReadPersonNumbers()
					Word.Close()
				elif extension==r".xlsx" or extension==r".xls":
					Excel = easyExcel(PersionalDir+file)
					Pages_adict[filename] = Excel.PageCount()
					Excel.Close()
		nTemp = 0
		bFirst = True
		for x in Pages_adict.keys():
			if '登记簿' in x:
				nTemp = nTemp + 1
			if '承包合同' in x:
				nTemp = nTemp + 1
			if '地块调查表' in x and bFirst:
				nTemp = nTemp + 1
				bFirst = False
			if '公示结果归户表' in x:
				nTemp = nTemp + 1
		if nTemp != 4: raise IOError
	except Exception as e:
		CDMF.print_red_text(fileOrDir + "    操作失败")
		f3.write( fileOrDir + "\n")
		f1.write( fileOrDir + "    操作失败" + "\n")
		if os.path.exists(PersionalDir +"卷内文件目录.doc"):
			os.remove(PersionalDir +"卷内文件目录.doc")
		bar.Move('')
		return
	else:
		pass
	finally:
		pass
	# 更新“卷内文件目录.doc”
	try:
		Word = easyWord(PersionalDir+"卷内文件目录.doc")
		# 填写目录的第1顺序号
		nTotalPages = 1
		Word.SetCell(1,2,HuZhu) # "责任者"
		if bWithTime:
			Word.SetCell(1,4,(AreaTimeAdict[HuZhuVillageCode])[1])  # "日期"
		Word.SetCell(1,5,nTotalPages)     # "页号"

		# 填写目录的第2顺序号
		for x in Pages_adict.keys():
			if '登记簿' in x:
				nTotalPages = Pages_adict[x]+nTotalPages
				bDjb = True
				break
		if not bDjb:
			raise OverflowError("没有找到 “承包经营权登记簿”")
		Word.SetCell(2,2,HuZhu)
		if bWithTime:
			Word.SetCell(2,4,(AreaTimeAdict[HuZhuVillageCode])[2])  # "日期"
		Word.SetCell(2,5,nTotalPages)

		# 填写目录的第3顺序号
		bDjb = False
		for x in Pages_adict.keys():
			if '承包方调查表' in x:
				nTotalPages = Pages_adict[x]+nTotalPages
				bDjb = True
				break
		if not bDjb:
			nTotalPages = nTotalPages + 1 #如果没有承包方调查表，默认承包方调查表为1页
		Word.SetCell(3,2,HuZhu)
		if bWithTime:
			Word.SetCell(3,4,(AreaTimeAdict[HuZhuVillageCode])[3])  # "日期"
		Word.SetCell(3,5,nTotalPages)

		# 填写目录的第4顺序号
		for x in Pages_adict.keys():
			if '地块调查表' in x:
				nTotalPages = Pages_adict[x]+nTotalPages
		Word.SetCell(4,2,HuZhu)
		if bWithTime:
			Word.SetCell(4,4,(AreaTimeAdict[HuZhuVillageCode])[4])  # "日期"
		Word.SetCell(4,5,nTotalPages)

		# 填写目录的第5顺序号
		for x in Pages_adict.keys():
			if '公示结果归户表' in x:
				nTotalPages = Pages_adict[x]+nTotalPages
				#承包方推荐证明
		Word.SetCell(5,2,HuZhu)
		if bWithTime:
			Word.SetCell(5,4,(AreaTimeAdict[HuZhuVillageCode])[5])  # "日期"
		Word.SetCell(5,5,nTotalPages)
		nTotalPages = nTotalPages+1

		# 填写目录的第6顺序号
		Word.SetCell(6,2,CunZhang+HuZhu)
		if bWithTime:
			Word.SetCell(6,4,(AreaTimeAdict[HuZhuVillageCode])[6])  # "日期"
		Word.SetCell(6,5,nTotalPages)
		for x in Pages_adict.keys():
			if '承包合同' in x:
				nTotalPages = Pages_adict[x]+nTotalPages

		# 填写目录的第7顺序号
		#户主户口本及身份证复印件
		Word.SetCell(7,2,HuZhu)
		if bWithTime:
			Word.SetCell(7,4,(AreaTimeAdict[HuZhuVillageCode])[7])  # "日期"
		Word.SetCell(7,5,str(nTotalPages)+"-"+str(nTotalPages+PersonNumber))
		nTotalPages = nTotalPages +PersonNumber
		Pages_adict.clear()
		Word.Close()
	except Exception as e:
		CDMF.print_red_text("出错！请检查是否在 "+PersionalDir+"存在 \"卷内文件目录.doc\" .")
	else:
		pass
	finally:
		pass

	# 更新软卷皮封面.doc”
	try:
		Word = easyWord(PersionalDir+"软卷皮封面.doc")

		mxxx = re.split(r'([镇村（])', Word.GetCell(2,0))  # r'([镇村（])'加括号保留分隔符
		mxxx[0] = '\r'+(AreaTimeAdict[HuZhuVillageCode])[0][0]
		mxxx[2] = (AreaTimeAdict[HuZhuVillageCode])[0][1]
		mxxx[4] = HuZhu
		Word.SetCell(2,0,''.join(mxxx),FontSize=18)

		mxxx = re.split(r'([自年(月至)])', Word.GetCell(3,0))
		if mxxx[-1]!='月':
			mxxx.pop()
		if mxxx[0]!='自':
			del mxxx[0]
		if len(mxxx) < 10:
			raise OverflowError("需要检查软卷皮封面模板中“自X年X月至X年X月”")
			return
		mxxx[1] = ''.join(list((AreaTimeAdict[HuZhuVillageCode])[7])[0:4])
		if list((AreaTimeAdict[HuZhuVillageCode])[7])[4]=='0':
			mxxx[3] = ''.join(list((AreaTimeAdict[HuZhuVillageCode])[7])[5:6])
		else:
			mxxx[3] = ''.join(list((AreaTimeAdict[HuZhuVillageCode])[7])[4:6])
		mxxx[7] = ''.join(list((AreaTimeAdict[HuZhuVillageCode])[1])[0:4])
		if list((AreaTimeAdict[HuZhuVillageCode])[1])[4]=='0':
			mxxx[9] = ''.join(list((AreaTimeAdict[HuZhuVillageCode])[1])[5:6])
		else:
			mxxx[9] = ''.join(list((AreaTimeAdict[HuZhuVillageCode])[1])[4:6])
		Word.SetCell(3,0,''.join(mxxx),FontSize=18)

		mxxx = re.split(r'([共件页])', Word.GetCell(4,0))
		if mxxx[-1]!='页':
			mxxx.pop()
		if mxxx[0]!='本卷':
			del mxxx[0]
		if len(mxxx) < 6:
			raise OverflowError("需要检查软卷皮封面模板中“本卷共X件X页”")
			return
		mxxx[2] = '   7   '
		mxxx[4] = '   '+str(nTotalPages)+'   '
		Word.SetCell(4,0,''.join(mxxx),FontSize=18)

		#设置表2的全宗号
		Word.SetCell(1,0,'53',TableIndex=1,FontSize=12)#小四
		#设置表2的分类号
		strTemp = 'TQ0202'+(AreaTimeAdict[HuZhuVillageCode])[0][2]+(AreaTimeAdict[HuZhuVillageCode])[8]
		Word.SetCell(1,1,strTemp,TableIndex=1,FontSize=12)#小四
		#设置表2的案卷号
		Word.SetCell(1,2,str(HuzhuPersonalCode),TableIndex=1,FontSize=12)#小四
		Word.Close()
	except Exception as e:
		CDMF.print_red_text("出错！请检查是否在 "+PersionalDir+"存在 \"软卷皮封面.doc\" .")
	else:
		pass
	finally:
		pass

	# gc.collect()
	f2.write(fileOrDir + "\n")
	f1.write(fileOrDir + "    操作成功" + "\n")
	print(fileOrDir + "    操作成功")
	bar.Move('')
# ---------------------------------------------------------------------------------------------------

class Job(object):
	"""docstring for Job"""
	def __init__(self, RootPath):
		#一些提示
		CDMF.set_cmd_color(cmd_format.FOREGROUND_RED | cmd_format.FOREGROUND_GREEN | \
			cmd_format.FOREGROUND_BLUE | cmd_format.FOREGROUND_INTENSITY)
		print("\n")
		print("==========================  特供赟哥软件  ==============================")
		print("|                                                                      |")
		print("|      将本程序放在根目录，运行之前请确保根目录下具有                  |")
		CDMF.print_red_text("|      (1) *包含每个村民的个人目录                                     |")
		CDMF.print_red_text("|      (2) *必须具有\"卷内文件目录\"文件模板(.doc或.docx)                |")
		print("|      (3) 可以添加\"地区代码及时间表\"文件模板(.xls或.xlsx)             |")
		print("|      (4) 可以添加\"软卷皮封面\"文件模板(.xls或.xlsx)                   |")
		print("|                                                                      |")
		print("========================================================================")
		self.RootPath = RootPath+"\\"
		self.bWithTime = True
		try:
			self.filesOrDirsInRoot = os.listdir(self.RootPath)
		except Exception:
			print("No such path: "+self.RootPath)
		else:
			#print("扫描 "+ self.RootPath)
			if os.path.exists(self.RootPath+"卷内文件目录.doc"):
				pass
			else:
				CDMF.print_red_text("在 "+ self.RootPath + " 没有找到\"卷内文件目录.doc\"")
				CDMF.print_red_text("程序中断，请完善相应资料！")
				quit = input("按任意键退出...")
				self.Status = False
				return
			# 询问是否全部重新计算
			while True:
				content = CDMF.print_green_text("是否需要新建或重新生成所有目录? 请输入y/Y或者n/N:")
				if content=="y" or content=="Y" or content=="n" or content=="N":
					break
			if content=="y" or content=="Y":
				self.bRegenerate = True
			else:
				self.bRegenerate = False


			# 询问是否需要将“软卷皮封面.doc”考入个人目录
			while True:
				content = CDMF.print_green_text("是否需要在个人目录自动填充\"软卷皮封面.doc\"? 请输入y/Y或者n/N:")
				if content=="y" or content=="Y" or content=="n" or content=="N":
					break
			if content=="y" or content=="Y":
				self.CopyFengmian = True
				if not os.path.exists(self.RootPath+"软卷皮封面.doc"):
					CDMF.print_red_text("在 "+ self.RootPath + " 没有找到\"软卷皮封面.doc\"")
					while True:
						content1 = CDMF.print_green_text("是否继续? 请输入y/Y或者n/N:")
						if content1=="y" or content1=="Y" or content1=="n" or content=="N":
							break
					if content1=="y" or content1=="Y":
						self.CopyFengmian = False
						pass
					else:
						CDMF.print_red_text("程序中断，请完善相应资料！")
						quit = input("按任意键退出...")
						self.Status = False
						return
			else:
				self.CopyFengmian = False

			if "地区代码及时间表.xlsx" in self.filesOrDirsInRoot:
				CDMF.print_blue_text("正在读取 \"地区代码及时间表.xlsx\"...")
				Excel = easyExcel(self.RootPath+"地区代码及时间表.xlsx")
				try:
					self.AreaTimeAdict = Excel.read_areacode_time()
				except Exception as e:
					CDMF.print_red_text("读取 \"地区代码及时间表.xlsx\" 出错！请检查该文件是否符合模板要求！")
				else:
					pass
				finally:
					pass
				Excel.Close()
			elif "地区代码及时间表.xls" in self.filesOrDirsInRoot:
				CDMF.print_blue_text("正在读取 \"地区代码及时间表.xls\"...")
				Excel = easyExcel(self.RootPath+"地区代码及时间表.xls")
				try:
					self.AreaTimeAdict = Excel.read_areacode_time()
				except Exception as e:
					CDMF.print_red_text("读取 \"地区代码及时间表.xlsx\" 出错！请检查该文件是否符合模板要求！")
				else:
					pass
				finally:
					pass
				Excel.Close()
			else:
				CDMF.print_red_text("在 "+self.RootPath+" 没有找到\"地区代码及时间表.xlsx\"或\"地区代码及时间表.xls\",无法完成目录表中时间自动填充!\"")

				while True:
					content = CDMF.print_green_text("是否继续? 请输入y/Y或者n/N:")
					if content=="y" or content=="Y" or content=="n" or content=="N":
						break
				if content=="y" or content=="Y":
					self.bWithTime = False
					pass
				else:
					CDMF.print_red_text("程序中断，请完善相应资料！")
					quit = input("按任意键退出...")
					self.Status = False
					return
			CDMF.print_blue_text("扫描待统计村民资料...,")
			self.nNumFile = 0;
			self.nNumNoContent = 0;
			for fileOrDir in self.filesOrDirsInRoot:
				if os.path.isdir(fileOrDir) and fileOrDir.startswith(('1','2','3','4','5','6','7','8','9','0')):
					self.nNumFile = self.nNumFile + 1
					if not os.path.exists(self.RootPath + fileOrDir + "\\" +"卷内文件目录.doc"): #扫描没有目录的
						self.nNumNoContent  = self.nNumNoContent + 1

			CDMF.print_blue_text("扫描完毕！共有 "+str(self.nNumFile) + " 户的资料,")
			if self.nNumFile==0:
				quit = input("按任意键退出...")
				self.Status = False
				return
			if self.bRegenerate:
				CDMF.print_blue_text("需要统计的有 "+str(self.nNumFile) + " 户.")
				self.nTotal = self.nNumFile
			else:
				CDMF.print_blue_text("需要统计的有 "+str(self.nNumNoContent) + " 户.")
				self.nTotal = self.nNumFile
				if self.nNumNoContent==0:
					CDMF.print_blue_text("已经没有需要统计的村民了.")
					quit = input("按任意键退出...")
					self.Status = False
					return
			self.Status = True

	def run(self,pros):
		f1 = open('全部操作.txt','w')
		f2 = open('操作成功.txt','w')
		f3 = open('操作失败.txt','w')
		CDMF.print_yellow_text("------------------------------ 开始统计 ----------------------------------")
		# CDMF.print_yellow_text("------------------------------------------------")
		# CDMF.print_yellow_text(" 序号        户主编号与名字           操作状态")
		# CDMF.print_yellow_text("------------------------------------------------")

		#多进程
		try:
			multiP = Pool(pros)
			bar = ProgressBar(total=self.nTotal,width=80)
			for fileOrDir in self.filesOrDirsInRoot:
				if os.path.isdir(fileOrDir) and fileOrDir.startswith(('1','2','3','4','5','6','7','8','9','0')):
					multiP.apply_async(Tasks, args=(fileOrDir,self.RootPath,self.AreaTimeAdict,self.bRegenerate,self.bWithTime,self.CopyFengmian,bar))
			multiP.close()
			multiP.join()
			self.Status = True
		except Exception as e:
			CDMF.print_red_text("运行出现错误！")
			self.Status = False
			raise 
		else:
			pass
		finally:
			pass
		if self.Status:
			CDMF.print_yellow_text("------------------------------ 统计完毕 ----------------------------------")
			f1.close()
			f2.close()
			f3.close()
			CDMF.print_blue_text("任务完成，可查看生成的.txt日志.")
			quit = input("按任意键退出...")
		else:
			CDMF.print_red_text("运行出现错误！")
			self.Status = False
			raise 

if __name__ == '__main__':
	rootpath = os.getcwd()
	Jobb = Job(rootpath)
	if(Jobb.Status): Jobb.run(2)
	