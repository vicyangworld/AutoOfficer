from wand.image import Image
import pyocr
import pyocr.builders
import io,sys,os,shutil
from PyPDF2.pdf import PdfFileReader
from PyPDF2.pdf import PdfFileWriter
import re
import CmdFormat
import datetime,time,socket
from PIL import ImageFilter
from PIL import Image as PI
import lisence
import win32com
from win32com.client import Dispatch, constants

VERSION = 'V1.1'
CDMF = CmdFormat.CmdFormat("PDF分离及识别器"+VERSION+"(试用版)")
ISOTIMEFORMAT='%Y-%m-%d %X'

def log(x):
	""" recording the status information"""
	if x==None:
		return
	with open('操作结果.txt', 'a+') as f:
		f.write(str(x)+'\n')

def readlisence():
	CDMF.print_blue_text('验证许可证....')
	time.sleep(2)
	addresses = lisence.get_mac_address()
	#CDMF.print_blue_text('本机物理地址为：'+s1.lower())
	if os.path.exists('lisence.lis'):
		try:
			with open('lisence.lis', 'r') as f:
				content = f.read()
				s2 = lisence.decrypt('cxr', content)
		except Exception as e:
			CDMF.print_red_text("验证码不正确，请联系管理员：QQ:35272212 手机：15934000850")
			return False
		if s2.lower() in addresses:
			#CDMF.print_blue_text('许可文件物理地址：'+s2.lower())
			CDMF.print_blue_text("验证成功,欢迎使用！")
			return True
		else:
			#CDMF.print_blue_text('许可文件物理地址：'+s2.lower())
			CDMF.print_red_text("验证失败,请联系管理员：QQ:35272212 手机：15934000850！")
			return False
	else:
		CDMF.print_red_text("没有许可文件lisence.lis，请联系管理员：QQ:35272212 手机：15934000850")
		return False

class easyExcel(object):
	"""docstring for easyExcel"""
	def __init__(self, FileName):
		self.ExcelApp = win32com.client.DispatchEx('Excel.Application')
		self.ExcelApp.Visible=False
		self.ExcelApp.DisplayAlerts = False
		self.DKBM = []
		self.CBFBM = []
		if FileName:
			self.FileName = FileName
			self.Xls = self.ExcelApp.Workbooks.Open(self.FileName)
		else:
			myLog("文件 "+FileName+" 没找到！")
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
			Activesheet = self.Xls.Worksheets(x)
			log("提取 "+Activesheet.Name +" 信息...")
			# UsedRange 从1开始
			nTempCode = Activesheet.Cells(1,4).Value
			nRows = Activesheet.UsedRange.Rows.Count
			nColumns = Activesheet.UsedRange.Columns.Count
			self.DKBM = []
			self.CBFBM = []
			for i in range(2,nRows+1):
				for j in range(3,5):
					string = Activesheet.Cells(i,j).Value
					if string!=None and j==3:
						self.DKBM.append(string)
					if string!=None and j==4:
						self.CBFBM.append(string)

class PDFspliter(object):
	"""docstring for ClassName"""
	def __init__(self, ROOTPATH):
		self.__RootPath = ROOTPATH;
		self.resPath = '其他资料\\承包方\\'
		self.DKBM = []
		self.CBFBM = []
	def __quiry(self,mes):
		global CDMF
		while True:
			content = CDMF.print_green_input_text(mes)
			if content=="y" or content=="Y" or content=="n" or content=="N":
				break
		if content=="y" or content=="Y":
			return True
		else:
			return False
	def __mkdir(self,path):
		path=path.strip()
		path=path.rstrip("\\")
		isExists=os.path.exists(path)
		if not isExists:
			os.makedirs(path)
			return True
		else:
			return False
	def __messages(self):
			CDMF.set_cmd_color(CmdFormat.FOREGROUND_RED | CmdFormat.FOREGROUND_GREEN | \
				CmdFormat.FOREGROUND_BLUE | CmdFormat.FOREGROUND_INTENSITY)
			print("\n")
			print("===================  PDF分离及识别器"+VERSION+"(试用版)  ======================")
			print("|                                                                      |")
			print("|      将本程序放在根目录，运行之前请确保根目录下具有                  |")
			CDMF.print_red_text("|      (1) *每户PDF文件                                                |")
			print("|      (2) 生成的文件放在了/其他资料/承包方/                           |")
			CDMF.print_red_text("|      (3) *注意，扫描质量不同，识别有可能失败 ！                      |")
			print("|                                                                      |")
			print("========================================================================")
	def __getPdfTxtAt(self,pageNum,bENHANCE):
		# print('---->>>>>'+str(pageNum))
		try:
			RESOLUTION = 250
			tempoutPdfName = 'temp.pdf'
			tempoutPdfNameWithAbsPath = os.path.join(self.__RootPath,tempoutPdfName)
			if os.path.exists(tempoutPdfNameWithAbsPath):
				os.remove(tempoutPdfNameWithAbsPath)
			pdfWriter = PdfFileWriter()     #生成一个空白的pdf文件
			pdfWriter.addPage(self.pdfReader.getPage(pageNum))
			with open(tempoutPdfNameWithAbsPath,'wb') as pdfOutput:
				pdfWriter.write(pdfOutput)                           #将复制的内容全部写入合并的pdf
			try:
				with Image(filename=tempoutPdfNameWithAbsPath,resolution=RESOLUTION) as image_pdf:
					image_jpeg = image_pdf.convert('jpeg')
			except Exception as e:
				raise e
				raise(r'Image(filename=tempoutPdfNameWithAbsPath,resolution=RESOLUTION) occurs error!' )
				quit = input("按任意键退出...")
				sys.exit(1)
			try:
				# img_page = Image(image=image_jpeg)
				req_image = image_jpeg.make_blob('jpeg')
			except Exception as e:
				raise e
				print('make_blob or ERROR!   '+ str(pageNum)+' 页失败！')
				quit = input("按任意键退出...")
				sys.exit(1)
			try:
				image_filtered = PI.open(io.BytesIO(req_image))
				# image_filtered= image_filtered.filter(ImageFilter.GaussianBlur(radius=1))
				# if bENHANCE:
				# 	image_filtered= image_filtered.filter(ImageFilter.EDGE_ENHANCE)
			except Exception as e:
				raise e
				print('PI.open ERROR!   '+ str(pageNum)+' 页失败！')
				quit = input("按任意键退出...")
				sys.exit(1)
			try:
				# print('>>> Debug:'+self.__lang)
				txt = self.__tool.image_to_string(
					image_filtered,
					lang=self.__lang,
					builder=pyocr.builders.TextBuilder()
				)
			except Exception as e:
				raise e
				print('image_to_string   '+ str(pageNum)+' 页失败！')
				quit = input("按任意键退出...")
				sys.exit(1)
			if os.path.exists(tempoutPdfNameWithAbsPath):
				os.remove(tempoutPdfNameWithAbsPath)
			return txt
		except Exception as e:
			raise e
			print('获取第 '+ str(pageNum)+' 页失败！')
			quit = input("按任意键退出...")
			sys.exit(1)

	def __writeToPdf(self,filename,beg,end):
		pdfWriter = PdfFileWriter()     #生成一个空白的pdf文件
		for x in range(beg,end+1):
			pdfWriter.addPage(self.pdfReader.getPage(x))
		with open(filename,'wb') as pdfOutput:              #将复制的内容全部写入合并的pdf
			pdfWriter.write(pdfOutput)

	def Run(self):
		timeString = time.strftime( ISOTIMEFORMAT, time.localtime(time.time()))
		hostName = socket.gethostname()
		log('-----------Log Time: '+timeString+ ' from '+hostName+'  --------------')
		self.__messages()
		#输入待处理文件绝对路径
		self.__PathOfInputFiles = CDMF.print_green_input_text('请输入待处理文件绝对路径:')
		while True:
			if self.__PathOfInputFiles=="":
				self.__PathOfInputFiles = CDMF.print_green_input_text('请输入待处理文件绝对路径:')
			else:
				if not os.path.exists(self.__PathOfInputFiles):
					CDMF.print_red_text("路径不存在: "+ self.__PathOfInputFiles)
					self.__PathOfInputFiles = CDMF.print_green_input_text('请重新输入待处理文件绝对路径:')
				else:
					break
		tools = pyocr.get_available_tools()[:]
		if len(tools)==0:
			print("No ocr tool found")
			sys.exit(1)
		else:
			print("Using '%s' " % (tools[0].get_name()))
		self.__tool = tools[0]
		templist = self.__tool.get_available_languages()
		# print('>>> Debug:')
		# print(templist)

		self.__lang = ""
		for x in templist:
			if 'chi' in x:
				self.__lang  = x #self.__tool.get_available_languages()[0]  #中文
		# print('>>> Debug:'+self.__lang)
		if self.__lang=="":
			print('在Tesseract-OCR安装目录中缺失OCR中文包 chi_sim.traineddata！！')
			quit = input("按任意键退出...")
			sys.exit(1)

		self.resPath = os.path.join(self.__PathOfInputFiles,self.resPath)
		if os.path.exists(self.resPath):
			bRegenerate = self.__quiry("是否需要重新生成？ 请输入y/Y或者n/N: ")
			if bRegenerate:
				try:
					shutil.rmtree(self.resPath)
				except Exception as e:
					print("删除旧文件失败，请查看合并文件是否被占用并解除占用！")
					quit = input("按任意键退出并重新启动...")
					sys.exit(1)
			os.makedirs(self.resPath)
		else:
			os.makedirs(self.resPath)

		allPdfFiles = os.listdir(self.__PathOfInputFiles)

		for x in allPdfFiles:
			if ("地块属性.xlsx" in x or "地块属性.xls" in x ) and not x.startswith('~'):
				self.DKSXFile = x
				print("正在提取 " + os.path.join(self.__PathOfInputFiles,x) + "信息")
				Excel = easyExcel(os.path.join(self.__PathOfInputFiles,x))
				Excel.read_areacode_time()
				self.DKBM = Excel.DKBM
				self.CBFBM = Excel.CBFBM
				Excel.Close()
				print("提取成功！")
		for x in self.DKBM:
			if x:
				PRE_CODE = x[0:14]
		CDMF.print_blue_text('正在扫描待处理文件...','')
		count = 0
		for file in allPdfFiles:
			(fname,extension) = os.path.splitext(file)
			if file.endswith('.pdf') and fname.isdigit():
				count +=1
		CDMF.print_blue_text(' 共有 '+str(count)+' 个文件需要处理')
		index = 0
		for file in allPdfFiles:
			(fname,extension) = os.path.splitext(file)
			if file.endswith('.pdf') and fname.isdigit():
				CBF = fname
				index +=1
				CDMF.print_yellow_text('正在处理 '+str(index)+"/"+str(count)+'    '+file, end='')
				starttime1 = datetime.datetime.now()
				try:
					self.pdfReader = PdfFileReader(open(self.__PathOfInputFiles+'\\'+file,'rb'))
				except Exception as e:
					print('打开 '+ os.path.join(self.__PathOfInputFiles,file)+ ' 失败！')
					quit = input("按任意键退出...")
					sys.exit(1)

				# 获取承包方代码以及地块代码（部分）
				AllPages = self.pdfReader.numPages
				CDMF.print_yellow_text('   共'+str(AllPages) + ' 页')
				curDK = []
				bRead = False
				for x in range(0,len(self.CBFBM)):
					if self.CBFBM[x]==CBF:
						try:
							curDK.append(self.DKBM[x][-5:])
							bRead = True
						except Exception as e:
							CDMF.pring_red_text(file + "在" + self.DKSXFile + " 编码未读取到! 将跳过此次处理" )
							bRead = False
				if bRead==False:
					continue

				CurrentPage = AllPages-1
				Page_GH = AllPages
				nHT = 0
				while True:
					CurrentPage -= 1
					print('   识别 '+str(CurrentPage+1)+'/'+str(AllPages)+'...')
					txt = self.__getPdfTxtAt(CurrentPage,False)
					#print(txt)
					if "归户表" in txt or "表6" in txt:
						self.__writeToPdf(self.resPath+"CBJYQGH"+PRE_CODE+CBF+".pdf",CurrentPage,AllPages-1)
						Page_GH = CurrentPage
						print('      成功生成归户表(表6)'+' page: '+str(CurrentPage+1)+'-'+str(AllPages))
						nJump = AllPages-2-CurrentPage
						CurrentPage = CurrentPage-nJump+1
					if "核实表" in txt or "表4" in txt:
						self.__writeToPdf(self.resPath+"CBJYQHS"+PRE_CODE+CBF+".pdf",CurrentPage,Page_GH-1)
						Page_HS = CurrentPage
						print('      成功生成核实(表4)'+' page: '+str(CurrentPage+1)+'-'+str(Page_GH))
					if "界址点成果表" in txt or "界址点坐标" in txt or "界址点编号" in txt:
						bRec = False
						digitals = re.findall(r'\d+', txt)
						digital_list=[]
						for digital in digitals:
							if digital.startswith('1'):
								if len(digital)>=14:
									digital_list.append(digital)
									bRec = True
						if bRec:
							DK = digital_list[0][-5:]
						if not bRec:
							digital_list=[]
							txt = self.__getPdfTxtAt(CurrentPage-1,True)
							digitals = re.findall(r'\d+', txt)
							if len(digitals[0])>=14:
									digital_list.append(digital)
							if len(digital_list)>=1:
								bRec = True
								DK = digital_list[0][-5:]
						if bRec:
							if DK in curDK:
								self.__writeToPdf(self.resPath+"CBFDKDCB"+PRE_CODE+DK+".pdf",CurrentPage-2,CurrentPage)
								print('      '+DK+'成功生成地块调查表'+PRE_CODE+DK+' page: '+str(CurrentPage-1)+'-'+str(CurrentPage+1))
							else:
								self.__writeToPdf(self.resPath+"CBFDKDCB"+PRE_CODE+DK+"_XXX.pdf",CurrentPage-2,CurrentPage)
								CDMF.print_red_text('      '+DK+'地块代码可能识别失败，但已保存为'+PRE_CODE+DK+"_XXX.pdf"+'  page: '+str(CurrentPage-2)+'-'+str(CurrentPage)+ '  请检查该文件并重命名！')
								log('      '+DK+'地块代码可能识别失败但已保存为'+PRE_CODE+DK+"_XXX.pdf"+' page: '+str(CurrentPage-1)+'-'+str(CurrentPage+1) +"   文件"+file+ '  请检查该文件并重命名！')
						else:
							self.__writeToPdf(self.resPath+"CBFDKDCB"+PRE_CODE+'XXXXX'+"_XXX.pdf",CurrentPage-2,CurrentPage)
							print('      有地块代码未识别成功，已经保存到为'+PRE_CODE+'XXXXX'+"_XXX.pdf"+'   请检查并重命名！')
						CurrentPage -=2

					bHT = False
					if "一式三份" in txt or "单位各一份" in txt or "另行拍卖" in txt or ("拍卖" in txt and "归户表" not in txt) or "四荒" in txt or "一百年" in txt or "使用期" in txt \
						or "鼓励" in txt or "明细" in txt or "承包期" in txt or "收回" \
						in txt or "另行发包" in txt or "上交国家" in txt or "另行发惩" in txt or "基础设施" in txt:

						bHT = True
						nHT += 1
						if nHT==1:
							end = CurrentPage
					else:
						bHT = False
					if not bHT and nHT >1:
						self.__writeToPdf(self.resPath+"HT"+PRE_CODE+CBF+".pdf",CurrentPage+1,end)
						print('      成功生成合同'+' page: '+str(CurrentPage+2)+'-'+str(end+1))
						break

				bCBFDCB = False
				txt = self.__getPdfTxtAt(0,False)
				if "表2" in txt or '调查表' in txt or '发包方' in txt or '承包方' in txt or '联系电话' in txt:
					self.__writeToPdf(self.resPath+"CBFDCB"+PRE_CODE+CBF+".pdf",0,0)
					bCBFDCB = True
					print('      成功生成承包方调查表（表2）'+' page: 1-1')
				#检查是否有证明，在第3页
				if bCBFDCB:
					pageNum = 2  #第三页，从0开始计算
				else:
					pageNum = 1  #第三页，从0开始计算
				CurrentPagetemp = pageNum
				while True:
					txt = self.__getPdfTxtAt(CurrentPagetemp,False)
					txt.replace("\n", "")
					txt.replace(' ', "")
					print('   识别 '+str(CurrentPagetemp+1)+'/'+str(AllPages)+'...')
					if "证明" in txt or "兹证明" in txt or "兹证" in txt or "情况属实" in txt or "承诺" in txt\
						or "承包方代表" in txt or "委托" in txt or "申请书" in txt\
						or "村委会" in txt or "变更承包" in txt or "意见" in txt or "发包方意见" in txt\
						or "常住" in txt or "登记" in txt or "曾用名" in txt or "籍贯" in txt \
						or "迁来" in txt or "本市" in txt \
						or "家庭成员共同推选" in txt or "受托方" in txt or "何时" in txt:
						print(txt)
						pass
					else:
						break
					CurrentPagetemp += 1

				self.__writeToPdf(self.resPath+"CBFMC"+PRE_CODE+CBF+".pdf",pageNum-1,CurrentPagetemp-1)
				print('      成功生成承包方身份证明'+' page:'+str(pageNum)+'-'+str(CurrentPagetemp))

				self.__writeToPdf(self.resPath+"CBFJTCY"+PRE_CODE+CBF+".pdf",CurrentPagetemp,CurrentPage)
				print('      成功生成家庭成员'+' page: '+str(CurrentPagetemp+1)+'-'+str(CurrentPage+1))

				endtime1 = datetime.datetime.now()
				CDMF.print_blue_text('      用时 '+str(endtime1 - starttime1))

if __name__ == '__main__':
	if not readlisence():
		quit = input("按任意键退出...")
		sys.exit(1)
	#time.sleep(1)
	ROOTPATH = os.getcwd()
	CDMF.print_white_text("当前工作目录："+ROOTPATH)
	CDMF.print_yellow_text("正在检测必要组件......")
	#time.sleep(1)
	dirs = os.listdir(r'C:\\Program Files (x86)\\')
	bImage = False
	bTesseract = False
	bGs = False
	for x in dirs:
		if 'ImageMagick' in x:
			print(x+'\r',end='')
			#time.sleep(1)
			print(x+'  已经安装!')
			bImage = True
		if 'Tesseract' in x:
			print(x+'\r',end='')
			#time.sleep(1)
			print(x+'  已经安装!')
			bTesseract = True
		if 'gs' in x and len(x)==2:
			print(x+'\r',end='')
			#time.sleep(1)
			bGs = True
			print(x+'  已经安装!')

	if not bImage:
		if not os.path.exists('./install_imagemagick.bat'):
			CDMF.print_red_text("当前目录不存 install_imagemagick.bat，请联系软件提供者. ")
			quit = input("按任意键退出...")
			sys.exit(1)
		os.system(ROOTPATH+'\\'+'install_imagemagick.bat')#安装必要的两个软件
	if not bTesseract:
		if not os.path.exists('./install_tesseract.bat'):
			CDMF.print_red_text("当前目录不存 install_tesseract.bat，请联系软件提供者. ")
			quit = input("按任意键退出...")
			sys.exit(1)
		os.system(ROOTPATH+'\\'+'install_tesseract.bat')#安装必要的两个软件
	if not bGs:
		if not os.path.exists('./install_gs.bat'):
			CDMF.print_red_text("当前目录不存 install_gs.bat，请联系软件提供者. ")
			quit = input("按任意键退出...")
			sys.exit(1)
		os.system(ROOTPATH+'\\'+'install_gs.bat')#安装必要的两个软件
	shutil.copyfile(os.path.join(ROOTPATH,'tesseract','tessdata','chi_sim.traineddata'),os.path.join(r'C:\Program Files (x86)\Tesseract-OCR\tessdata','chi_sim.traineddata'))
	starttime = datetime.datetime.now()
	Job = PDFspliter(ROOTPATH)
	Job.Run()
	CDMF.print_yellow_text("任务完成！")
	endtime = datetime.datetime.now()
	print('总共用时 '+str(endtime - starttime))
	quit = input("按任意键退出...")
