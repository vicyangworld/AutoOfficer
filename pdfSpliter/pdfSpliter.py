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

class PDFspliter(object):
	"""docstring for ClassName"""
	def __init__(self, ROOTPATH):
		self.__RootPath = ROOTPATH;
		self.resPath = '其他资料/承包方/'
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

		if not self.__PathOfInputFiles.endswith('\\'):
			self.__PathOfInputFiles += "\\"

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
		if not self.__mkdir(self.resPath) and len(os.listdir(self.resPath))!=0:
			bRegenerate = self.__quiry("是否需要重新生成？ 请输入y/Y或者n/N: ")
			if bRegenerate:
				try:
					shutil.rmtree(self.resPath)
				except Exception as e:
					print("删除旧文件失败，请查看合并文件是否被占用并解除占用！")
					quit = input("按任意键退出并重新启动...")
					sys.exit(1)
				self.__mkdir(self.resPath)
		allPdfFiles = os.listdir(self.__PathOfInputFiles)
		CDMF.print_blue_text('正在扫描...','')
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
				index +=1
				print('正在处理 '+str(index)+"/"+str(count)+'    '+file, end='')
				starttime1 = datetime.datetime.now()
				try:
					self.pdfReader = PdfFileReader(open(self.__PathOfInputFiles+'\\'+file,'rb'))
				except Exception as e:
					print('打开 '+ self.__PathOfInputFiles+'\\'+file + '失败！')
					quit = input("按任意键退出...")
					sys.exit(1)
				
				# 获取承包方代码以及地块代码（部分）
				AllPages = self.pdfReader.numPages
				print('   '+str(AllPages) + ' 页')
				bFlag = False
				Code_DK = []
				CountTemp = 1
				TotalDK = 0
				for pageNum in range(AllPages):
					if pageNum<AllPages-4:
						continue
					digitals=""
					txt = self.__getPdfTxtAt(pageNum,True)
					#print(txt)
					if "农村土地" in txt or "归户表" in txt or "表6" in txt:
						CountTemp = 1
						Code_DK = []
						bFlag = True
						continue
					if bFlag:
						digitals = re.findall(r'\d+', txt)
						# print(digitals)
						try:
							if len(digitals[0])<3:
								TotalDK += int(digitals[0])
							for digital in digitals:
								if digital.startswith(('0','1','2','3','4','5','6','7','8','9')):
									if len(digital)==5:
										Code_DK.append(digital)
						except Exception as e:
							TotalDK = len(Code_DK)+1
						# print(Code_DK)
						CountTemp += 1

				if len(Code_DK)==0:
					print("归户表格式错误！归户表设置最多为4页")
				if TotalDK != len(Code_DK):
					CDMF.print_red_text("      有地块代码未识别成功！！请查看“操作结果.txt”")
					log("失败：有地块代码未识别失败："+file+"  成功识别："+ "/".join(Code_DK))
					# print(TotalDK)
					# print(Code_DK)
				# # 从界址点成果表中读取地块代码以及承包方代码
				CurrentPage = AllPages-2*CountTemp-1
				Pre_Code_DK=""
				LoopNum = 0
				while True:
					LoopNum += 1
					txt = self.__getPdfTxtAt(CurrentPage,False)
					# print(txt)
					Code_DK_FULL=""
					Code_CBF = ""
					if "界址点成果表" in txt or "界址点坐标" in txt or "界址点编号" in txt:
						# print(txt)
						digitals = re.findall(r'\d+', txt)
						# print(digitals)
						# Code_DK_FULL = digitals[0]
						# Code_CBF = digitals[1]
						# print(digitals)
						for digital in digitals:
							if digital.startswith('1'):
								if len(digital)==19:
									Code_DK_FULL = digital
								elif len(digital)==18:
									Code_CBF = digital
						break
					CurrentPage += 1
					if LoopNum==2:
						break
				if Code_DK_FULL=="" and Code_CBF=="":
					CDMF.print_red_text(file+" 中界址点成果表地块代码与承包方代码均未正常识别,将跳过此次处理")
					log("失败：界址点成果表地块代码与承包方代码均未正常识别："+file)
					continue
				if Code_DK_FULL!="" and Code_DK_FULL.startswith('140'):
					Pre_Code_DK = Code_DK_FULL[0:14]
				if Code_CBF!="" and Code_CBF.startswith('140'):
					Pre_Code_DK = Code_CBF[0:14]
				if not Code_CBF.startswith('140'):
					Code_CBF = Pre_Code_DK+Code_CBF[-4:]
				if len(Pre_Code_DK)!=14 or len(Code_CBF)!=18:
					CDMF.print_red_text('      承包方代码识别失败,将跳过此次处理')
					log("失败：承包方代码识别失败："+file)
					continue
				print('      地块代码以及承包方代码识别成功')
				# print('编码 '+str(Pre_Code_DK)+' '+str(Code_DK)+' '+str(Code_CBF))

				#-------  分解PDF--------------
				#（1）承包方调查表（表2）
				bCBFDCB = False
				txt = self.__getPdfTxtAt(0,False)
				if "表2" in txt or '调查表' in txt:
					self.__writeToPdf(self.resPath+"CBFDCB"+Code_CBF+".pdf",0,0)
					bCBFDCB = True
					print('      成功生成承包方调查表（表2）'+' page: 1-1')
				#(2)分离归户表(表6)
				self.__writeToPdf(self.resPath+"CBJYQGH"+Code_CBF+".pdf",AllPages-CountTemp,AllPages-1)
				print('      成功生成归户表(表6)'+' page: '+str(AllPages-CountTemp+1)+'-'+str(AllPages))
				#(3)分离核实表 (表4)
				if LoopNum==1:
					self.__writeToPdf(self.resPath+"CBJYQDCHS"+Code_CBF+".pdf",AllPages-2*CountTemp,AllPages-CountTemp-1)
					print('      成功生成核实表 (表4)'+' page: '+str(AllPages-2*CountTemp+1)+'-'+str(AllPages-CountTemp))
				else:
					self.__writeToPdf(self.resPath+"CBJYQDCHS"+Code_CBF+".pdf",AllPages-2*CountTemp+1,AllPages-CountTemp-1)
					print('      成功生成核实表 (表4)'+' page: '+str(AllPages-2*CountTemp+1+1)+'-'+str(AllPages-CountTemp))			
				#(4) 地块调查表
				for ii in range(0,len(Code_DK)):
					temp= AllPages-2*CountTemp+(LoopNum-1)-3*len(Code_DK) + ii*3
					if ii==0:
						p_beg = temp+1
					if ii==len(Code_DK)-1:
						p_end = temp+3
					self.__writeToPdf(self.resPath+"CBFDKDCB"+Pre_Code_DK+Code_DK[ii]+".pdf",temp,temp+2)
				print('      成功生成地块调查表'+' page: '+str(p_beg+1)+'-'+str(p_end))
				#---------------------------------------------
				#(5) 承包方代表身份证及证明
				#检查是否有证明，在第3页
				if bCBFDCB:
					pageNum = 2  #第三页，从0开始计算
				else:
					pageNum = 1  #第三页，从0开始计算
				txt = self.__getPdfTxtAt(pageNum,False)
				# print(txt)
				bZM=False
				if "证明" in txt or "兹证明" in txt or "兹证" in txt or "情况属实" in txt or "承诺":
					bZM = True
				if "常住" in txt or "登记" in txt:
					bZM = True
				if "签发" in txt or "机关" in txt or "身份" in txt or "姓名" in txt or "地址" in txt:
					bZM = False
				if bZM:
					self.__writeToPdf(self.resPath+"CBFMC"+Code_CBF+".pdf",pageNum-1,pageNum)
					print('      成功生成承包方身份证明'+' page:'+str(pageNum)+'-'+str(pageNum+1))
				else:
					self.__writeToPdf(self.resPath+"CBFMC"+Code_CBF+".pdf",pageNum-1,pageNum-1)
					print('      成功生成承包方身份证明'+' page: '+str(pageNum)+'-'+str(pageNum))
				#（6）合同
				CurrentPage = AllPages-2*CountTemp-3*len(Code_DK)-1+(LoopNum-1)-1
				LoopNumtemp = 0
				offset = 0
				while True:
					LoopNumtemp += 1
					txt = self.__getPdfTxtAt(CurrentPage,True)
					if "通知书" in txt or "指界" in txt or "领导小组" in txt:
						self.__writeToPdf(self.resPath+"HT"+Code_CBF+".pdf",CurrentPage-2,CurrentPage-1)
						print('      成功生成合同'+' page: '+str(CurrentPage-1)+'-'+str(CurrentPage))
						offset = 3
						break
					if ("一式三份" in txt or "单位各一份" in txt or "另行拍卖" in txt or "合同" in txt 
						or "一" in txt or "二" in txt or "三" in txt or "四" in txt or "五" in txt or "六" in txt or "七" in txt
						or "拍卖" in txt):
						if LoopNumtemp<=2:
							self.__writeToPdf(self.resPath+"HT"+Code_CBF+".pdf",CurrentPage-1,CurrentPage)  #没有指导书的情况或者没有识别
							print('      成功生成合同'+' page: '+str(CurrentPage)+'-'+str(CurrentPage+1))
							offset = 2
						elif LoopNumtemp==3:
							self.__writeToPdf(self.resPath+"HT"+Code_CBF+".pdf",CurrentPage,CurrentPage+1)  #没有指导书的情况或者没有识别
							print('      成功生成合同'+' page: '+str(CurrentPage+1)+'-'+str(CurrentPage+2))
							offset = 1
						break;
					CurrentPage = CurrentPage - 1
					if LoopNumtemp>3:
						break
				#(7)家庭成员
				CurrentPage = AllPages-2*CountTemp-3*len(Code_DK)-1+(LoopNum-1)-offset
				if bZM:
					JTCY_beg = 3
				else:
					JTCY_beg = 2
				self.__writeToPdf(self.resPath+"CBFJTCY"+Code_CBF+".pdf",JTCY_beg,CurrentPage)
				print('      成功生成家庭成员'+' page: '+str(JTCY_beg+1)+'-'+str(CurrentPage+1))
				endtime1 = datetime.datetime.now()
				CDMF.print_blue_text('      用时 '+str(endtime1 - starttime1))

if __name__ == '__main__':
	if not readlisence():
		quit = input("按任意键退出...")
		sys.exit(1)
	time.sleep(1)
	ROOTPATH = os.getcwd()
	CDMF.print_white_text("当前工作目录："+ROOTPATH)
	CDMF.print_yellow_text("正在检测必要组件......")
	time.sleep(1)
	dirs = os.listdir(r'C:\\Program Files (x86)\\')
	bImage = False
	bTesseract = False
	bGs = False
	for x in dirs:
		if 'ImageMagick' in x:
			print(x+'\r',end='')
			time.sleep(1)
			print(x+'  已经安装!')
			bImage = True
		if 'Tesseract' in x:
			print(x+'\r',end='')
			time.sleep(2)
			print(x+'  已经安装!')
			bTesseract = True
		if 'gs' in x and len(x)==2:
			print(x+'\r',end='')
			time.sleep(2)
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
