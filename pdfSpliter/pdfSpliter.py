from wand.image import Image
from PIL import Image as PI
import pyocr
import pyocr.builders
import io,sys,os,shutil
from PyPDF2.pdf import PdfFileReader
from PyPDF2.pdf import PdfFileWriter
import re
import CmdFormat
import datetime,time,socket
from PIL import ImageFilter
import lisence

CDMF = CmdFormat.CmdFormat("PDF分离及识别器")
ISOTIMEFORMAT='%Y-%m-%d %X'

def log(x):
	""" recording the status information"""
	if x==None:
		return
	with open('操作结果.txt', 'a+') as f:
		f.write(str(x)+'\n')

def readlisence():
	CDMF.print_blue_text('验证许可证')
	s1 = lisence.get_mac_address()
	if os.path.exists('lisence.lis'):
		try:
			with open('lisence.lis', 'r') as f:
				content = f.read()
				s2 = lisence.decrypt('cxr', content)
		except Exception as e:
			CDMF.print_red_text("验证码不正确，请联系管理员：QQ:529301432 手机：18801415145")
			return False
		if s1.lower() == s2.lower():
			CDMF.print_blue_text("验证成功,欢迎使用！")
			return True
		else:
			CDMF.print_red_text("验证码不正确,请联系管理员：QQ:529301432 手机：18801415145！")
			return False
	else:
		CDMF.print_red_text("没有许可文件lisence.lis，请联系管理员：QQ:529301432 手机：18801415145")
		return False


class PDFspliter(object):
	"""docstring for ClassName"""
	def __init__(self, ROOTPATH):
		self.__RootPath = ROOTPATH;
		self.resPath = ROOTPATH+'/其他资料/承包方/'
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
			print("==========================  欢迎使用  ==================================")
			print("|                                                                      |")
			print("|      将本程序放在根目录，运行之前请确保根目录下具有                  |")
			CDMF.print_red_text("|      (1) *每户PDF文件                                                |")
			print("|      (2) 生成的文件放在了/其他资料/承包方/                           |")
			CDMF.print_red_text("|      (3) *注意，扫描质量不同，识别有可能失败 ！                      |")
			print("|                                                                      |")
			print("========================================================================")
	def Run(self):
		if not readlisence():
			sys.exit(1)
		timeString = time.strftime( ISOTIMEFORMAT, time.localtime(time.time()))
		hostName = socket.gethostname()
		log('-----------Log Time: '+timeString+ ' from '+hostName+'  --------------')
		self.__messages()
		tools = pyocr.get_available_tools()[:]
		if len(tools)==0:
			print("No ocr tool found")
			sys.exit(1)
		# else:
		# 	print("Using '%s' " % (tools[0].get_name()))
		self.__tool = tools[0]
		self.__lang  = self.__tool.get_available_languages()[0]  #中文

		if not self.__mkdir(self.resPath) and len(os.listdir(self.resPath))!=0:
			bRegenerate = self.__quiry("是否需要重新生成？ 请输入y/Y或者n/N: ")
			if bRegenerate:
				try:
					shutil.rmtree(self.resPath)
					self.__mkdir(self.resPath)
				except Exception as e:
					print("删除旧文件失败，请查看合并文件是否被占用！")
					sys.exit(1)
		allPdfFiles = os.listdir(self.__RootPath)
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
				print('正在处理 '+str(index)+"/"+str(count))
				starttime1 = datetime.datetime.now()
				self.pdfReader = PdfFileReader(open(self.__RootPath+'\\'+file,'rb'))
				# 获取承包方代码以及地块代码（部分）
				AllPages = self.pdfReader.numPages
				bFlag = False
				Code_DK = []
				CountTemp = 1
				def getPdfTxtAt(pageNum,bENHANCE):
					# print('---->>>>>'+str(pageNum))
					RESOLUTION = 200
					tempoutPdfName = 'temp.pdf'
					pdfWriter = PdfFileWriter()     #生成一个空白的pdf文件
					pdfWriter.addPage(self.pdfReader.getPage(pageNum))
					pdfOutput = open('./'+tempoutPdfName,'wb')
					pdfWriter.write(pdfOutput)                           #将复制的内容全部写入合并的pdf
					pdfOutput.close()
					# print('--->'+'open pdf')
					with Image(filename='./'+tempoutPdfName,resolution=RESOLUTION) as image_pdf:
					# print('--->'+'convert pdf to jpeg')
						image_jpeg = image_pdf.convert('jpeg')
					img_page = Image(image=image_jpeg)
					req_image = img_page.make_blob('jpeg')

					# print('--->'+'recognite')
					image_filtered = PI.open(io.BytesIO(req_image))
					# image_filtered= image_filtered.filter(ImageFilter.GaussianBlur(radius=1))
					# if bENHANCE:
					# 	image_filtered= image_filtered.filter(ImageFilter.EDGE_ENHANCE)
					txt = self.__tool.image_to_string(
						image_filtered,
						lang=self.__lang,
						builder=pyocr.builders.TextBuilder()
					)
					os.remove('./'+tempoutPdfName)
					return txt

				def writeToPdf(filename,beg,end):
					pdfWriter = PdfFileWriter()     #生成一个空白的pdf文件
					for x in range(beg,end+1):
						pdfWriter.addPage(self.pdfReader.getPage(x))
					pdfOutput = open(filename,'wb')
					pdfWriter.write(pdfOutput)                         #将复制的内容全部写入合并的pdf
					pdfOutput.close()
				TotalDK = 0
				for pageNum in range(AllPages):
					if pageNum<AllPages-4:
						continue
					digitals=""
					txt = getPdfTxtAt(pageNum,True)
					# print(txt)
					if "农村土地" in txt or "归户表" in txt or "表6" in txt:
						CountTemp = 1
						Code_DK = []
						bFlag = True
						continue
					if bFlag:
						digitals = re.findall(r'\d+', txt)
						# print('DDDDD  ','')
						# print(digitals)
						if len(digitals[0])<3:
							TotalDK += int(digitals[0])
						for digital in digitals:
							if digital.startswith(('0','1','2','3','4','5','6','7','8','9')):
								if len(digital)==5:
									Code_DK.append(digital)
						# print(Code_DK)
						CountTemp += 1
				if len(Code_DK)==0:
					print("归户表格式错误！归户表设置最多为4页")
				if TotalDK != len(Code_DK):
					CDMF.print_red_text("有地块代码未识别成功！！")
					log("失败：有地块代码未识别失败："+file+"  成功识别："+ "/".join(Code_DK))
					print(TotalDK)
					print(Code_DK)
				txt = getPdfTxtAt(AllPages-2*CountTemp-1,False)
				# print(txt)
				Pre_Code_DK=""
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
				print('编码 '+str(Pre_Code_DK)+' '+str(Code_DK)+' '+str(Code_CBF))

				#-------  分解PDF--------------
				#（1）承包方调查表（表2）
				writeToPdf(self.resPath+"CBFDCB"+Code_CBF+".pdf",0,0)
				print('      成功生成承包方调查表（表2）')
				#(2)分离归户表(表6)
				writeToPdf(self.resPath+"CBJYQGH"+Code_CBF+".pdf",AllPages-CountTemp,AllPages-1)
				print('      成功生成归户表(表6)')
				#(3)分离核实表 (表4)
				writeToPdf(self.resPath+"CBJYQDCHS"+Code_CBF+".pdf",AllPages-2*CountTemp,AllPages-CountTemp-1)
				print('      成功生成核实表 (表4)')
				#(4) 地块调查表
				for ii in range(0,len(Code_DK)):
					temp= AllPages-2*CountTemp-3*len(Code_DK) + ii*3
					writeToPdf(self.resPath+"CBFDKDCB"+Pre_Code_DK+Code_DK[ii]+".pdf",temp,temp+2)
				print('      成功生成地块调查表')
				#---------------------------------------------
				#(5) 承包方代表身份证及证明
				#检查是否有证明，在第3页
				pageNum = 2  #第三页，从0开始计算
				txt = getPdfTxtAt(pageNum,False)
				# print(txt)
				bZM=False
				if "证明" in txt or "兹证明" in txt or "情况属实" in txt:
					bZM = True
				if bZM:
					writeToPdf(self.resPath+"CBFMC"+Code_CBF+".pdf",1,2)
				else:
					writeToPdf(self.resPath+"CBFMC"+Code_CBF+".pdf",1,1)
				print('      成功生成承包方身份证明')
				#（6）合同
				CurrentPage = AllPages-2*CountTemp-3*len(Code_DK)-1
				while True:
					txt = getPdfTxtAt(CurrentPage,True)
					if "一式三份" in txt or "单位各一份" in txt or "另行拍卖" in txt:
						break;
					CurrentPage = CurrentPage - 1
				writeToPdf(self.resPath+"HT"+Code_CBF+".pdf",CurrentPage,CurrentPage+1)
				print('      成功生成合同')

				#(7)家庭成员
				if bZM:
					JTCY_beg = 3
				else:
					JTCY_beg = 2
				writeToPdf(self.resPath+"CBFJTCY"+Code_CBF+".pdf",JTCY_beg,CurrentPage)
				print('      成功生成家庭成员')
				endtime1 = datetime.datetime.now()
				CDMF.print_blue_text('      用时 '+str(endtime1 - starttime1))

if __name__ == '__main__':
	starttime = datetime.datetime.now()
	ROOTPATH = os.getcwd()
	Job = PDFspliter(ROOTPATH)
	Job.Run()
	CDMF.print_yellow_text("任务完成！")
	endtime = datetime.datetime.now()
	print('总共用时 '+str(endtime - starttime))
	quit = input("按任意键退出...")