from wand.image import Image
from PIL import Image as PI
import pyocr
import pyocr.builders
import io,sys,os,shutil
from PyPDF2.pdf import PdfFileReader
from PyPDF2.pdf import PdfFileWriter
import re
import CmdFormat

CDMF = CmdFormat.CmdFormat("PDF分离及识别器")
RESOLUTION = 150
tools = pyocr.get_available_tools()[:]
if len(tools)==0:
	print("No ocr tool found")
	sys.exit(1)
else:
	print("Using '%s' " % (tools[0].get_name()))
tool = tools[0]
lang  = tool.get_available_languages()[0]  #中文

tempoutPdfName = 'temp.pdf'
pdfReader = PdfFileReader(open('E:\\PythonProjects\\pdfSpliter\\0001.pdf','rb'))


def __quiry(mes):
	while True:
		content = CDMF.print_green_input_text(mes)
		if content=="y" or content=="Y" or content=="n" or content=="N":
			break
	if content=="y" or content=="Y":
		return True
	else:
		return False

def mkdir(path):
	path=path.strip()
	path=path.rstrip("\\")
	isExists=os.path.exists(path)
	if not isExists:
		os.makedirs(path)
		return True
	else:
		return False

resPath = os.getcwd()+'/其他资料/承包方/'
if not mkdir(resPath) and len(os.listdir(resPath))!=0:
	bRegenerate = __quiry("是否需要重新生成？ 请输入y/Y或者n/N: ")
	if bRegenerate:
		try:
			shutil.rmtree(resPath)
			mkdir(resPath)
		except Exception as e:
			print("删除旧合并文件失败，请查看合并文件是否被占用！")
			sys.exit(1)

#表2承包方调查表
Page_CBFDCB_beg = 0;
Page_CBFDCB_end = 0;
#承包方代表身份证及证明
Page_CBFDB_beg = 1;
Page_CBFDB_end = 1;

# 获取承包方代码以及地块代码（部分）
AllPages = pdfReader.numPages
bFlag = False
Code_DK = []
CountTemp = 1

for pageNum in range(AllPages):
	if pageNum<AllPages-4:
		continue
	print('--->'+str(pageNum))
	pdfWriter = PdfFileWriter()     #生成一个空白的pdf文件
	pdfWriter.addPage(pdfReader.getPage(pageNum))
	pdfOutput = open('./'+tempoutPdfName,'wb')
	pdfWriter.write(pdfOutput)                           #将复制的内容全部写入合并的pdf
	pdfOutput.close()
	print('--->'+'open pdf')
	image_pdf = Image(filename='./'+tempoutPdfName,resolution=RESOLUTION)
	print('--->'+'convert pdf to jpeg')
	image_jpeg = image_pdf.convert('jpeg')
	img_page = Image(image=image_jpeg)
	req_image = img_page.make_blob('jpeg')
	print('--->'+'recognite')
	txt = tool.image_to_string(
		PI.open(io.BytesIO(req_image)),
		lang=lang,
		builder=pyocr.builders.TextBuilder()
	)
	os.remove(tempoutPdfName)
	# print(txt)
	if "农村土地" in txt or "归户表" in txt or "表6" in txt:
		CountTemp = 1
		Code_DK = []
		bFlag = True
		continue
	if bFlag:
		digitals = re.findall(r'\d+', txt)
		for digital in digitals:
			if digital.startswith(('0','1','2','3','4','5','6','7','8','9')):
				if len(digital)==5:
					Code_DK.append(digital)
		print(Code_DK)
		CountTemp += 1


if len(Code_DK)==0:
	print("归户表格式错误！归户表设置最多为4页")

# 获取承包方代码以及地块代码（完善）
# print(CountTemp)
pageNum = AllPages-2*CountTemp-1
print('--->'+str(pageNum))
pdfWriter = PdfFileWriter()     #生成一个空白的pdf文件
pdfWriter.addPage(pdfReader.getPage(pageNum))
pdfOutput = open('./'+tempoutPdfName,'wb')
pdfWriter.write(pdfOutput)                           #将复制的内容全部写入合并的pdf
pdfOutput.close()
print('--->'+'open pdf')
image_pdf = Image(filename='./'+tempoutPdfName,resolution=RESOLUTION)
print('--->'+'convert pdf to jpeg')
image_jpeg = image_pdf.convert('jpeg')
img_page = Image(image=image_jpeg)
req_image = img_page.make_blob('jpeg')
print('--->'+'recognite')
txt = tool.image_to_string(
	PI.open(io.BytesIO(req_image)),
	lang=lang,
	builder=pyocr.builders.TextBuilder()
)
os.remove(tempoutPdfName)
print(txt)
Pre_Code_DK=""
if "界址点成果表" in txt or "界址点坐标" in txt or "界址点编号" in txt:
	digitals = re.findall(r'\d+', txt) #filter(txt.isdigit(), txt.encode('gbk'))
	for digital in digitals:
		if digital.startswith('140'):
			if len(digital)==19:
				Code_DK_FULL = digital
			else:
				Code_CBF = digital
if Code_DK_FULL:
	Pre_Code_DK = Code_DK_FULL[0:15]
else:
	Pre_Code_DK = Code_CBF[0:15]
if  Pre_Code_DK==None or Pre_Code_DK=="":
	print("界址点成果表所在位置错误！归户表设置最多为4页")
	sys.exit(1)
if Code_CBF==None:
	print("承包方代码读取错误！请检查最后一个界址点成果表")
	sys.exit(1)
print('编码 '+str(Pre_Code_DK)+' '+str(Code_DK)+' '+str(Code_CBF))

#-------  先分解部分PDF--------------
#（1）承包方调查表（表2）
pdfWriter = PdfFileWriter()     #生成一个空白的pdf文件
pdfWriter.addPage(pdfReader.getPage(0))
outPdfName = resPath+"CBFDCB"+Code_CBF+".pdf"
pdfOutput = open(outPdfName,'wb')
pdfWriter.write(pdfOutput)                         #将复制的内容全部写入合并的pdf
pdfOutput.close()

#(2)分离归户表(表6)
pdfWriter = PdfFileWriter()     #生成一个空白的pdf文件
for x in range(0,CountTemp):
	pdfWriter.addPage(pdfReader.getPage(AllPages-CountTemp+x))
outPdfName = resPath+"CBJYQGH"+Code_CBF+".pdf"
pdfOutput = open(outPdfName,'wb')
pdfWriter.write(pdfOutput)                         #将复制的内容全部写入合并的pdf
pdfOutput.close()

#(2)分离核实表 (表4)
pdfWriter = PdfFileWriter()     #生成一个空白的pdf文件
for x in range(0,CountTemp):
	pdfWriter.addPage(pdfReader.getPage(AllPages-2*CountTemp+x))
outPdfName = resPath+"CBJYQDCHS"+Code_CBF+".pdf"
pdfOutput = open(outPdfName,'wb')
pdfWriter.write(pdfOutput)                         #将复制的内容全部写入合并的pdf
pdfOutput.close()

#(2) 地块调查表
for ii in range(0,len(Code_DK)):
	temp= AllPages-2*CountTemp-3*len(Code_DK) + ii*3
	pdfWriter = PdfFileWriter()     #生成一个空白的pdf文件
	pdfWriter.addPage(pdfReader.getPage(temp))
	pdfWriter.addPage(pdfReader.getPage(temp+1))
	pdfWriter.addPage(pdfReader.getPage(temp+2))
	pdfOutput = open(resPath+"CBFDKDCB"+Pre_Code_DK+Code_DK[ii]+".pdf",'wb')
	pdfWriter.write(pdfOutput)                         #将复制的内容全部写入合并的pdf
	pdfOutput.close()
#---------------------------------------------
#(3) 承包方代表身份证及证明
pdfWriter = PdfFileWriter()     #生成一个空白的pdf文件
pdfWriter.addPage(pdfReader.getPage(1))
#检查是否有证明，在第3页
pageNum = 2  #第三页，从0开始计算
print('--->'+str(pageNum))
pdfWriter1 = PdfFileWriter()     #生成一个空白的pdf文件
pdfWriter1.addPage(pdfReader.getPage(pageNum))
pdfOutput = open('./'+tempoutPdfName,'wb')
pdfWriter1.write(pdfOutput)                           #将复制的内容全部写入合并的pdf
pdfOutput.close()
print('--->'+'open pdf')
image_pdf = Image(filename='./'+tempoutPdfName,resolution=RESOLUTION)
print('--->'+'convert pdf to jpeg')
image_jpeg = image_pdf.convert('jpeg')
img_page = Image(image=image_jpeg)
req_image = img_page.make_blob('jpeg')
print('--->'+'recognite')
txt = tool.image_to_string(
	PI.open(io.BytesIO(req_image)),
	lang=lang,
	builder=pyocr.builders.TextBuilder()
)
os.remove(tempoutPdfName)
print(txt)
bZM=False
if "证明" in txt or "兹证明" in txt or "情况属实" in txt:
	pdfWriter.addPage(pdfReader.getPage(2))
	bZM = True
pdfOutput = open(resPath+"CBFMC"+Code_CBF+".pdf",'wb')
pdfWriter.write(pdfOutput)                         #将复制的内容全部写入合并的pdf
pdfOutput.close()

#（4）合同
CurrentPage = AllPages-2*CountTemp-3*len(Code_DK)-1
while True:
	print('--->'+str(CurrentPage))
	pdfWriter = PdfFileWriter()     #生成一个空白的pdf文件
	pdfWriter.addPage(pdfReader.getPage(CurrentPage))
	pdfOutput = open('./'+tempoutPdfName,'wb')
	pdfWriter.write(pdfOutput)                           #将复制的内容全部写入合并的pdf
	pdfOutput.close()
	print('--->'+'open pdf')
	image_pdf = Image(filename='./'+tempoutPdfName,resolution=RESOLUTION)
	print('--->'+'convert pdf to jpeg')
	image_jpeg = image_pdf.convert('jpeg')
	img_page = Image(image=image_jpeg)
	req_image = img_page.make_blob('jpeg')
	print('--->'+'recognite')
	txt = tool.image_to_string(
		PI.open(io.BytesIO(req_image)),
		lang=lang,
		builder=pyocr.builders.TextBuilder()
	)
	os.remove(tempoutPdfName)
	print(txt)
	if "一式三份" in txt or "单位各一份" in txt or "另行拍卖" in txt:
		break;
	CurrentPage = CurrentPage - 1

pdfWriter = PdfFileWriter()     #生成一个空白的pdf文件
pdfWriter.addPage(pdfReader.getPage(CurrentPage))
pdfWriter.addPage(pdfReader.getPage(CurrentPage+1))
pdfOutput = open(resPath+"CBYJ"+Code_CBF+".pdf",'wb')
pdfWriter.write(pdfOutput)                         #将复制的内容全部写入合并的pdf
pdfOutput.close()

#(5)家庭成员
if bZM:
	JTCY_beg = 3
else:
	JTCY_beg = 2

pdfWriter = PdfFileWriter()     #生成一个空白的pdf文件
for x in range(JTCY_beg,CurrentPage):
	pdfWriter.addPage(pdfReader.getPage(x))
pdfOutput = open(resPath+"CBFJTCY"+Code_CBF+".pdf",'wb')
pdfWriter.write(pdfOutput)                         #将复制的内容全部写入合并的pdf
pdfOutput.close()


sys.exit(1)


for pageNum in range(pdfReader.numPages):
	if pageNum<55:
		continue
	print('--->'+str(pageNum))
	pdfWriter = PdfFileWriter()     #生成一个空白的pdf文件
	pdfWriter.addPage(pdfReader.getPage(pageNum))
	pdfOutput = open('./'+outPdfName,'wb')
	pdfWriter.write(pdfOutput)                           #将复制的内容全部写入合并的pdf
	pdfOutput.close()
	print('--->'+'open pdf')
	image_pdf = Image(filename='./'+outPdfName,resolution=RESOLUTION)
	print('--->'+'convert pdf to jpeg')
	image_jpeg = image_pdf.convert('jpeg')
	img_page = Image(image=image_jpeg)
	req_image = img_page.make_blob('jpeg')
	print('--->'+'recognite')
	txt = tool.image_to_string(
		PI.open(io.BytesIO(req_image)),
		lang=lang,
		builder=pyocr.builders.TextBuilder()
	)
	os.remove(outPdfName)
	print(txt)
	Page_CBFJTCY_beg = 2
	#证明
	if "证明" in txt or "兹证明" in txt or "情况属实" in txt:
		Page_ZM_beg = pageNum
		Page_ZM_end = pageNum
		#承包方家庭成员
		Page_CBFJTCY_beg = pageNum+1
		print('承包方证明'+str(Page_ZM_beg)+' '+str(Page_ZM_end))
	#找合同关键字
	if "一式三份" in txt or "单位各一份" in txt or "另行拍卖" in txt:
		Page_CBFJTCY_end = pageNum-1
		Page_HT_beg = pageNum;
		Page_HT_end = pageNum+1;
		print('承包方家庭成员'+' '+str(Page_CBFJTCY_beg)+str(Page_CBFJTCY_end))
		print('合成 '+str(Page_HT_beg)+' '+str(Page_HT_end))
	#找承包地块调查表（表3）关键字
	if "承包地块调查表" in txt or "表3" in txt:
		Page_CBDKDCB_beg = pageNum;
	#提取承包方代码和承包地块代码
	if "界址点成果表" in txt or "界址点坐标" in txt or "界址点编号" in txt:
		digitals = re.findall(r'\d+', txt) #filter(txt.isdigit(), txt.encode('gbk'))
		for digital in digitals:
			if digital.startswith('140'):
				if len(digital)==19:
					Code_DK = digital
				else:
					Code_CBF = digital
		print('编码 '+str(Code_CBF)+' '+str(Code_DK))
	#提取承包经营权调查结果核实表（表4）
	if "承包经营权" in txt or "结果核实表" in txt or "表4" in txt:
		Page_CBJYQHSB_beg = pageNum;
		Page_CBDKDCB_end = pageNum-1;
		#print('结果核实表 '+str(Page_CBDKDCB_beg)+' '+str(Page_CBDKDCB_end))
		print('结果核实表 '+str(Page_CBDKDCB_end))
	#提取承包经营权调查结果核实表（表6）
	if "农村土地" in txt or "归户表" in txt or "表6" in txt:
		Page_GHB_beg = pageNum;
		age_CBJYQHSB_end = pageNum-1
		print('归户表 '+str(Page_GHB_beg))
		break

	#os.system("pause")

Page_GHB_end = pdfReader.numPages;
