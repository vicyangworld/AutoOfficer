import PyPDF2, os
import sys
import CmdFormat
import shutil

CDMF = CmdFormat.CmdFormat("PDF分离器")


class PDFMerger(object):
	"""docstring for PDFMerger"""
	def __init__(self,ROOTPATH):
		self.__ROOTPATH = ROOTPATH+"\\"
		self.__countriesCount = 0
		self.__currentCoutry=""
		self.bRegenerate = True

	def __messages(self):
		CDMF.set_cmd_color(CmdFormat.FOREGROUND_RED | CmdFormat.FOREGROUND_GREEN | \
			CmdFormat.FOREGROUND_BLUE | CmdFormat.FOREGROUND_INTENSITY)
		print("\n")
		print("==========================  欢迎使用  ==================================")

	def __quiry(self,mes):
		while True:
			content = CDMF.print_green_input_text(mes)
			if content=="y" or content=="Y" or content=="n" or content=="N":
				break
		if content=="y" or content=="Y":
			return True
		else:
			return False
	def Run(self):
		self.__messages()
		allFiles = os.listdir(self.__ROOTPATH)
		CDMF.print_blue_text("扫描待统计村民资料...,")
		nNumFile = 0;
		nNumNoContent = 0;
		for fileOrDir in allFiles:
			if fileOrDir.startswith(('1','2','3','4','5','6','7','8','9','0')) and fileOrDir.endswith('.pdf'):
				nNumFile = nNumFile + 1
		CDMF.print_blue_text("扫描完毕！共有 "+str(nNumFile) + " 户的资料,",end='')
		CDMF.print_blue_text("需要统计的有 "+str(nNumFile) + " 户.")
		#多个村
		bdeleteOrg = self.__quiry("是否删掉原文件(请输入y或n):")
		index = 1
		for file in allFiles:
			filefull = os.path.join(self.__ROOTPATH,file)
			if not os.path.isdir(filefull):
				if filefull.endswith('.pdf'):        #找到以.pdf结尾的文件
					(filepath,tempfilename) = os.path.split(filefull)
					(filename,extension) = os.path.splitext(tempfilename)
					if filename.startswith(('1','2','3','4','5','6','7','8','9','0')):
						pdfWriter = PyPDF2.PdfFileWriter()     #生成一个空白的pdf文件
						inPDFfile = open(filefull,'rb')
						pdfReader = PyPDF2.PdfFileReader(inPDFfile)   #以只读方式依次打开pdf文件
						for pageNum in range(pdfReader.numPages):
							if pageNum<pdfReader.numPages-2:
							   pdfWriter.addPage(pdfReader.getPage(pageNum))    #将打开的pdf文件内容一页一页的复制到新建的空白pdf里    
						outPdfName = self.__ROOTPATH+'\\'+'Res_'+filename+'.pdf'
						pdfOutput = open(outPdfName,'wb')
						pdfWriter.write(pdfOutput)                           #将复制的内容全部写入合并的pdf
						pdfOutput.close()
						inPDFfile.close()
						outPdfName=""   #清空outPdfName
					CDMF.print_yellow_text(str(index)+'/'+str(nNumFile)+' --->  '+file+"  成功！")
					index += 1
					if bdeleteOrg:
						os.remove(filefull)

if __name__ == '__main__':
	ROOTPATH = os.getcwd()
	Job = PDFMerger(ROOTPATH)
	Job.Run()
	CDMF.print_yellow_text("任务完成！")
	quit = input("按任意键退出...")
