import PyPDF2, os
import sys
import CmdFormat
import shutil

CDMF = CmdFormat.CmdFormat("Get4PDFPages")
__ROOTPATH = os.getcwd()
allCountries = os.listdir(__ROOTPATH)
def mkdir(path):
	path=path.strip()
	path=path.rstrip("\\")
	if not os.path.exists(path):
		os.makedirs(path)
		return True
	else:
		return False
def __isValidCountry(subPath):
	pdfFileCount = 0
	if not os.path.isdir(subPath):
		return False
	for pdfFile in os.listdir(subPath):
		if pdfFile.endswith('.pdf'):
			pdfFileCount += 1
	if pdfFileCount < 4:
		return False
	return True
def __quiry(mes):
	while True:
		content = CDMF.print_green_input_text(mes)
		if content=="y" or content=="Y" or content=="n" or content=="N":
			break
	if content=="y" or content=="Y":
		return True
	else:
		return False
def __getValidCountryCount(allCountries):
	__countriesCount = 0
	for country in allCountries:
		if os.path.isdir(country) and __isValidCountry(country):
			__countriesCount += 1
	return __countriesCount

CDMF.print_blue_text("共有 "+str(__getValidCountryCount(allCountries))+" 个村庄需要统计")

for country in allCountries:
	coutryPath = os.path.join(__ROOTPATH,country)
	if os.path.isdir(coutryPath)  and __isValidCountry(coutryPath) and not country.startswith('_result_'):
		newFolder = os.path.join(__ROOTPATH,'_result_'+country)
		if not mkdir(newFolder) and len(os.listdir(newFolder)) != 0:
			CDMF.print_red_text(newFolder+"已经存在并且不为空文件夹！")
			if __quiry("是否需要清空该文件夹(是y:否n)？"):
				try:
					shutil.rmtree(newFolder)
					mkdir(newFolder)
				except Exception as e:
					raise("清空失败，请查看文件是否被占用！")
					sys.exit(1)
			else:
				continue
		if not __isValidCountry(coutryPath):
			CDMF.print_red_text('In '+ coutryPath + " 没有足够的pdf文件，至少应该大于4，跳过此次处理")
			continue
		listAllfilesInACountry = os.listdir(coutryPath)
		listAllfilesInACountry.reverse()
		for x in range(0,4):
			CDMF.print_yellow_text("复制"+os.path.join(coutryPath,listAllfilesInACountry[x])+"到"+newFolder+"成功！")
			try:
				shutil.copy(os.path.join(coutryPath,listAllfilesInACountry[x]),newFolder)
			except Exception as e:
				raise("复制"+os.path.join(coutryPath,listAllfilesInACountry[x])+"到"+newFolder+"失败！")
				sys.exit(1)

CDMF.print_yellow_text("任务完成！")
quit = input("按任意键退出...")

