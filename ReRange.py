import os,sys,shutil

RootPath = os.getcwd()
allCountries = os.listdir(RootPath)
for x in allCountries:
	if os.path.isdir(x):
		dirsInOneCountry = os.listdir(x)
		for y in dirsInOneCountry:
			if '承包地块调查表' in y:
				shutil.copytree(os.path.join(RootPath,x,y),os.path.join(RootPath,x,"结果"))  #结果文件
				print(y)
			# if '承包方调查表' in y:
			# 	files = os.listdir(y)

				