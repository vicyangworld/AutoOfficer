import os,sys,shutil

def mkdir(path):
	path=path.strip()
	path=path.rstrip("\\")
	isExists=os.path.exists(path)
	if not isExists:
		os.makedirs(path)
		return True
	else:
		return False

def ReRange(path):
	RootPath = path
	allCountries = os.listdir(RootPath)
	for x in allCountries:
		if os.path.isdir(x) and x.endswith('村'):
			dirsInOneCountry = os.listdir(x)
			resPath = os.path.join(RootPath,x,'temp\\')
			bResPath = False;
			if os.path.exists(resPath):
				if os.listdir(resPath):
					bResPath = True;
				else:
					shutil.rmtree(resPath)

			for y in dirsInOneCountry:
				if not bResPath and '承包地块调查表' in y:
					shutil.copytree(os.path.join(RootPath,x,y),resPath)  #结果文件
					break

			persons=os.listdir(resPath)
			subCountryPath = os.path.join(RootPath,x)

			for y in dirsInOneCountry:
				subCountryPath = os.path.join(RootPath,x,y)	
				if '承包方调查表' in y:
					CBFDCBs = os.listdir(subCountryPath)
					for z in CBFDCBs:
						name = z.split('-')[1]
						if name in persons:
							shutil.copy(os.path.join(subCountryPath,z),os.path.join(resPath,name,z))
				if '公示结果归户表' in y:
					GHBs = os.listdir(subCountryPath)
					for z in GHBs:
						name = z.split('-')[0]
						if name in persons:
							shutil.copy(os.path.join(subCountryPath,z),os.path.join(resPath,name,z))
				if '农户代表声明书' in y:
					SMSs = os.listdir(subCountryPath)
					for z in SMSs:
						name = z.split('-')[1]
						if name in persons:
							shutil.copy(os.path.join(subCountryPath,z),os.path.join(resPath,name,z))
				if '承包合同' in y:
					HTs = os.listdir(subCountryPath)
					for z in HTs:
						name = z.split('-')[1]
						if name in persons:
							shutil.copy(os.path.join(subCountryPath,z),os.path.join(resPath,name,z))
				if '发包方调查表' in y:
					FBFDCBs = os.listdir(subCountryPath)
					if len(FBFDCBs) == 1:
						FBFDM1 = FBFDCBs[0].split('(')[1]
						FBFDM = FBFDM1.split(')')[0]
					else:
						print("发包方调查表中有多个文件，请检查")
						quit('请按任意键退出...')
						sys.exti(1)
			# rename
			if FBFDM!="":
				for y in dirsInOneCountry:
					if '承包方调查表' in y:
						CBFDCBs = os.listdir(os.path.join(RootPath,x,y))
						for z in CBFDCBs:
							CBF_CODE_NAME = '_'.join(z.split('-')[0:2])
							name = z.split('-')[1]
							if name in persons:
								os.rename(os.path.join(resPath,name),os.path.join(resPath,FBFDM+CBF_CODE_NAME))
								if not os.path.exists(os.path.join(RootPath,x,FBFDM+CBF_CODE_NAME)):
									shutil.move(os.path.join(resPath,FBFDM+CBF_CODE_NAME),os.path.join(RootPath,x)) 
	shutil.rmtree(resPath)
if __name__ == '__main__':
	ReRange(os.getcwd())
