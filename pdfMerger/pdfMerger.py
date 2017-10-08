import PyPDF2, os
import sys
import CmdFormat
import shutil

CDMF = CmdFormat.CmdFormat("PDF合并器")


class PDFMerger(object):
    """docstring for PDFMerger"""
    def __init__(self,ROOTPATH,DIRSNAMEs):
        self.__ROOTPATH = ROOTPATH+"\\"
        self.__DIRSNAMEs = DIRSNAMEs
        self.__countriesCount = 0
        self.__currentCoutry=""
        self.bRegenerate = True
    def __checkDirs(self,foldersInACountry):
        if self.__DIRSNAMEs[0] not in foldersInACountry:
            #CDMF.print_red_text("在 "+self.__currentCoutry+" 没有找到 "+ self.__DIRSNAMEs[0]+" !")
            return False
        if self.__DIRSNAMEs[1] not in foldersInACountry:
           # CDMF.print_red_text("在 "+self.__currentCoutry+" 没有找到 "+ self.__DIRSNAMEs[1]+" !")
            return False
        if self.__DIRSNAMEs[2] not in foldersInACountry:
            #CDMF.print_red_text("在 "+self.__currentCoutry+" 没有找到 "+ self.__DIRSNAMEs[2]+" !")
            return False
        return True

    def __messages(self):
        CDMF.set_cmd_color(CmdFormat.FOREGROUND_RED | CmdFormat.FOREGROUND_GREEN | \
            CmdFormat.FOREGROUND_BLUE | CmdFormat.FOREGROUND_INTENSITY)
        print("\n")
        print("==========================  欢迎使用  ==================================")
        print("|                                                                      |")
        print("|      将本程序放在根目录，运行之前请确保根目录下具有                  |")
        CDMF.print_red_text("|      (1) *各个村的目录(一个村一个目录)                               |")
        CDMF.print_red_text("|      (2) *每个村目录下包含<'表3','草图PDF','界址点PDF'>              |")
        CDMF.print_red_text("|      (3) *请注意以上三个文件夹中的pdf文件需一一对应                  |")
        print("|      (3) 合并后的文件放在了每个村目录下的<合并>文件夹内              |")
        print("|                                                                      |")
        print("========================================================================")
    def __quiry(self,mes):
        while True:
            content = CDMF.print_green_input_text(mes)
            if content=="y" or content=="Y" or content=="n" or content=="N":
                break
        if content=="y" or content=="Y":
            return True
        else:
            return False

    def __checkRegenerate(self,allCountries):
        isempty = True
        for country in allCountries:
            if os.path.isdir(country):
                self.__currentCoutry = country
                foldersInACountry = os.listdir(country)
                if self.__checkDirs(foldersInACountry):                    
                    coutryPath = os.path.join(self.__ROOTPATH,country)
                    mergerPath = coutryPath + "\\合并\\"
                    if os.path.exists(mergerPath):
                        tempList = os.listdir(mergerPath)
                        if len(tempList)==0:
                            isempty = True
                        else:
                            isempty = False
                            break
        if not isempty:
            self.bRegenerate = self.__quiry("注意：已经存在合并文件，是否重新生成? 请输入y/Y或者n/N:");
            if self.bRegenerate:
                try:
                    shutil.rmtree(mergerPath)
                except Exception as e:
                    raise("删除旧合并文件失败，请查看合并文件是否被占用！")
                    sys.exit(1)

    def __getValidCountryCount(self,allCountries):
        for country in allCountries:
            if os.path.isdir(country):
                self.__currentCoutry = country
                foldersInACountry = os.listdir(country)
                if self.__checkDirs(foldersInACountry):
                    self.__countriesCount  += 1

    def Run(self):
        self.__messages()
        allCountries = os.listdir(self.__ROOTPATH)
        self.__checkRegenerate(allCountries)
        self.__getValidCountryCount(allCountries)  
        CDMF.print_yellow_text("共有 "+str(self.__countriesCount)+" 个村庄需要统计") 
        #多个村
        index = 1
        for country in allCountries:
            coutryPath = os.path.join(self.__ROOTPATH,country)
            pdfFiles1 = []
            pdfFiles2 = []
            pdfFiles3 = []
            if os.path.isdir(country):
                #一个村
                self.__currentCoutry = country

                foldersInACountry = os.listdir(country)
                if not self.__checkDirs(foldersInACountry):
                    continue
                #获取需要合并的文件---------
                for subDir in foldersInACountry:    #遍历所在文件夹内的文件\
                    tempFullName2 = os.path.join(coutryPath,subDir)
                    if os.path.isdir(tempFullName2):
                        files = os.listdir(tempFullName2)
                        if self.__DIRSNAMEs[0] in subDir:
                            for pdfFile in  files:
                               if pdfFile.endswith('.pdf'):        #找到以.pdf结尾的文件
                                    pdfFiles1.append(pdfFile)        #将pdf文件装进pdfFiles数组内
                        elif self.__DIRSNAMEs[1] in tempFullName2:
                            for pdfFile in  files:
                                if pdfFile.endswith('.pdf'):        #找到以.pdf结尾的文件
                                    pdfFiles2.append(pdfFile)        #将pdf文件装进pdfFiles数组内
                        elif self.__DIRSNAMEs[2] in tempFullName2:
                            for pdfFile in  files:
                                if pdfFile.endswith('.pdf'):        #找到以.pdf结尾的文件
                                    pdfFiles3.append(pdfFile)        #将pdf文件装进pdfFiles数组内
                        else:
                            continue
                    #----------------------------
                mergerPath = coutryPath + "\\合并\\"
                if not os.path.exists(mergerPath):
                    os.mkdir(mergerPath)
                pdfFiles1.sort()
                pdfFiles2.sort()
                pdfFiles3.sort()

                length = min([len(pdfFiles1),len(pdfFiles2),len(pdfFiles3)])
                for x in range(0,length):
                    if x<length-1:
                        CDMF.print_blue_text(str(index)+"/"+str(self.__countriesCount)+"  "+country+"    共有 "+str(length) + " 个需要合并"+"  "+str(x+1)+"/"+str(length)+"\r",'')
                    else:
                        CDMF.print_blue_text(str(index)+"/"+str(self.__countriesCount)+"  "+country+"    共有 "+str(length) + " 个需要合并"+"  "+str(x+1)+"/"+str(length),'')
                    pdfWriter = PyPDF2.PdfFileWriter()     #生成一个空白的pdf文件 
                    pdfReader = PyPDF2.PdfFileReader(open(os.path.join(coutryPath,self.__DIRSNAMEs[0],pdfFiles1[x]),'rb'))   #以只读方式依次打开pdf文件
                    for pageNum in range(pdfReader.numPages):
                        pdfWriter.addPage(pdfReader.getPage(pageNum))    #将打开的pdf文件内容一页一页的复制到新建的空白pdf里    
                    pdfReader = PyPDF2.PdfFileReader(open(os.path.join(coutryPath,self.__DIRSNAMEs[1],pdfFiles2[x]),'rb'))   #以只读方式依次打开pdf文件
                    for pageNum in range(pdfReader.numPages):
                        pdfWriter.addPage(pdfReader.getPage(pageNum))    #将打开的pdf文件内容一页一页的复制到新建的空白pdf里    
                    pdfReader = PyPDF2.PdfFileReader(open(os.path.join(coutryPath,self.__DIRSNAMEs[2],pdfFiles3[x]),'rb'))   #以只读方式依次打开pdf文件
                    for pageNum in range(pdfReader.numPages):
                        pdfWriter.addPage(pdfReader.getPage(pageNum))    #将打开的pdf文件内容一页一页的复制到新建的空白pdf里    
                    
                    outPdfName = 'CBFDKDCB'+pdfFiles3[x].split('表')[1]
                    # outPdfName = mergerPath+tempstr+'.pdf'
                    # outPdfName = mergerPath+pdfFiles3[x]
                    if os.path.exists(outPdfName) and not self.bRegenerate:
                        continue
                    pdfOutput = open(mergerPath+outPdfName,'wb') 
                    pdfWriter.write(pdfOutput)                           #将复制的内容全部写入合并的pdf
                    pdfOutput.close()
                    outPdfName=""   #清空outPdfName
                index += 1
                CDMF.print_blue_text("     成功！")

if __name__ == '__main__':
    ROOTPATH = os.getcwd()
    DIRSNAMEs=['表3','草图PDF','界址点PDF']
    Job = PDFMerger(ROOTPATH,DIRSNAMEs)
    Job.Run()
    CDMF.print_yellow_text("任务完成！")
    quit = input("按任意键退出...")
