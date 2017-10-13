CLS 
@echo off 
ECHO. 
ECHO 安装windows环境下必要软件1：tesseract-ocr
ECHO 请稍等... 
start /wait tesseract-ocr-setup-3.02.02.exe /s /v/qn
ECHO.
ECHO 安装windows环境下必要软件2：ImageMagick
ECHO 请稍等... 
start /wait ImageMagick-6.9.9-19-Q8-x86-dll.exe
ECHO.
EXIT 
