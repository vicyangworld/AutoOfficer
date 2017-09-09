import ctypes
import os
STD_INPUT_HANDLE = -10
STD_OUTPUT_HANDLE= -11
STD_ERROR_HANDLE = -12

FOREGROUND_BLACK = 0x0
FOREGROUND_BLUE = 0x01 # text color contains blue.
FOREGROUND_GREEN= 0x02 # text color contains green.
FOREGROUND_RED = 0x04 # text color contains red.
FOREGROUND_INTENSITY = 0x08 # text color is intensified.

BACKGROUND_BLUE = 0x10 # background color contains blue.
BACKGROUND_GREEN= 0x20 # background color contains green.
BACKGROUND_RED = 0x40 # background color contains red.
BACKGROUND_INTENSITY = 0x80 # background color is intensified.
#上面这一大段都是在设置前景色和背景色，其实可以用数字直接设置，我的代码直接用数字设置颜色
    
class  CmdFormat(object):
    """docstring for  CmdFormat"""
    std_out_handle = ctypes.windll.kernel32.GetStdHandle(STD_OUTPUT_HANDLE)

    def __init__(self, WinTitle="Console Window",\
        color=FOREGROUND_RED | FOREGROUND_GREEN | FOREGROUND_BLUE | FOREGROUND_INTENSITY,\
                ):
        super( CmdFormat, self).__init__()
        self.WinTitle = WinTitle
        os.system("title " + WinTitle)

    def set_cmd_color(self, color, handle=std_out_handle):
        bool = ctypes.windll.kernel32.SetConsoleTextAttribute(handle, color)
        return bool
    
    def reset_color(self):
        self.set_cmd_color(FOREGROUND_RED | FOREGROUND_GREEN | FOREGROUND_BLUE | FOREGROUND_INTENSITY)
        #初始化颜色为黑色背景，纯白色字，CMD默认是灰色字体的
    def print_white_text(self,print_text,endd='\n'):
        self.reset_color()
        print(print_text,end=endd)

    def print_red_text(self, print_text,endd='\n'):
        self.set_cmd_color(4 | 8)
        print(print_text,end=endd)
        self.reset_color()
        #红色字体
        
    def print_green_text(self, print_text):
        self.set_cmd_color(FOREGROUND_GREEN | FOREGROUND_INTENSITY)
        c = input(print_text)
        self.reset_color()
        return c
        #绿色字体。实现的是，让用户输入的字体是绿色的，记得返回函数值。
        
    def print_yellow_text(self, print_text,endd='\n'): 
        self.set_cmd_color(6 | 8)
        print(print_text,end=endd)
        self.reset_color()
        #黄色字体

    def print_blue_text(self, print_text,endd='\n'): 
        self.set_cmd_color(1 | 10)
        print(print_text,end=endd)
        self.reset_color()
        #蓝色字体
if __name__ == '__main__':
    clr = CmdFormat("赟哥特供")
    clr.set_cmd_color(FOREGROUND_RED | FOREGROUND_GREEN | FOREGROUND_BLUE | FOREGROUND_INTENSITY)
    clr.print_red_text('red')
    clr.print_green_text("输入： ")
    clr.print_blue_text('blue')
    clr.print_yellow_text('yellow')
    input()