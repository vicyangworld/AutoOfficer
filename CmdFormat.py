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

    def print_white_text(self,print_text,end='\n'):
        self.reset_color()
        print(print_text,end=end)

    def print_red_text(self, print_text,end='\n'):
        self.set_cmd_color(4 | 8)
        print(print_text,end=end)
        self.reset_color()

    def print_green_input_text(self, print_text):
        self.set_cmd_color(FOREGROUND_GREEN | FOREGROUND_INTENSITY)
        c = input(print_text)
        self.reset_color()
        return c

    def print_green_text(self, print_text,end='\n'):
        self.set_cmd_color(FOREGROUND_GREEN | FOREGROUND_INTENSITY)
        print(print_text,end=end)
        self.reset_color()

    def print_yellow_text(self, print_text,end='\n'):
        self.set_cmd_color(6 | 8)
        print(print_text,end=end)
        self.reset_color()

    def print_blue_text(self, print_text,end='\n'):
        self.set_cmd_color(1 | 10)
        print(print_text,end=end)
        self.reset_color()

if __name__ == '__main__':
    clr = CmdFormat("Window Title")
    clr.set_cmd_color(FOREGROUND_RED | FOREGROUND_GREEN | FOREGROUND_BLUE | FOREGROUND_INTENSITY)
    clr.print_red_text('red')
    clr.print_green_text("green")
    clr.print_green_input_text("input: ")
    clr.print_blue_text('blue')
    clr.print_yellow_text('yellow')
    input()