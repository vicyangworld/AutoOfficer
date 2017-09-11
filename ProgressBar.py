import sys, time
from CmdFormat import CmdFormat

class ProgressBar(CmdFormat):
    def __init__(self, count = 0, total = 0, width = 80, bWithheader=True, bWithPercent=True,barColor='white'):
        super(CmdFormat, self).__init__()
        self.count = count
        self.total = total
        self.width = width
        self.bWithheader = bWithheader
        self.bWithPercent = bWithPercent
        self.__barColor = barColor
    def __Set_bar_color(self):
        if type(self.__barColor) != type('a'):
            raise TypeError("Wrong argument type of __Set_bar_color(color) in class ProgressBar！")
        if self.__barColor=='red':
            self.set_cmd_color(4|8)
        if self.__barColor=='green':
            self.set_cmd_color(2|8)
        if self.__barColor=='blue':
            self.set_cmd_color(1|10)
        if self.__barColor=='yellow':
            self.set_cmd_color(6|8)
    def Move(self, s):
        self.count += 1
        sys.stdout.write(' '*(self.width + 20) + '\r')
        sys.stdout.flush()
        print(s)
        progress = self.width * self.count / self.total
        if(self.bWithheader):sys.stdout.write('{0:3}/{1:3}:'.format(self.count, self.total))
        percent = progress * 100.0 / self.total

        if(self.bWithPercent): 
            self.__Set_bar_color()
            sys.stdout.write('[' + int(progress)*'>' + int(self.width - progress)*'-' + ']' + ' %.2f' % progress + '%' + '\r')
            self.reset_color()
        else:
            self.__Set_bar_color()
            sys.stdout.write('[' + int(progress)*'>' + int(self.width - progress)*'-' + ']'+'\r')
            self.reset_color()
        if progress == self.width:
            sys.stdout.write('\n')
        sys.stdout.flush()
    def Set_cmd_color(self,color):
        if type(color) != type('a'):
            raise TypeError("Wrong argument type of __Set_bar_color(color) in class ProgressBar！")
        if color=='red':
            self.set_cmd_color(4|8)
        if color=='green':
            self.set_cmd_color(2|8)
        if color=='blue':
            self.set_cmd_color(1|10)
        if color=='yellow':
            self.set_cmd_color(6|8)
=
if __name__ == '__main__':
    bar = ProgressBar(total = 15,bWithheader=True,bWithPercent=True,barColor='green')
    for i in range(15):
        bar.Set_cmd_color('red')
        bar.Move('sdfds ')
        time.sleep(1)

        