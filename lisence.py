# import win32com.client
import uuid
import CmdFormat

import os,sys

CDMF = CmdFormat.CmdFormat("PDF分离及识别器 注册机")

# def encrypt(key,content): # key:密钥,content:明文
#     EncryptedData = win32com.client.Dispatch('CAPICOM.EncryptedData')
#     EncryptedData.Algorithm.KeyLength = 5
#     EncryptedData.Algorithm.Name = 2
#     EncryptedData.SetSecret(key)
#     EncryptedData.Content = content
#     return EncryptedData.Encrypt()

# def decrypt(key,content): # key:密钥,content:密文
#     EncryptedData = win32com.client.Dispatch('CAPICOM.EncryptedData')
#     EncryptedData.Algorithm.KeyLength = 5
#     EncryptedData.Algorithm.Name = 2
#     EncryptedData.SetSecret(key)
#     EncryptedData.Decrypt(content)
#     str = EncryptedData.Content
#     return str




def encrypt(keyy, s):
    key = 15
    b = bytearray(str(s).encode("gbk"))
    n = len(b) # 求出 b 的字节数
    c = bytearray(n*2)
    j = 0
    for i in range(0, n):
        b1 = b[i]
        b2 = b1 ^ key # b1 = b2^ key
        c1 = b2 % 16
        c2 = b2 // 16 # b2 = c2*16 + c1
        c1 = c1 + 65
        c2 = c2 + 65 # c1,c2都是0~15之间的数,加上65就变成了A-P 的字符的编码
        c[j] = c1
        c[j+1] = c2
        j = j+2
    return c.decode("gbk")
 
def decrypt(keyy, s):
    key = 15
    c = bytearray(str(s).encode("gbk"))
    n = len(c) # 计算 b 的字节数
    if n % 2 != 0 :
        return ""
    n = n // 2
    b = bytearray(n)
    j = 0
    for i in range(0, n):
        c1 = c[j]
        c2 = c[j+1]
        j = j+2
        c1 = c1 - 65
        c2 = c2 - 65
        b2 = c2*16 + c1
        b1 = b2^ key
        b[i]= b1
    try:
        return b.decode("gbk")
    except:
        return "failed"


def get_mac_address():
    mac=uuid.UUID(int = uuid.getnode()).hex[-12:] 
    return "-".join([mac[e:e+2] for e in range(0,11,2)])

def generate_lisence(s):
    with open('./lisence.lis', 'w') as f:
        f.write(s)
    if os.path.exists('./lisence.lis'):
        CDMF.print_yellow_text('已经在当前目录生成许可lisence.lis')
        tt = input('按任意键退出...')
        sys.exit()
    else:
        CDMF.print_red_text('许可生成失败！')


if __name__ == '__main__':
    # s1 = encrypt('cxr', '1C:C1:DE:34:E1:1E')
    # s2 = decrypt('cxr', s1)
    # print(s1)
    # print(s2)
    key = 'cxr'
    mac_address = input('请输入机器物理地址：')
    generate_lisence(encrypt(key,mac_address))



# MGEGCSsGAQQBgjdYA6BUMFIGCisGAQQBgjdYAwGgRDBCAgMCAAECAmYBAgFABAgq
# GpllWj9cswQQh/fnBUZ6ijwKDTH9DLZmBgQYmfaZ3VFyS/lq391oDtjlcRFGnXpx
# lG7o
# hello world
