import win32com.client
import uuid
def encrypt(key,content): # key:密钥,content:明文
    EncryptedData = win32com.client.Dispatch('CAPICOM.EncryptedData')
    EncryptedData.Algorithm.KeyLength = 5
    EncryptedData.Algorithm.Name = 2
    EncryptedData.SetSecret(key)
    EncryptedData.Content = content
    return EncryptedData.Encrypt()

def decrypt(key,content): # key:密钥,content:密文
    EncryptedData = win32com.client.Dispatch('CAPICOM.EncryptedData')
    EncryptedData.Algorithm.KeyLength = 5
    EncryptedData.Algorithm.Name = 2
    EncryptedData.SetSecret(key)
    EncryptedData.Decrypt(content)
    str = EncryptedData.Content
    return str

def get_mac_address(): 
    mac=uuid.UUID(int = uuid.getnode()).hex[-12:] 
    return ":".join([mac[e:e+2] for e in range(0,11,2)])

def generate_lisence(s):
    with open('./lisence.lis', 'w') as f:
        f.write(s)

if __name__ == '__main__':
    s1 = encrypt('cxr', '1C:C1:DE:34:E1:1E')
    s2 = decrypt('cxr', s1)
    print(s1)
    print(s2)
    generate_lisence(s1)



# MGEGCSsGAQQBgjdYA6BUMFIGCisGAQQBgjdYAwGgRDBCAgMCAAECAmYBAgFABAgq
# GpllWj9cswQQh/fnBUZ6ijwKDTH9DLZmBgQYmfaZ3VFyS/lq391oDtjlcRFGnXpx
# lG7o
# hello world