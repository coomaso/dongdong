from Crypto.Cipher import AES
from Crypto.Util.Padding import unpad
import base64

def decrypt_code(code_value, encrypted_data):
    # 将 codeValue 转换为字符串并编码为字节
    key = str(code_value).encode('utf-8')
    # AES-128 需要 16 字节密钥，填充或截断到 16 字节
    key = key.ljust(16, b'\0')[:16]  # 填充零并截断
    
    # Base64 解码密文
    ciphertext = base64.b64decode(encrypted_data)
    
    # AES-ECB 解密
    cipher = AES.new(key, AES.MODE_ECB)
    decrypted = cipher.decrypt(ciphertext)
    
    # 去除 PKCS#7 填充
    plaintext = unpad(decrypted, AES.block_size).decode('utf-8')
    
    return plaintext

# 示例数据
code_value = 1747204625326
encrypted_data = "tydAd9ijOGonZN2I/FGYsQ=="

# 解密并输出
code = decrypt_code(code_value, encrypted_data)
print(f"Decrypted Code: {code}")  # 输出 sZi0
