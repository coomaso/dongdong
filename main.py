from Crypto.Cipher import AES
import base64
from Crypto.Util.Padding import unpad
import hashlib

def decrypt_code(code_value, encrypted_data):
    # 生成密钥和 IV（与 JavaScript 代码中的值匹配）
    key_hex = "6875616E6779696E6875616E6779696E"  # 代码中的固定密钥
    iv_str = "sskjKingFree5138"                   # 代码中的固定 IV
    
    # 将 hex 密钥转换为字节
    key = bytes.fromhex(key_hex)
    iv = iv_str.encode('utf-8').ljust(16, b'\0')[:16]  # IV 固定为 16 字节
    
    # Base64 解码密文
    ciphertext = base64.b64decode(encrypted_data)
    
    # AES-CBC 解密（无填充）
    cipher = AES.new(key, AES.MODE_CBC, iv=iv)
    decrypted = cipher.decrypt(ciphertext)
    
    # 处理无填充：去除末尾的空字节或特殊字符
    plaintext = decrypted.decode('utf-8', errors='ignore').rstrip('\x00').strip()
    return plaintext

# 测试数据
code_value = 1747204625326  # 注意：此处 codeValue 未参与密钥生成，实际密钥为固定值
encrypted_data = "tydAd9ijOGonZN2I/FGYsQ=="

# 解密
code = decrypt_code(code_value, encrypted_data)
print(f"Decrypted Code: {code}")  # 输出 sZi0
