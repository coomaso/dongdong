from Crypto.Cipher import AES
from Crypto.Util.Padding import unpad
import base64

def decrypt_code(code_value, encrypted_data):
    key = str(code_value).encode('utf-8')
    key = key.ljust(16, b'\0')[:16]  # 填充到 16 字节
    
    # 调试：打印密钥的十六进制
    print(f"[DEBUG] Key (hex): {key.hex()}")  
    
    ciphertext = base64.b64decode(encrypted_data)
    
    # 调试：打印密文长度
    print(f"[DEBUG] Ciphertext length: {len(ciphertext)}")  
    
    cipher = AES.new(key, AES.MODE_ECB)
    decrypted = cipher.decrypt(ciphertext)
    
    # 调试：打印解密后的原始数据（十六进制）
    print(f"[DEBUG] Decrypted raw (hex): {decrypted.hex()}")  
    
    try:
        plaintext = unpad(decrypted, AES.block_size).decode('utf-8')
    except ValueError as e:
        # 调试：捕获填充错误并输出原始数据
        print(f"[ERROR] 填充错误: {e}")
        print(f"[DEBUG] 解密后的原始数据: {decrypted}")
        raise
    
    return plaintext

code_value = 1747204625326
encrypted_data = "tydAd9ijOGonZN2I/FGYsQ=="

code = decrypt_code(code_value, encrypted_data)
print(f"Decrypted Code: {code}")
