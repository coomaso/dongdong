import hashlib

def decrypt_code(code_value, encrypted_data):
    # 生成 MD5 哈希密钥（16 字节）
    key_str = str(code_value)
    key = hashlib.md5(key_str.encode()).digest()  # 关键修改
    
    ciphertext = base64.b64decode(encrypted_data)
    cipher = AES.new(key, AES.MODE_ECB)
    decrypted = cipher.decrypt(ciphertext)
    
    # 调试：打印解密后的原始数据（十六进制）
    print(f"[DEBUG] Decrypted raw (hex): {decrypted.hex()}")  
    
    # 直接截取前4字节（假设明文为4字符）
    plaintext = decrypted[:4].decode('utf-8')  # 忽略填充
    return plaintext
