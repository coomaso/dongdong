from Crypto.Cipher import AES
import base64
import hashlib
from Crypto.Util.Padding import unpad

def decrypt_code_raw(code_value, encrypted_data):
    # 生成 MD5 密钥
    key = hashlib.md5(str(code_value).encode()).digest()

    ciphertext = base64.b64decode(encrypted_data)
    cipher = AES.new(key, AES.MODE_ECB)
    decrypted_bytes = cipher.decrypt(ciphertext)
    return decrypted_bytes

# 运行 raw 解密
code_value = 1747204625326
encrypted_data = "tydAd9ijOGonZN2I/FGYsQ=="
raw_decrypted = decrypt_code_raw(code_value, encrypted_data)
print(f"Raw Decrypted Bytes: {raw_decrypted}")

def decrypt_code(code_value, encrypted_data):
    # 生成 MD5 密钥
    key = hashlib.md5(str(code_value).encode()).digest()

    ciphertext = base64.b64decode(encrypted_data)
    cipher = AES.new(key, AES.MODE_ECB)
    decrypted = cipher.decrypt(ciphertext)

    try:
        # 尝试 PKCS#7 解填充
        plaintext = unpad(decrypted, AES.block_size).decode('utf-8')
    except ValueError:
        # 若填充错误，直接截取前4字节
        plaintext = decrypted[:4].decode('utf-8', errors='ignore')

    return plaintext

# 运行解密
code_value = 1747204625326
encrypted_data = "tydAd9ijOGonZN2I/FGYsQ=="
code = decrypt_code(code_value, encrypted_data)
print(f"Decrypted Code: {code}")
