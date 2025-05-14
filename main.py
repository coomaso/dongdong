from Crypto.Cipher import AES
import base64

def decrypt_code(encrypted_data: str) -> str:
    try:
        # 正确的密钥长度：16 字节（AES-128）
        key_hex = "6875616E6779696E6875616E6779696E"
        key = bytes.fromhex(key_hex)
        iv = b'sskjKingFree5138'  # IV 也是 16 字节

        ciphertext = base64.b64decode(encrypted_data)
        cipher = AES.new(key, AES.MODE_CBC, iv)
        decrypted = cipher.decrypt(ciphertext)

        # 若非 PKCS7 填充，则使用手动去尾部空字节
        plaintext = decrypted.decode('utf-8', errors='ignore').rstrip('\x00').strip()
        return plaintext

    except Exception as e:
        return f"[解密失败] 错误: {str(e)}"

# 测试用例
encrypted_data = "tydAd9ijOGonZN2I/FGYsQ=="
result = decrypt_code(encrypted_data)
print(f"Decrypted Code: {result}")
