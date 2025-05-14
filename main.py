from Crypto.Cipher import AES
import base64

def decrypt_code(encrypted_data: str) -> str:
    try:
        key = b'huangyinghuangying'  # 注意不是 fromhex，而是直接 UTF-8 编码
        iv = b'sskjKingFree5138'

        ciphertext = base64.b64decode(encrypted_data)
        cipher = AES.new(key, AES.MODE_CBC, iv)
        decrypted = cipher.decrypt(ciphertext)

        # 如果不是用标准填充方式，手动 strip 空字符
        plaintext = decrypted.decode('utf-8', errors='ignore').rstrip('\x00').strip()
        return plaintext

    except Exception as e:
        return f"[解密失败] 错误: {str(e)}"

# 测试
encrypted_data = "tydAd9ijOGonZN2I/FGYsQ=="
result = decrypt_code(encrypted_data)
print(f"Decrypted Code: {result}")  # 预期应为 sZi0 或类似结构
