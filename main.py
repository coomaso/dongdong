from Crypto.Cipher import AES
import base64

class CryptoUtils:
    @staticmethod
    def decrypt_cx_data(encrypted_b64: str) -> str:
        """信用分加密数据解密"""
        # 固定密钥（对应 JavaScript 代码中的 huangyinhuanhuan...）
        key = bytes.fromhex("6875616E6779696E6875616E6779696E")
        
        # 固定 IV（对应 sskjKingFree5138）
        iv = "sskjKingFree5138".encode('utf-8').ljust(16, b'\0')[:16]
        
        # 解密处理
        cipher = AES.new(key, AES.MODE_CBC, iv=iv)
        ciphertext = base64.b64decode(encrypted_b64)
        decrypted = cipher.decrypt(ciphertext)
        
        # 清洗解密结果
        return decrypted.decode('utf-8', errors='ignore').split('\x00')[0].strip()

# 使用示例
if __name__ == "__main__":
    data = "tydAd9ijOGonZN2I/FGYsQ=="
    print(CryptoUtils.decrypt_cx_data(data))  # 输出 sZi0
