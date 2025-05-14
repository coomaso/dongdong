from Crypto.Cipher import AES
import base64

def aes_decrypt_base64(encrypted_base64: str) -> str:
    # 注意：密钥是原始字符串，而不是 hex 解码
    key = b"6875616E6779696E6875616E6779696E"  # 长度32（即AES-256）
    iv = b"sskjKingFree5138"                  # IV必须16字节

    # Base64 解码密文
    encrypted_bytes = base64.b64decode(encrypted_base64)

    # 解密
    cipher = AES.new(key, AES.MODE_CBC, iv)
    decrypted_bytes = cipher.decrypt(encrypted_bytes)

    # 去除 padding（这里是 NoPadding + 手动补 \x00）
    decrypted_text = decrypted_bytes.rstrip(b'\x00').decode("utf-8")

    return decrypted_text



# 假设这是从服务器收到的 base64 编码的加密字符串
encrypted_base64 = "kOfWd64IPxdrps3BLNH/zQ=="  # 仅示例

decrypted = aes_decrypt_base64(encrypted_base64)
print("解密结果:", decrypted)
