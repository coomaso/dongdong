from Crypto.Cipher import AES
import base64

def aes_decrypt_base64(encrypted_base64: str) -> str:
    # 密钥和 IV（必须为 bytes）
    key = bytes.fromhex("6875616E6779696E6875616E6779696E")  # hex 解码后的16字节 key
    iv = "sskjKingFree5138".encode("utf-8")                  # iv 转为 bytes

    # 解码 base64 编码的密文
    encrypted_bytes = base64.b64decode(encrypted_base64)

    # 创建 AES 解密器
    cipher = AES.new(key, AES.MODE_CBC, iv)

    # 解密
    decrypted_bytes = cipher.decrypt(encrypted_bytes)

    # 去除填充的 \x00 字符
    decrypted_text = decrypted_bytes.rstrip(b'\x00').decode("utf-8")

    return decrypted_text


# 假设这是从服务器收到的 base64 编码的加密字符串
encrypted_base64 = "kOfWd64IPxdrps3BLNH/zQ=="  # 仅示例

decrypted = aes_decrypt_base64(encrypted_base64)
print("解密结果:", decrypted)
