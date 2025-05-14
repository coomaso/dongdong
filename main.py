import requests
import base64
from Crypto.Cipher import AES
import time

# 解密函数
def aes_decrypt_base64(encrypted_base64: str) -> str:
    key = b"6875616E6779696E6875616E6779696E"  # AES-256 密钥
    iv = b"sskjKingFree5138"                  # IV
    encrypted_bytes = base64.b64decode(encrypted_base64)
    cipher = AES.new(key, AES.MODE_CBC, iv)
    decrypted_bytes = cipher.decrypt(encrypted_bytes)
    return decrypted_bytes.rstrip(b'\x00').decode("utf-8")

# 1. 获取当前时间戳（毫秒）
timestamp = str(int(time.time() * 1000))

# 2. 请求验证码接口
code_url = f"https://www.ycjsjg.net/ycdc/bakCmisYcOrgan/getCreateCode?codeValue={timestamp}"
response = requests.get(code_url)
resp_json = response.json()

if resp_json["code"] != 0:
    print("验证码请求失败：", resp_json)
    exit()

# 3. 解密验证码
encrypted_data = resp_json["data"]
decrypted_code = aes_decrypt_base64(encrypted_data)
print("解密后的验证码:", decrypted_code)

# 4. 拼接最终目标URL
final_url = (
    "https://www.ycjsjg.net/ycdc/bakCmisYcOrgan/getCurrentIntegrityPage"
    f"?pageSize=10&cioName=%E5%85%AC%E5%8F%B8&page=0"
    f"&code={decrypted_code}&codeValue={timestamp}"
)

# 5. 请求数据
final_response = requests.get(final_url)
final_json = final_response.json()

# 6. 打印结果
print("返回结果：")
print(final_json)
