import requests
import base64
from Crypto.Cipher import AES
from Crypto.Util.Padding import unpad
import time
from urllib.parse import quote

# 配置常量
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
}
AES_KEY = bytes.fromhex("6875616E6779696E6875616E6779696E")  # AES-256 密钥（正确十六进制）
AES_IV = b"sskjKingFree5138"  # 16字节IV


def aes_decrypt_base64(encrypted_base64: str) -> str:
    """AES-CBC解密Base64编码数据"""
    try:
        encrypted_bytes = base64.b64decode(encrypted_base64)
        cipher = AES.new(AES_KEY, AES.MODE_CBC, AES_IV)
        decrypted_bytes = cipher.decrypt(encrypted_bytes)
        # 移除PKCS#7填充（根据实际情况选择）
        return unpad(decrypted_bytes, AES.block_size).decode("utf-8")
    except (ValueError, KeyError) as e:
        raise RuntimeError(f"解密失败: {e}")


def get_timestamp() -> str:
    """获取当前毫秒时间戳"""
    return str(int(time.time() * 1000))


def fetch_code(timestamp: str) -> dict:
    """请求验证码接口"""
    url = f"https://www.ycjsjg.net/ycdc/bakCmisYcOrgan/getCreateCode?codeValue={timestamp}"
    try:
        response = requests.get(url, headers=HEADERS, timeout=10)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"请求验证码接口失败: {e}")


def fetch_data(code: str, timestamp: str) -> dict:
    """请求最终数据接口"""
    final_url = (
        "https://www.ycjsjg.net/ycdc/bakCmisYcOrgan/getCurrentIntegrityPage"
        f"?pageSize=10&cioName=%E5%85%AC%E5%8F%B8&page=0"
        f"&code={quote(code)}&codeValue={timestamp}"
    )
    try:
        response = requests.get(final_url, headers=HEADERS, timeout=10)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"请求数据接口失败: {e}")


def main():
    try:
        # 1. 获取时间戳
        timestamp = get_timestamp()
        print(f"生成时间戳: {timestamp}")

        # 2. 获取验证码
        code_response = fetch_code(timestamp)
        if code_response.get("code") != 0:
            print("验证码请求失败：", code_response)
            return

        # 3. 解密验证码
        encrypted_data = code_response["data"]
        decrypted_code = aes_decrypt_base64(encrypted_data)
        print("解密后的验证码:", decrypted_code)

        # 4. 请求数据
        data_response = fetch_data(decrypted_code, timestamp)
        print("\n返回结果：")
        print(data_response)

    except Exception as e:
        print(f"程序执行出错: {e}")
        exit(1)


if __name__ == "__main__":
    main()
