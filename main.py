import requests
import base64
from Crypto.Cipher import AES
from Crypto.Util.Padding import unpad
import time
from urllib.parse import quote

# 配置常量
HEADERS = {
    "Accept": "application/json",
    "Accept-Encoding": "gzip, deflate, br",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,vi;q=0.7",
    "Connection": "keep-alive",
    "Content-Type": "application/json; charset=utf-8",
    "Cookie": "Hm_lvt_b97569d26a525941d8d163729d284198=1745837143; Hm_lvt_e8002ef3d9e0d8274b5b74cc4a027d08=1745837143",
    "Host": "www.ycjsjg.net",
    "Referer": "https://www.ycjsjg.net/xxgs/",
    "Sec-Ch-Ua": '"Chromium";v="122", "Not(A:Brand";v="24", "Google Chrome";v="122"',
    "Sec-Ch-Ua-Mobile": "?0",
    "Sec-Ch-Ua-Platform": '"Windows"',
    "Sec-Fetch-Dest": "empty",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Site": "same-origin",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.6261.95 Safari/537.36"
}

AES_KEY = bytes.fromhex("6875616E6779696E6875616E6779696E")  # AES-256 密钥
AES_IV = b"sskjKingFree5138"  # 16字节IV

def aes_decrypt_base64(encrypted_base64: str) -> str:
    """AES-CBC解密Base64编码数据"""
    try:
        encrypted_bytes = base64.b64decode(encrypted_base64)
        cipher = AES.new(AES_KEY, AES.MODE_CBC, AES_IV)
        decrypted_bytes = cipher.decrypt(encrypted_bytes)
        return unpad(decrypted_bytes, AES.block_size).decode("utf-8")
    except (ValueError, KeyError) as e:
        raise RuntimeError(f"解密失败: {e}")

def get_timestamp() -> str:
    """获取当前毫秒时间戳"""
    return str(int(time.time() * 1000))

def create_session() -> requests.Session:
    """创建带持久化Cookie的会话"""
    session = requests.Session()
    # 手动设置初始Cookie
    cookies = {
        "Hm_lvt_b97569d26a525941d8d163729d284198": "1745837143",
        "Hm_lvt_e8002ef3d9e0d8274b5b74cc4a027d08": "1745837143"
    }
    session.cookies.update(requests.utils.cookiejar_from_dict(cookies))
    return session

def fetch_code(session: requests.Session, timestamp: str) -> dict:
    """请求验证码接口"""
    url = f"https://www.ycjsjg.net/ycdc/bakCmisYcOrgan/getCreateCode?codeValue={timestamp}"
    try:
        response = session.get(url, headers=HEADERS, timeout=10)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"验证码接口请求失败: {e}")

def fetch_data(session: requests.Session, code: str, timestamp: str) -> dict:
    """请求最终数据接口"""
    final_url = (
        "https://www.ycjsjg.net/ycdc/bakCmisYcOrgan/getCurrentIntegrityPage"
        f"?pageSize=10&cioName=%E5%85%AC%E5%8F%B8&page=0"
        f"&code={quote(code)}&codeValue={timestamp}"
    )
    try:
        response = session.get(final_url, headers=HEADERS, timeout=10)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"数据接口请求失败: {e}")

def main():
    try:
        # 创建持久化会话
        with create_session() as session:
            # 1. 获取时间戳
            timestamp = get_timestamp()
            print(f"生成时间戳: {timestamp}")

            # 2. 获取验证码
            code_response = fetch_code(session, timestamp)
            if code_response.get("code") != 0:
                print("验证码请求失败：", code_response)
                return

            # 3. 解密验证码
            encrypted_data = code_response["data"]
            decrypted_code = aes_decrypt_base64(encrypted_data)
            print("解密后的验证码:", decrypted_code)

            # 4. 请求数据
            data_response = fetch_data(session, decrypted_code, timestamp)
            print("\n返回结果：")
            print(data_response)

    except Exception as e:
        print(f"程序执行出错: {e}")
        exit(1)

if __name__ == "__main__":
    main()
