import requests
import base64
from Crypto.Cipher import AES
import time
import socket
from urllib.parse import quote
import random

# 配置常量
HEADERS = {
    "Accept": "application/json",
    "Accept-Encoding": "gzip, deflate, br",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,vi;q=0.7",
    "Connection": "keep-alive",
    "Content-Type": "application/json; charset=utf-8",
    "Host": "106.15.60.27:22222",
    "Referer": "http://106.15.60.27:22222/xxgs/",
    "Sec-Ch-Ua": '"Chromium";v="122", "Not(A:Brand";v="24", "Google Chrome";v="122"',
    "Sec-Ch-Ua-Mobile": "?0",
    "Sec-Ch-Ua-Platform": '"Windows"',
    "Sec-Fetch-Dest": "empty",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Site": "same-origin",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.6261.95 Safari/537.36"
}

RETRY_COUNT = 3               # 请求重试次数
TIMEOUT = 15                  # 请求超时时间（秒）
PAGE_SIZE = 10

def safe_request(session: requests.Session, url: str) -> requests.Response:
    """带自动重试的安全请求"""
    for attempt in range(RETRY_COUNT):
        try:
            if attempt > 0:
                time.sleep(random.uniform(0.5, 2.5))
            response = session.get(url, headers=HEADERS, timeout=TIMEOUT)
            response.raise_for_status()
            return response
        except requests.exceptions.Timeout:
            print(f"↺ 请求超时，正在重试 ({attempt+1}/{RETRY_COUNT})...")
        except requests.exceptions.RequestException as e:
            raise RuntimeError(f"请求失败: {str(e)}")
    raise RuntimeError(f"超过最大重试次数 ({RETRY_COUNT})")

def aes_decrypt_base64(encrypted_base64: str) -> str:
    """AES-CBC解密Base64数据"""
    key = b"6875616E6779696E6875616E6779696E"
    iv = b"sskjKingFree5138"
    encrypted_bytes = base64.b64decode(encrypted_base64)
    cipher = AES.new(key, AES.MODE_CBC, iv)
    decrypted_bytes = cipher.decrypt(encrypted_bytes)
    return decrypted_bytes.rstrip(b'\x00').decode("utf-8")

def create_session() -> requests.Session:
    return requests.Session()

def main():
    print("=== 启动数据获取程序 ===")
    try:
        with create_session() as session:
            timestamp = str(int(time.time() * 1000))
            print(f"[步骤1] 生成时间戳: {timestamp}")

            code_url = f"http://106.15.60.27:22222/ycdc/bakCmisYcOrgan/getCreateCode?codeValue={timestamp}"
            print(f"[步骤2] 请求验证码接口: {code_url}")
            code_response = safe_request(session, code_url).json()
            if code_response.get("code") != 0:
                print(f"[错误] 验证码接口返回异常: {code_response}")
                exit(1)

            encrypted_data = code_response["data"]
            print(f"[步骤3] 解密验证码数据: {encrypted_data[:15]}...")

            decrypted_code = aes_decrypt_base64(encrypted_data)
            print(f"[成功] 解密结果: {decrypted_code}")

            # 获取第一页用于计算总页数
            first_url = (
                "http://106.15.60.27:22222/ycdc/bakCmisYcOrgan/getCurrentIntegrityPage"
                f"?pageSize={PAGE_SIZE}&cioName=%E5%85%AC%E5%8F%B8&page=1"
                f"&code={quote(decrypted_code)}&codeValue={timestamp}"
            )
            print(f"[步骤4] 请求第一页数据: {first_url[:80]}...")
            first_data = safe_request(session, first_url).json()

            total = int(first_data.get("total", 0))
            total_pages = (total + PAGE_SIZE - 1) // PAGE_SIZE
            print(f"[信息] 总记录数: {total}, 总页数: {total_pages}")

            all_data = []
            for page in range(1, total_pages + 1):
                page_url = (
                    "http://106.15.60.27:22222/ycdc/bakCmisYcOrgan/getCurrentIntegrityPage"
                    f"?pageSize={PAGE_SIZE}&cioName=%E5%85%AC%E5%8F%B8&page={page}"
                    f"&code={quote(decrypted_code)}&codeValue={timestamp}"
                )
                print(f"[→] 获取第 {page} 页: {page_url}")
                try:
                    page_data = safe_request(session, page_url).json()

                    print("\n=== 最终数据 ===")
                    print(page_data)

                    print("\n=== 解密结果 ===")
                    print(aes_decrypt_base64(page_data.get("data", "")))

                    all_data.extend(page_data.get("data", []))
                except Exception as e:
                    print(f"[×] 第 {page} 页请求失败: {e}")

            print(f"\n=== 共获取记录数: {len(all_data)} ===")

    except Exception as e:
        print(f"\n!!! 程序执行失败 !!!\n错误原因: {str(e)}")
        print("""
        常见解决方案：
        1. 检查系统时间是否准确
        2. 尝试更换网络环境
        3. 等待5分钟后重试
        4. 检查Cookie是否过期
        """)
        exit(1)

if __name__ == "__main__":
    main()
