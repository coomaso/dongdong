import requests
import base64
from Crypto.Cipher import AES
from Crypto.Util.Padding import unpad
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

AES_KEY = bytes.fromhex("6875616E6779696E6875616E6779696E")  # AES-256密钥
AES_IV = b"sskjKingFree5138"  # 16字节IV
RETRY_COUNT = 3               # 请求重试次数
TIMEOUT = 15                  # 请求超时时间（秒）

def network_check() -> bool:
    """执行基础网络诊断"""
    try:
        # DNS解析检查
        ip = socket.gethostbyname('www.ycjsjg.net')
        print(f"[网络诊断] DNS解析成功 → IP地址: {ip}")
        
        # 端口连通性检查
        with socket.create_connection((ip, 443), timeout=5) as sock:
            print("[网络诊断] 443端口可访问")
            return True
    except Exception as e:
        print(f"[网络诊断] 失败: {str(e)}")
        return False

def create_session() -> requests.Session:
    """创建带持久化Cookie的会话"""
    session = requests.Session()
    # 更新初始Cookie
    session.cookies.update({
        "Hm_lvt_b97569d26a525941d8d163729d284198": "1745837143",
        "Hm_lvt_e8002ef3d9e0d8274b5b74cc4a027d08": "1745837143"
    })
    return session

def safe_request(session: requests.Session, url: str) -> requests.Response:
    """带自动重试的安全请求"""
    for attempt in range(RETRY_COUNT):
        try:
            # 添加随机延迟防止封禁
            if attempt > 0:
                time.sleep(random.uniform(0.5, 2.5))
            
            response = session.get(
                url,
                headers=HEADERS,
                timeout=TIMEOUT,
                # verify=False  # 如需调试SSL可取消注释
            )
            response.raise_for_status()
            return response
        except requests.exceptions.Timeout:
            print(f"↺ 请求超时，正在重试 ({attempt+1}/{RETRY_COUNT})...")
        except requests.exceptions.RequestException as e:
            raise RuntimeError(f"请求失败: {str(e)}")
    raise RuntimeError(f"超过最大重试次数 ({RETRY_COUNT})")

def aes_decrypt_base64(encrypted_base64: str) -> str:
    """AES-CBC解密Base64数据"""
    try:
        encrypted_bytes = base64.b64decode(encrypted_base64)
        cipher = AES.new(AES_KEY, AES.MODE_CBC, AES_IV)
        decrypted_bytes = cipher.decrypt(encrypted_bytes)
        return unpad(decrypted_bytes, AES.block_size).decode("utf-8")
    except (ValueError, KeyError) as e:
        raise RuntimeError(f"解密失败: {str(e)}")

def main():
    print("=== 启动数据获取程序 ===")
    
    # 网络环境检查
    if not network_check():
        print("""
        !!! 网络连接异常 !!!
        可能原因：
        1. 本地网络未连接
        2. 防火墙阻止访问
        3. 目标服务器不可用
        4. DNS解析失败
        建议操作：
        - 检查网络连接
        - 尝试使用手机热点
        - 联系网络管理员
        """)
        exit(1)

    try:
        with create_session() as session:
            # 生成时间戳
            timestamp = str(int(time.time() * 1000))
            print(f"[步骤1] 生成时间戳: {timestamp}")

            # 请求验证码
            code_url = f"https://www.ycjsjg.net/ycdc/bakCmisYcOrgan/getCreateCode?codeValue={timestamp}"
            print(f"[步骤2] 请求验证码接口: {code_url}")
            
            code_response = safe_request(session, code_url).json()
            if code_response.get("code") != 0:
                print(f"[错误] 验证码接口返回异常: {code_response}")
                exit(1)

            # 解密验证码
            encrypted_data = code_response["data"]
            print(f"[步骤3] 解密验证码数据: {encrypted_data[:15]}...")
            
            decrypted_code = aes_decrypt_base64(encrypted_data)
            print(f"[成功] 解密结果: {decrypted_code}")

            # 请求最终数据
            final_url = (
                "https://www.ycjsjg.net/ycdc/bakCmisYcOrgan/getCurrentIntegrityPage"
                f"?pageSize=10&cioName=%E5%85%AC%E5%8F%B8&page=0"
                f"&code={quote(decrypted_code)}&codeValue={timestamp}"
            )
            print(f"[步骤4] 请求数据接口: {final_url[:80]}...")
            
            final_data = safe_request(session, final_url).json()
            
            # 输出结果
            print("\n=== 最终数据 ===")
            print(final_data)

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
