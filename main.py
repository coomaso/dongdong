import requests
import base64
import json
from Crypto.Cipher import AES
import time
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
PAGE_RETRY_MAX = 2           # 单页最大重试次数
TIMEOUT = 15                  # 请求超时时间（秒）
PAGE_SIZE = 10

# AES配置
AES_KEY = b"6875616E6779696E6875616E6779696E"
AES_IV = b"sskjKingFree5138"

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
    """增强版AES解密函数"""
    if not encrypted_base64:
        raise ValueError("加密数据为空，无法解密")
    
    try:
        encrypted_bytes = base64.b64decode(encrypted_base64)
        cipher = AES.new(AES_KEY, AES.MODE_CBC, AES_IV)
        decrypted_bytes = cipher.decrypt(encrypted_bytes)
        return decrypted_bytes.rstrip(b'\x00').decode("utf-8")
    except Exception as e:
        raise RuntimeError(f"解密失败: {str(e)}")

def get_new_code(session: requests.Session) -> tuple:
    """获取新验证码和时间戳"""
    timestamp = str(int(time.time() * 1000))
    code_url = f"http://106.15.60.27:22222/ycdc/bakCmisYcOrgan/getCreateCode?codeValue={timestamp}"
    
    try:
        response = safe_request(session, code_url).json()
        if response.get("code") != 0:
            raise RuntimeError(f"验证码接口异常: {response}")
        return aes_decrypt_base64(response["data"]), timestamp
    except Exception as e:
        raise RuntimeError(f"获取新验证码失败: {str(e)}")

def parse_response_data(encrypted_data: str) -> dict:
    """健壮的数据解析方法"""
    if not encrypted_data:
        return {"error": "empty data"}
    
    try:
        decrypted_str = aes_decrypt_base64(encrypted_data)
        return json.loads(decrypted_str)
    except json.JSONDecodeError:
        return {"error": "invalid json format"}
    except Exception as e:
        return {"error": str(e)}

def process_page(session: requests.Session, page: int, code: str, timestamp: str) -> tuple:
    """处理单个页面并返回数据"""
    page_url = (
        "http://106.15.60.27:22222/ycdc/bakCmisYcOrgan/getCurrentIntegrityPage"
        f"?pageSize={PAGE_SIZE}&cioName=%E5%85%AC%E5%8F%B8&page={page}"
        f"&code={quote(code)}&codeValue={timestamp}"
    )
    
    try:
        response = safe_request(session, page_url)
        page_response = response.json()
        
        if "data" not in page_response or not page_response["data"]:
            print(f"第 {page} 页数据为空，触发验证码刷新")
            raise RuntimeError("empty response data")
            
        page_data = parse_response_data(page_response["data"])
        print(page_data)
        if "error" in page_data:
            print(f"第 {page} 页数据解析错误: {page_data['error']}")
            raise RuntimeError("invalid page data")
            
        return page_data.get("rows", []), page_data.get("total", 0)
    except Exception as e:
        print(f"第 {page} 页处理失败: {str(e)}")
        raise

def main():
    print("=== 启动数据获取程序 ===")
    session = requests.Session()
    all_data = []
    
    try:
        # 初始获取验证码
        current_code, current_ts = get_new_code(session)
        print(f"[初始化] 验证码: {current_code} | 时间戳: {current_ts}")
        
        # 获取第一页确定总数
        first_data, total = process_page(session, 1, current_code, current_ts)
        total_pages = (total + PAGE_SIZE - 1) // PAGE_SIZE
        print(f"[初始化] 总记录数: {total} | 总页数: {total_pages}")
        
        # 分页处理
        page = 1
        while page <= total_pages:
            retry_count = 0
            success = False
            
            while retry_count < PAGE_RETRY_MAX and not success:
                try:
                    print(f"\n[处理中] 第 {page} 页 (重试次数: {retry_count})")
                    page_data, _ = process_page(session, page, current_code, current_ts)
                    print(page_data)
                    all_data.extend(page_data['data'])
                    print(f"[成功] 第 {page} 页获取到 {len(page_data['data'])} 条记录")
                   
                    success = True
                    page += 1
                    
                except Exception as e:
                    retry_count += 1
                    print(f"[重试] 第 {page} 页第 {retry_count} 次重试")
                    
                    # 获取新验证码
                    try:
                        current_code, current_ts = get_new_code(session)
                        print(f"[刷新] 新验证码: {current_code} | 新时间戳: {current_ts}")
                    except Exception as e:
                        print(f"[警告] 验证码刷新失败: {str(e)}")
                        break
                    
            if not success:
                print(f"[终止] 第 {page} 页超过最大重试次数")
                page += 1  # 跳过失败页

        print(f"\n=== 数据获取完成 ===")
        print(f"总获取记录数: {len(all_data)}")
        print("示例数据:", json.dumps(all_data[:2], indent=2, ensure_ascii=False))

    except Exception as e:
        print(f"\n!!! 程序执行失败 !!!\n错误原因: {str(e)}")
    finally:
        session.close()

if __name__ == "__main__":
    main()
