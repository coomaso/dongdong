import requests
import base64
import json
from Crypto.Cipher import AES
import time
from urllib.parse import quote
import random
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

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
            print(f"正在请求: {url}")  # 添加请求URL日志
            response = session.get(url, headers=HEADERS, timeout=TIMEOUT)
            response.raise_for_status()
            return response
        except requests.exceptions.Timeout:
            print(f"↺ 请求超时，正在重试 ({attempt+1}/{RETRY_COUNT})...")
        except requests.exceptions.RequestException as e:
            print(f"请求异常: {str(e)}")  # 打印具体异常信息
            if attempt < RETRY_COUNT - 1:
                print(f"正在进行第 {attempt+2} 次尝试...")
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
        print(f"解密失败，原始数据: {encrypted_base64[:50]}...")  # 打印部分原始数据
        raise RuntimeError(f"解密失败: {str(e)}")

def get_new_code(session: requests.Session) -> tuple:
    """获取新验证码和时间戳"""
    timestamp = str(int(time.time() * 1000))
    code_url = f"http://106.15.60.27:22222/ycdc/bakCmisYcOrgan/getCreateCode?codeValue={timestamp}"

    try:
        response = safe_request(session, code_url).json()
        print(f"验证码接口响应: {json.dumps(response, ensure_ascii=False)[:100]}...")  # 打印部分响应
        if response.get("code") != 0:
            raise RuntimeError(f"验证码接口异常: {response}")
        return aes_decrypt_base64(response["data"]), timestamp
    except Exception as e:
        print(f"获取验证码失败，URL: {code_url}")  # 打印失败的URL
        raise RuntimeError(f"获取新验证码失败: {str(e)}")

def parse_response_data(encrypted_data: str) -> dict:
    """健壮的数据解析方法"""
    if not encrypted_data:
        print("警告: 收到空的加密数据")  # 添加警告日志
        return {"error": "empty data"}

    try:
        decrypted_str = aes_decrypt_base64(encrypted_data)
        print(f"解密后的数据样本: {decrypted_str[:100]}...")  # 打印解密后的数据样本
        return json.loads(decrypted_str)
    except json.JSONDecodeError as e:
        print(f"JSON解析错误，数据样本: {encrypted_data[:50]}...")  # 打印错误数据样本
        return {"error": f"invalid json format: {str(e)}"}
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
        print(f"第 {page} 页API响应状态: {page_response.get('code', '未知')}")  # 打印响应状态

        if "data" not in page_response or not page_response["data"]:
            print(f"第 {page} 页数据为空，触发验证码刷新")
            raise RuntimeError("empty response data")

        page_data = parse_response_data(page_response["data"])
        
        if "error" in page_data:
            print(f"第 {page} 页数据解析错误: {page_data['error']}")
            raise RuntimeError(f"invalid page data: {page_data['error']}")
        
        records = page_data.get("data", [])
        print(f"第 {page} 页解析出 {len(records)} 条记录")  # 明确记录数量
        
        # 检查解析出的数据是否有效
        if not records:
            print(f"警告: 第 {page} 页解析出空记录列表")
            
        return records, page_data.get("total", 0)
    except Exception as e:
        print(f"第 {page} 页处理失败: {str(e)}")
        raise

def export_to_excel(data, filename="企业数据.xlsx"):
    """增强版Excel导出函数(支持复杂数据结构)"""
    
    # ==================== 前置校验 ====================
    if not data:
        print("错误: 输入数据为空，取消导出操作")
        return
    
    # ==================== 数据预处理 ====================
    processed_data = []
    error_count = 0
    
    for idx, item in enumerate(data, 1):
        try:
            # 类型安全检查
            if not isinstance(item, dict):
                print(f"警告: 第{idx}条数据不是字典格式，已跳过")
                error_count += 1
                continue
            
            clean_item = {}
            for key, value in item.items():
                # 处理不同类型数据
                if isinstance(value, (list, dict)):
                    # 将复杂结构转为JSON字符串
                    clean_item[key] = json.dumps(value, ensure_ascii=False) if value else ""
                elif isinstance(value, (int, float)):
                    # 数值类型直接保留
                    clean_item[key] = value
                elif value is None:
                    # 空值处理
                    clean_item[key] = ""
                else:
                    # 其他类型转为字符串
                    clean_item[key] = str(value)
            
            # 检查是否存在有效数据
            if any(clean_item.values()):
                processed_data.append(clean_item)
            else:
                print(f"警告: 第{idx}条数据所有字段均为空，已跳过")
                error_count += 1
                
        except Exception as e:
            print(f"处理第{idx}条数据时发生错误: {str(e)}")
            error_count += 1
    
    # ==================== 有效性检查 ====================
    if not processed_data:
        print(f"错误: 所有{len(data)}条数据均无效，取消导出")
        return
    
    print(f"数据预处理完成，有效数据: {len(processed_data)}条，无效数据: {error_count}条")
    
    # ==================== 创建Excel ====================
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "企业信息"
    except Exception as e:
        print(f"创建Excel工作簿失败: {str(e)}")
        return
    
    # ==================== 生成动态表头 ====================
    header_set = set()
    for item in processed_data:
        header_set.update(item.keys())
    
    # 排序表头(可选)
    headers = sorted(header_set, key=lambda x: (
        # 自定义排序逻辑
        ("cioName", "eqtName", "csf").index(x) if x in ("cioName", "eqtName", "csf") else 999,
        x
    ))
    
    # ==================== 样式配置 ====================
    # 表头样式
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )
    
    # 数据样式
    data_alignment = Alignment(vertical="center", wrap_text=True)
    
    # ==================== 写入数据 ====================
    try:
        # 添加表头
        ws.append(headers)
        
        # 设置表头样式
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            cell.border = thin_border
        
        # 写入数据行
        for row_idx, item in enumerate(processed_data, 2):
            row = []
            for header in headers:
                # 安全获取值
                value = item.get(header, "")
                
                # 处理超长内容
                if isinstance(value, str) and len(value) > 32767:
                    value = value[:30000] + "...[数据截断]"
                
                row.append(value)
            
            ws.append(row)
            
            # 设置数据行样式
            for col_idx in range(1, len(headers)+1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.alignment = data_alignment
                cell.border = thin_border
                
                # 自动识别数值类型
                if isinstance(cell.value, (int, float)):
                    cell.number_format = "0.00"
    except Exception as e:
        print(f"写入Excel数据时发生错误: {str(e)}")
        return
    
    # ==================== 格式优化 ====================
    try:
        # 自动调整列宽
        for col_idx, header in enumerate(headers, 1):
            max_length = 0
            column = get_column_letter(col_idx)
            
            # 计算最大长度
            for cell in ws[column]:
                try:
                    length = len(str(cell.value))
                    if length > max_length:
                        max_length = length
                except:
                    pass
            
            # 设置列宽(限制最大50字符)
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column].width = adjusted_width
        
        # 添加标题行
        title = f"企业信用数据 ({time.strftime('%Y-%m-%d')})"
        ws.insert_rows(1)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
        title_cell = ws.cell(row=1, column=1, value=title)
        title_cell.font = Font(size=14, bold=True)
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        
        # 冻结表头
        ws.freeze_panes = "A2"
    except Exception as e:
        print(f"格式优化时发生错误: {str(e)}")
    
    # ==================== 保存文件 ====================
    try:
        # 检查文件名后缀
        if not filename.endswith(".xlsx"):
            filename += ".xlsx"
            
        # 检查目录是否存在
        os.makedirs(os.path.dirname(filename), exist_ok=True)
        
        # 保存文件
        wb.save(filename)
        print(f"成功导出文件: {os.path.abspath(filename)}")
        print(f"总数据量: {len(processed_data)}条")
        print(f"列字段: {', '.join(headers)}")
    except PermissionError:
        print(f"错误: 文件 {filename} 被其他程序占用，请关闭后重试")
    except Exception as e:
        print(f"保存文件失败: {str(e)}")
        
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

        if total == 0:
            print("错误: API返回总记录数为0，无需继续处理")
            return

        # 分页处理
        page = 1
        while page <= total_pages:
            retry_count = 0
            success = False

            while retry_count < PAGE_RETRY_MAX and not success:
                try:
                    print(f"\n[处理中] 第 {page} 页 (重试次数: {retry_count})")
                    page_data, _ = process_page(session, page, current_code, current_ts)
                    
                    if page_data:
                        print(f"[成功获取数据] 第 {page} 页 {len(page_data)} 条记录")
                        all_data.extend(page_data)
                        success = True
                        page += 1
                    else:
                        print(f"[警告] 第 {page} 页获取到空数据，尝试刷新验证码")
                        raise RuntimeError("empty page data")

                except Exception as e:
                    retry_count += 1
                    print(f"[重试] 第 {page} 页第 {retry_count} 次重试: {str(e)}")

                    # 获取新验证码
                    try:
                        current_code, current_ts = get_new_code(session)
                        print(f"[刷新] 新验证码: {current_code} | 新时间戳: {current_ts}")
                    except Exception as e:
                        print(f"[警告] 验证码刷新失败: {str(e)}")
                        break

            if not success:
                print(f"[终止] 第 {page} 页超过最大重试次数，跳过此页")
                page += 1  # 跳过失败页

        print(f"\n=== 数据获取完成 ===")
        print(f"总获取记录数: {len(all_data)}")
        
        # 导出数据前再次检查
        if all_data:
            print("示例数据:", json.dumps(all_data[:2], indent=2, ensure_ascii=False))
            export_to_excel(all_data)
        else:
            print("错误: 没有获取到任何有效数据，无法导出Excel")

    except Exception as e:
        print(f"\n!!! 程序执行失败 !!!\n错误原因: {str(e)}")
    finally:
        session.close()

if __name__ == "__main__":
    main()
