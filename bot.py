import os
import json
import requests
from glob import glob

# ======= 企业微信机器人 Webhook 地址，从环境变量读取 =======
WEBHOOK_URL = os.getenv("WEBHOOK_URL")
if not WEBHOOK_URL:
    raise ValueError("环境变量 WEBHOOK_URL 未设置，请设置你的企业微信机器人 Webhook 地址")

# ======= 获取最新的文件路径 =======
def get_latest_file(directory, pattern):
    files = glob(os.path.join(directory, pattern))
    if not files:
        return None
    return max(files, key=os.path.getmtime)

# ======= 发送文本消息，企业名包含“盛荣”则红色高亮 =======
def send_text_msg(title, data_list):
    content = f"**📊 {title}**\n\n"
    # 企业微信支持简单 markdown，红色字体用 > <font color="red">text</font>
    for item in data_list:
        name = item['企业名称']
        score = item['诚信分值']
        rank = item['排名']
        if "盛荣" in name:
            # 红色高亮
            line = f"{rank}. > **<font color=\"red\">{name}</font>** （{score}分）\n"
        else:
            line = f"{rank}. {name}（{score}分）\n"
        content += line

    payload = {
        "msgtype": "markdown",
        "markdown": {
            "content": content
        }
    }
    r = requests.post(WEBHOOK_URL, json=payload)
    r.raise_for_status()

# ======= 上传文件到企业微信并发送 =======
def send_file_msg(filepath):
    filename = os.path.basename(filepath)
    # 获取媒体 ID
    upload_url = f"{WEBHOOK_URL}&type=file"
    with open(filepath, 'rb') as f:
        files = {'file': (filename, f, 'application/octet-stream')}
        res = requests.post(upload_url, files=files)
    res.raise_for_status()
    media_id = res.json().get("media_id")
    if not media_id:
        print(f"❌ 上传文件失败：{res.text}")
        return

    # 发送文件消息
    payload = {
        "msgtype": "file",
        "file": {
            "media_id": media_id
        }
    }
    r = requests.post(WEBHOOK_URL, json=payload)
    r.raise_for_status()
    print(f"✅ 文件已发送: {filename}")

# ======= 主函数 =======
def main():
    output_dir = "excel_output"

    # 获取最新的 JSON 文件并读取内容
    json_file = get_latest_file(output_dir, "*.json")
    if not json_file:
        print("未找到 JSON 文件")
        return

    with open(json_file, "r", encoding="utf-8") as f:
        json_data = json.load(f)

    timestamp_raw = json_data.get("TIMEamp", None)
    if timestamp_raw:
        dt = datetime.strptime(timestamp_raw, "%Y%m%d_%H%M%S")
        dt_bj = dt + timedelta(hours=8)
        timestamp = dt_bj.strftime("%Y-%m-%d")
    else:
        timestamp = "未知时间"
        
    title = f"宜昌市企业诚信分值 Top10（{timestamp}）"
    data_list = json_data.get("DATAlist")

    if not data_list:
        print("JSON 内容为空或格式不正确")
        return

    # 发送文本消息
    send_text_msg(title, data_list)

    # 只发送最新的 XLSX 附件
    filepath = get_latest_file(output_dir, "*.xlsx")
    if filepath:
        send_file_msg(filepath)

if __name__ == "__main__":
    main()
