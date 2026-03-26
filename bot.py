import os
import json
import requests
from datetime import datetime, timedelta
import urllib.parse

# ======= 企业微信机器人 Webhook 地址，从环境变量读取 =======
WEBHOOK_URL = os.getenv("WEBHOOK_URL")
if not WEBHOOK_URL:
    raise ValueError("环境变量 WEBHOOK_URL 未设置，请设置你的企业微信机器人 Webhook 地址")

# ======= 从 WEBHOOK_URL 中提取 key =======
def get_key_from_webhook(url):
    parsed = urllib.parse.urlparse(url)
    query = urllib.parse.parse_qs(parsed.query)
    key = query.get('key', [None])[0]
    if not key:
        raise ValueError("无法从 WEBHOOK_URL 中提取 key")
    return key

# ======= 发送文本消息，企业名包含“盛荣”则红色高亮 =======
def send_text_msg(title, data_list):
    content = f"## 🏆 {title}\n\n"
    for item in data_list:
        name = item['企业名称']
        score = item['诚信分值']
        rank = item['排名']
        if "盛荣" in name:
            line = f'> **<font color="red">{rank}. {name} {score}</font>**\n'
        else:
            line = f'> {rank}. {name} <font color="comment">{score}</font>\n'
        content += line

    payload = {
        "msgtype": "markdown",
        "markdown": {"content": content}
    }
    r = requests.post(WEBHOOK_URL, json=payload)
    r.raise_for_status()
    print("✅ 文本消息已发送")

# ======= 主函数 =======
def main():
    output_dir = "excel_output"
    # 指定文件名
    filename = "建筑工程总承包信用分排序_top10.json"
    filepath = os.path.join(output_dir, filename)

    if not os.path.isfile(filepath):
        print(f"❌ 文件不存在：{filepath}")
        return

    with open(filepath, "r", encoding="utf-8") as f:
        json_data = json.load(f)

    timestamp_raw = json_data.get("TIMEamp", None)
    if timestamp_raw:
        dt = datetime.strptime(timestamp_raw, "%Y%m%d_%H%M%S")
        dt_bj = dt + timedelta(hours=8)  # 转北京时间
        timestamp = dt_bj.strftime("%Y-%m-%d")
    else:
        timestamp = "未知时间"

    title = f"宜昌施工总承包诚信分 Top10 {timestamp}"
    data_list = json_data.get("DATAlist")

    if not data_list:
        print("JSON 内容为空或格式不正确")
        return

    # 发送文本消息
    send_text_msg(title, data_list)

if __name__ == "__main__":
    main()