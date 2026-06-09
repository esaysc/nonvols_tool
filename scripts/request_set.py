import requests
import json
import time
import pandas as pd
import sys
import warnings
import random
import os

from requests.exceptions import RequestException, Timeout, SSLError
from urllib3.exceptions import InsecureRequestWarning

# 正确过滤业务无关的警告
warnings.filterwarnings("ignore", category=DeprecationWarning)
warnings.filterwarnings("ignore", category=UserWarning)
warnings.filterwarnings("ignore", category=InsecureRequestWarning)

# ===================== 基础配置 =====================
COMP_ID = 10046
API_URL_BASE = "https://surveysc.iearthtime.com:5088/surveysc/api/restUploadFieldscreenForTB?version=36&skipResponse=true"
QUERY_API_URL = "https://surveysc.iearthtime.com:5088/surveysc/api/restFindTaskAttributeByTaskidForClassify"
EXCEL_FILE = "data/new_data.xlsx"

FIELD_TYPE_MAP = {
    "10818": "string",
    "10745": "string",
    "10746": "string",
    "10821": "number",
    "10819": "string",
    "10756": "string",
    "10748": "number",
    "10747": "string",
    "10822": "number",
    "10820": "string",
    "10757": "string",
    "10751": "number",
    "10530": "string",
}

EXCEL_COLUMN_TO_CUSTOMFIELD = {
    "矢量报错": "10818",
    "农作物种植用地属性": "10745",
    "2025年夏收主要作物": "10746",
    "夏收主要作物面积（%）": "10821",
    "夏收作物套种": "10819",
    "夏收次要作物": "10756",
    "夏收次要作物面积（%）": "10748",
    "2025年秋收主要作物": "10747",
    "秋收主要作物面积（%）": "10822",
    "秋收作物套种": "10820",
    "秋收次要作物": "10757",
    "秋收次要作物面积（%）": "10751",
    "备注": "10530",
}

REQUEST_HEADERS = {
    "Accept": "*/*",
    "Accept-Encoding": "gzip, deflate, br, zstd",
    "Accept-Language": "zh-CN,zh;q=0.9",
    "Connection": "keep-alive",
    "Content-Type": "application/json; charset=UTF-8",
    "Host": "surveysc.iearthtime.com:5088",
    "Origin": "https://surveysc.iearthtime.com:5088",
    "Referer": "https://surveysc.iearthtime.com:5088/surveysc/mapClassifyTaskDetail.html",
    "sec-ch-ua": '"Chromium";v="148", "Google Chrome";v="148", "Not/A)Brand";v="99"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"Windows"',
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "same-origin",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36",
    "X-Requested-With": "XMLHttpRequest",
}

QUERY_HEADERS = REQUEST_HEADERS.copy()
QUERY_HEADERS["Content-Type"] = "application/x-www-form-urlencoded; charset=UTF-8"

# ===================== 1. 加载Excel =====================
def load_excel_data(excel_path):
    if not os.path.exists(excel_path):
        print(f"❌ Excel 文件不存在：{excel_path}")
        sys.exit(1)

    try:
        df = pd.read_excel(excel_path, dtype=str)
    except Exception as e:
        print(f"❌ Excel 读取失败：{str(e)}")
        sys.exit(1)

    required_cols = ["taskid", "tbid"]
    for col in required_cols:
        if col not in df.columns:
            print(f"❌ 缺少列：{col}")
            sys.exit(1)

    records = []
    invalid_rows = []
    for idx, row in df.iterrows():
        try:
            tbid = int(float(row["tbid"]))
            taskid = int(float(row["taskid"]))
        except:
            invalid_rows.append(idx + 2)
            continue

        row_data = {}
        for excel_col, field_code in EXCEL_COLUMN_TO_CUSTOMFIELD.items():
            if excel_col not in df.columns:
                row_data[field_code] = None
                continue
            raw_val = row[excel_col]
            if pd.isna(raw_val) or str(raw_val).strip() == "" or str(raw_val).strip().lower() == "nan":
                cleaned = None
            else:
                cleaned = str(raw_val).strip()
                if FIELD_TYPE_MAP[field_code] == "number" and cleaned.isdigit():
                    cleaned = int(cleaned)
            row_data[field_code] = cleaned

        records.append((tbid, taskid, row_data))

    if not records:
        print("❌ 无有效数据")
        sys.exit(1)
    print(f"✅ 加载完成：{len(records)} 个地块")
    return records

# ===================== 2. Cookie =====================
def load_and_validate_cookie():
    try:
        with open("config.txt", "r", encoding="utf-8") as f:
            cleaned_cookie = f.read().replace("\n", "").replace("\r", "").strip()
        if not cleaned_cookie:
            print("❌ cookie 为空")
            sys.exit(1)
        print("✅ cookie 加载完成")
        return cleaned_cookie
    except:
        print("❌ 请创建 config.txt 并填入cookie")
        sys.exit(1)

# ===================== 3. 安全确认 =====================
def security_confirm(records):
    print("\n" + "="*60)
    print("🔒 安全确认环节")
    print(f"📌 数据来源：{EXCEL_FILE}")
    print(f"📋 总地块数：{len(records)}")
    confirm = input("\n输入 y/yes 开始提交：").strip().lower()
    if confirm not in ["y", "yes"]:
        print("❌ 已退出")
        sys.exit(0)

# ===================== 4. 查询接口 =====================
def query_task_attribute(session, tbid, taskid, cookie):
    try:
        headers = QUERY_HEADERS.copy()
        headers["Cookie"] = cookie
        data = {"taskid": taskid, "tbid": tbid, "version": 36}
        res = session.post(QUERY_API_URL, headers=headers, data=data, verify=False, timeout=15)
        if res.status_code == 200:
            print("✅ 读取任务属性成功")
    except:
        pass

# ===================== 5. 构造请求体 =====================
def build_request_body(tbid, taskid, row_data):
    body = []
    base = {"compid": COMP_ID, "taskid": taskid, "tbid": tbid}
    for code, vt in FIELD_TYPE_MAP.items():
        item = {"customfield": int(code), "valuetype": vt, **base}
        val = row_data.get(code)
        if vt == "string":
            item["stringvalue"] = val
        else:
            item["numbervalue"] = val
        body.append(item)
    return body

# ===================== 6. 批量提交 =====================
def batch_submit(records, cookie):
    success = []
    fail = []
    total = len(records)

    for i, (tbid, taskid, data) in enumerate(records, 1):
        print(f"\n[{i}/{total}] 提交 tbid={tbid} taskid={taskid}")

        # 每个地块独立会话（模拟真人）
        session = requests.Session()
        headers = REQUEST_HEADERS.copy()
        headers["Cookie"] = cookie

        # 查询接口
        query_task_attribute(session, tbid, taskid, cookie)
        wait_q = random.randint(1,2)
        time.sleep(1)

        # 构造提交体
        body = build_request_body(tbid, taskid, data)

        # 动态 hasError
        vector_err = data.get("10818")
        hasError = "true" if vector_err == "是" else "false"
        url = f"{API_URL_BASE}&hasError={hasError}"
        print(f"🔗 提交URL：{hasError}")

        ok = False
        msg = ""
        try:
            resp = session.post(url, json=body, headers=headers, verify=False, timeout=25)
            if resp.status_code == 200:
                try:
                    j = resp.json()
                    if j.get("sucessful") or j.get("successful"):
                        ok = True
                except:
                    ok = True
        except Exception as e:
            msg = str(e)

        if ok:
            print("✅ 提交成功")
            success.append((tbid, taskid))
        else:
            print(f"❌ 失败：{msg}")
            fail.append((tbid, taskid, msg))

        time.sleep(1)

    # 结果
    print("\n" + "="*60)
    print(f"📊 完成：总计{total} 成功{len(success)} 失败{len(fail)}")
    print("="*60)

# ===================== 主入口 =====================
if __name__ == "__main__":
    cookie = load_and_validate_cookie()
    records = load_excel_data(EXCEL_FILE)
    security_confirm(records)
    batch_submit(records, cookie)
    print("\n🔹 程序结束")