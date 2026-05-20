import requests
import json
import time
import warnings
warnings.filterwarnings("ignore")

# ===================== 配置区 =====================
def load_cookie():
    try:
        with open("config.txt", "r", encoding="utf-8") as f:
            return f.read().strip()
    except:
        print("❌ 请创建 config.txt 并放入Cookie")
        return ""

COOKIE = load_cookie()

# 要提交的地块列表
TB_TASK_LIST = [
    (138541434, 1707382),
    (138541433, 1707382),
]

# 提交内容
FILL_DATA = {
    "10818": "否",
    "10745": "旱地",
    "10746": "豆类",
    "10821": 100,
    "10819": "否",
    "10756": None,
    "10748": None,
    "10747": "甘薯",
    "10822": 100,
    "10820": "否",
    "10757": None,
    "10751": None,
    "10530": None
}

# 接口
URL = "https://surveysc.iearthtime.com:5088/surveysc/api/restUploadFieldscreenForTB?version=36&hasError=false&skipResponse=true"

# 请求头
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Content-Type": "application/json;charset=UTF-8",
    "Cookie": COOKIE,
    "Referer": "https://surveysc.iearthtime.com:5088/surveysc/mapClassifyTaskDetail.html",
    "Origin": "https://surveysc.iearthtime.com:5088",
    "Accept": "application/json, text/javascript, */*",
    "X-Requested-With": "XMLHttpRequest"
}

# ===================== 构造请求体 =====================
def build_post_data(tbid, taskid):
    arr = []
    base = {"compid": 10046, "taskid": int(taskid), "tbid": int(tbid)}

    field_map = {
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
        "10530": "string"
    }

    for cf, vt in field_map.items():
        item = {"customfield": int(cf), "valuetype": vt, **base}
        val = FILL_DATA[cf]

        if vt == "string":
            item["stringvalue"] = val
        if vt == "number":
            item["numbervalue"] = val
        arr.append(item)
    return arr

# ===================== 延迟批量提交 =====================
def batch_submit():
    print("🚀 开始延迟批量提交（每个间隔 1.5 秒）...\n")
    s = requests.Session()

    for index, (tbid, taskid) in enumerate(TB_TASK_LIST, 1):
        data = build_post_data(tbid, taskid)
        
        print(f"===== 正在提交第 {index} 个 | tbid={tbid}, taskid={taskid} =====")
        try:
            resp = s.post(URL, json=data, headers=HEADERS, verify=False, timeout=15)
            print(f"状态码: {resp.status_code}")
            print(f"返回: {resp.text}")

            if resp.status_code == 200 and "sucessful" in resp.text:
                print("✅ 提交成功！\n")
            else:
                print("❌ 提交失败！\n")

        except Exception as e:
            print(f"❌ 异常: {str(e)}\n")

        # 延迟 1.5 秒再提交下一个（防封、防失败）
        time.sleep(1.5)

    print("🎉 全部批量提交完成！")

if __name__ == "__main__":
    batch_submit()