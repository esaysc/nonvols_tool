import requests
from bs4 import BeautifulSoup
import warnings
warnings.filterwarnings("ignore")

# 读取Cookie
def load_cookie():
    try:
        with open("config.txt", "r", encoding="utf-8") as f:
            return f.read().strip()
    except:
        return ""

COOKIE = load_cookie()

# 要验证的地块
tbid = 138541433
taskid = 1707382

# 接口地址
url = "https://surveysc.iearthtime.com:5088/surveysc/api/restFindTaskAttributeByTaskidForClassify"

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Content-Type": "application/x-www-form-urlencoded",
    "Cookie": COOKIE
}

data = {
    "taskid": taskid,
    "tbid": tbid,
    "version": "36"
}

# 安全获取下拉框选中值
def get_selected(soup, field_id):
    el = soup.find(id=field_id)
    if not el:
        return ""
    opt = el.find("option", selected=True)
    return opt.text.strip() if opt else ""

# 安全获取值
def get_val(soup, field_id):
    el = soup.find(id=field_id)
    return el["value"].strip() if el and el.get("value") else ""

# 请求
resp = requests.post(url, data=data, headers=headers, verify=False)
json_data = resp.json()

if json_data.get("sucessful"):
    html = json_data["message"]
    soup = BeautifulSoup(html, "html.parser")

    print("🔍 验证结果：")
    print("========================================")
    print(f"tbid         : {tbid}")
    print(f"taskid       : {taskid}")
    print("----------------------------------------")
    print(f"矢量报错     : {get_selected(soup, 'field_10818')}")
    print(f"种植属性     : {get_selected(soup, 'field_10745')}")
    print(f"夏收主要作物 : {get_selected(soup, 'field_10746')}")
    print(f"夏收面积     : {get_val(soup, 'field_10821')}")
    print(f"夏收套种     : {get_selected(soup, 'field_10819')}")
    print(f"秋收主要作物 : {get_selected(soup, 'field_10747')}")
    print(f"秋收面积     : {get_val(soup, 'field_10822')}")
    print(f"秋收套种     : {get_selected(soup, 'field_10820')}")
    print("========================================")
    print("✅ 数据已成功保存！")
else:
    print("❌ 查询失败：", json_data.get("message"))