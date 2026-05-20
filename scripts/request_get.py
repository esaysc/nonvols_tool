import requests
from bs4 import BeautifulSoup
import warnings
warnings.filterwarnings("ignore")

# ======================================
# 从外部文件 config.txt 读取 Cookie
# ======================================
def load_cookie():
    try:
        with open("config.txt", "r", encoding="utf-8") as f:
            return f.read().strip()
    except Exception as e:
        print("❌ 未找到 config.txt 文件或读取失败:", e)
        return ""

YOUR_COOKIE = load_cookie()
print(f"已加载 Cookie: {'成功' if YOUR_COOKIE else '未找到或为空'}")

# 接口信息
url = "https://surveysc.iearthtime.com:5088/surveysc/api/restFindTaskAttributeByTaskidForClassify"
data = {
    "taskid": "1707383",
    "tbid": "138541433",
    "version": "36"
}

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Content-Type": "application/x-www-form-urlencoded",
    "Cookie": YOUR_COOKIE
}

# 请求
response = requests.post(url, data=data, headers=headers, verify=False)
json_data = response.json()

# 安全获取 selected 文本（修复崩溃问题）
def get_selected_text(soup, field_id):
    el = soup.find(id=field_id)
    if not el:
        return ""
    opt = el.find("option", selected=True)
    return opt.text.strip() if opt else ""

# 安全获取 value
def get_value(soup, field_id):
    el = soup.find(id=field_id)
    return el["value"].strip() if el and el.get("value") else ""

if json_data.get("sucessful"):
    html = json_data["message"]
    soup = BeautifulSoup(html, "html.parser")
    print("soup")
    result = {
        "tbid": get_value(soup, "edit_tbid"),
        "taskid": get_value(soup, "edit_taskid"),
        "地块编号": soup.find("td", string="地块编号").find_next("td").text.strip() if soup.find("td", string="地块编号") else "",
        "矢量报错": get_selected_text(soup, "field_10818"),
        "农作物种植用地属性": get_selected_text(soup, "field_10745"),
        "2025年夏收主要作物": get_selected_text(soup, "field_10746"),
        "夏收主要作物面积(%)": get_value(soup, "field_10821"),
        "夏收作物套种": get_selected_text(soup, "field_10819"),
        "夏收次要作物": get_selected_text(soup, "field_10756"),
        "夏收次要作物面积(%)": get_value(soup, "field_10748"),
        "2025年秋收主要作物": get_selected_text(soup, "field_10747"),
        "秋收主要作物面积(%)": get_value(soup, "field_10822"),
        "秋收作物套种": get_selected_text(soup, "field_10820"),
        "秋收次要作物": get_selected_text(soup, "field_10757"),
        "秋收次要作物面积(%)": get_value(soup, "field_10751"),
        "备注": get_value(soup, "field_10530")
    }

    print("=" * 60)
    print("✅ 地块全部信息获取成功")
    print("=" * 60)
    for k, v in result.items():
        print(f"{k:>20} : {v}")

else:
    print("❌ 失败：", json_data.get("message"))