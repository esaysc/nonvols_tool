import requests
import json
import warnings
warnings.filterwarnings("ignore")

# ===================== 【固定你的目标地块】 =====================
COOKIE = open("config.txt", "r", encoding="utf-8").read().strip()

tbid = 138541432
taskid = 1707383

# 【严格按照你要的空值结构】
payload = [
    {"customfield":10818,"valuetype":"string","stringvalue":"否","compid":10046,"taskid":taskid,"tbid":tbid},
    {"customfield":10745,"valuetype":"string","stringvalue":"","compid":10046,"taskid":taskid,"tbid":tbid},
    {"customfield":10746,"valuetype":"string","stringvalue":"","compid":10046,"taskid":taskid,"tbid":tbid},
    {"customfield":10821,"valuetype":"number","numbervalue":100,"compid":10046,"taskid":taskid,"tbid":tbid},
    {"customfield":10819,"valuetype":"string","stringvalue":"否","compid":10046,"taskid":taskid,"tbid":tbid},
    {"customfield":10756,"valuetype":"string","stringvalue":"","compid":10046,"taskid":taskid,"tbid":tbid},
    {"customfield":10748,"valuetype":"number","numbervalue":None,"compid":10046,"taskid":taskid,"tbid":tbid},
    {"customfield":10747,"valuetype":"string","stringvalue":"","compid":10046,"taskid":taskid,"tbid":tbid},
    {"customfield":10822,"valuetype":"number","numbervalue":100,"compid":10046,"taskid":taskid,"tbid":tbid},
    {"customfield":10820,"valuetype":"string","stringvalue":"否","compid":10046,"taskid":taskid,"tbid":tbid},
    {"customfield":10757,"valuetype":"string","stringvalue":"","compid":10046,"taskid":taskid,"tbid":tbid},
    {"customfield":10751,"valuetype":"number","numbervalue":None,"compid":10046,"taskid":taskid,"tbid":tbid},
    {"customfield":10530,"valuetype":"string","stringvalue":"","compid":10046,"taskid":taskid,"tbid":tbid}
]

# 接口（和你前端完全一致）
url = "https://surveysc.iearthtime.com:5088/surveysc/api/restUploadFieldscreenForTB?version=36&hasError=false&skipResponse=true"

# 请求头
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Content-Type": "application/json;charset=UTF-8",
    "Cookie": COOKIE,
    "Referer": "https://surveysc.iearthtime.com:5088/surveysc/mapClassifyTaskDetail.html",
    "Origin": "https://surveysc.iearthtime.com:5088",
    "X-Requested-With": "XMLHttpRequest"
}

# 提交
print("🚀 正在提交：taskid=1707383 tbid=138541432")
resp = requests.post(url, json=payload, headers=headers, verify=False)

print("状态码:", resp.status_code)
print("返回:", resp.text)

if "sucessful" in resp.text:
    print("✅ ✅ ✅ 提交成功！数据已完全按你要求设置！")
else:
    print("❌ 提交失败")