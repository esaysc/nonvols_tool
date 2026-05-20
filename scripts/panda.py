import requests
import pandas as pd
import time

# ---------------------- 1. 配置信息 ----------------------
login_url = "https://mstf.widthsoft.com/width-website-webapi/publicController/login.do"
data_api_url = "https://mstf.widthsoft.com/width-website-webapi/fundService/financialPublic.do"
detail_api_url = "https://mstf.widthsoft.com/width-website-webapi/fundService/getVoucher.do"

headers = {
    "Accept": "*/*",
    "Accept-Encoding": "gzip, deflate, br, zstd",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
    "Connection": "keep-alive",
    "Content-Type": "application/json;charset=UTF-8",
    "Host": "mstf.widthsoft.com",
    "Idcard": "JPDm25+h/0kGoSc8Lv+jwi2xfcb+Hn4KaQhmMUgKdQ4oFLERos07iL5AaE8QiV1H0VoxeGtf28ylfxqBnXe1Cqyve1tyAJwSXqlXnQKobwB/qsQ52k2PtxRlRSK2r7cloysMiJvz7T/U7zPrtSDo+PuabVmRHbxjs1LOJdlbbJqjq4hbfe8fI3HObz8pv0mkN+L13Ruy8LVZAh3piDNicmA0dUgMk4PBVQEgxHhnS3TsSc3hlXU8G3kC4SVsHZoxVRm8T72lVCly2C5PtOggbq9lKwdHfWo0AA/dtBXb5SkjO385GzSwEIC9+rcxp9ARMjZy7pEDXlrh67KO235tiQ==",
    "Nmbtoken": "null",
    "Origin": "https://mstf.widthsoft.com",
    "Referer": "https://mstf.widthsoft.com/Login",
    "Sec-Ch-Ua": '"Microsoft Edge";v="147", "Not.A/Brand";v="8", "Chromium";v="147"',
    "Sec-Ch-Ua-Mobile": "?0",
    "Sec-Ch-Ua-Platform": '"Windows"',
    "Sec-Fetch-Dest": "empty",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Site": "same-origin",
    "Sshtoken": "null",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36 Edg/147.0.0.0",
    "Webdomain": "mstf.widthsoft.com"
}

session = requests.Session()

# ---------------------- 2. 账号登录 ----------------------
def login():
    payload = {
        "identityId": "JPDm25+h/0kGoSc8Lv+jwi2xfcb+Hn4KaQhmMUgKdQ4oFLERos07iL5AaE8QiV1H0VoxeGtf28ylfxqBnXe1Cqyve1tyAJwSXqlXnQKobwB/qsQ52k2PtxRlRSK2r7cloysMiJvz7T/U7zPrtSDo+PuabVmRHbxjs1LOJdlbbJqjq4hbfe8fI3HObz8pv0mkN+L13Ruy8LVZAh3piDNicmA0dUgMk4PBVQEgxHhnS3TsSc3hlXU8G3kC4SVsHZoxVRm8T72lVCly2C5PtOggbq9lKwdHfWo0AA/dtBXb5SkjO385GzSwEIC9+rcxp9ARMjZy7pEDXlrh67KO235tiQ==",
        "name": "王小惠",
        "appType": "公开查询"
    }
    res = session.post(login_url, json=payload, headers=headers)
    try:
        response_json = res.json()
        print("登录响应：", response_json)
        return response_json
    except Exception as e:
        print(f"解析登录响应失败: {e}")
        return None

# ---------------------- 3. 请求业务API拿数据 ----------------------
def get_voucher_list(year):
    payload = {
        "unitId": "51140314918002",
        "startMonth": 1,
        "endMonth": 12,
        "year": year
    }
    res = session.post(data_api_url, headers=headers, json=payload)
    try:
        return res.json()
    except Exception as e:
        print(f"获取 {year} 年业务数据失败: {e}")
        return None

def get_voucher_detail(voucher_id, year):
    """获取凭证详细信息"""
    if not voucher_id:
        return None
    payload = {
        "voucherId": voucher_id,
        "year": year
    }
    res = session.post(detail_api_url, headers=headers, json=payload)
    try:
        return res.json()
    except Exception as e:
        print(f"获取凭证详情失败({voucher_id}, {year}): {e}")
        return None

def process_list_with_details(item_list, list_name, year):
    """通用处理函数：抓取详情并生成概览和详情数据"""
    detailed_rows = []
    simple_rows = []
    
    for i, item in enumerate(item_list):
        v_id = item.get("voucherId")
        summary = item.get("summary", "")
        
        if summary == "合计" or not v_id:
            continue
        
        print(f"  [{i+1}/{len(item_list)}] 正在抓取{year}年{list_name}: {summary[:20]}...")
        detail_res = get_voucher_detail(v_id, year)
        
        voucher_no = ""
        if detail_res and detail_res.get("success") and detail_res.get("data"):
            v_data = detail_res["data"]
            master = v_data.get("master", {})
            details = v_data.get("details", [])
            voucher_no = master.get("voucherNo", "")

            for d in details:
                detailed_rows.append({
                    "年份": year,
                    "凭证号": voucher_no,
                    f"{list_name}摘要(总)": summary,
                    "分录摘要": d.get("summary"),
                    "会计科目": d.get("fullsubjectName"),
                    "方向": d.get("orientation"),
                    "金额": d.get("amount"),
                    "日期": master.get("makingDate"),
                    "制单人": master.get("makingUser"),
                    "审核人": master.get("verifyUser"),
                    "出纳": master.get("cashier"),
                    "凭证ID": v_id
                })
        
        simple_rows.append({
            "年份": year,
            "凭证号": voucher_no,
            "摘要": summary,
            "金额": item.get("amount"),
            "日期": item.get("makingDate"),
            "凭证ID": v_id,
            "备注": "" if voucher_no else "详情抓取失败"
        })
        time.sleep(0.3)
    
    return simple_rows, detailed_rows

# ---------------------- 4. 主逻辑 ----------------------
def main():
    login_res = login()
    if not (login_res and login_res.get("success")):
        print("登录失败，停止后续操作")
        return

    all_simple_inc = []
    all_detailed_inc = []
    all_simple_pay = []
    all_detailed_pay = []

    years = range(2019, 2027)
    for year in years:
        print(f"\n>>> 正在处理 {year} 年数据...")
        json_data = get_voucher_list(year)
        if not json_data or "data" not in json_data or not json_data["data"]:
            print(f"  {year} 年没有数据或获取失败")
            continue
        
        data = json_data["data"]
        
        # 处理收入
        income_list = data.get("incomeList", [])
        if income_list:
            s_inc, d_inc = process_list_with_details(income_list, "收入", year)
            all_simple_inc.extend(s_inc)
            all_detailed_inc.extend(d_inc)

        # 处理支出
        pay_list = data.get("payList", [])
        if pay_list:
            s_pay, d_pay = process_list_with_details(pay_list, "支出", year)
            all_simple_pay.extend(s_pay)
            all_detailed_pay.extend(d_pay)

    # 导出 Excel
    output_file = "财务公开数据_2019-2026.xlsx"
    try:
        with pd.ExcelWriter(output_file) as writer:
            if all_simple_inc:
                pd.DataFrame(all_simple_inc).to_excel(writer, sheet_name="收入明细(不含详情)", index=False)
            if all_detailed_inc:
                pd.DataFrame(all_detailed_inc).to_excel(writer, sheet_name="收入明细(含详情)", index=False)
            if all_simple_pay:
                pd.DataFrame(all_simple_pay).to_excel(writer, sheet_name="支出明细(不含详情)", index=False)
            if all_detailed_pay:
                pd.DataFrame(all_detailed_pay).to_excel(writer, sheet_name="支出明细(含详情)", index=False)
        print(f"\n所有年份数据导出完成：{output_file}")
    except PermissionError:
        print(f"错误：无法写入文件 {output_file}。请确保该文件已关闭！")

if __name__ == "__main__":
    main()
