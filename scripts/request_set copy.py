import requests
import json
import time
import pandas as pd
import sys
import warnings
import random  # 新增随机数模块
from requests.exceptions import RequestException, Timeout, SSLError
from urllib3.exceptions import InsecureRequestWarning

# 正确过滤业务无关的警告
warnings.filterwarnings("ignore", category=DeprecationWarning)
warnings.filterwarnings("ignore", category=UserWarning)
warnings.filterwarnings(
    "ignore", category=InsecureRequestWarning
)  # 适配verify=False的SSL警告

# ===================== 基础配置（业务固定值，请勿随意修改） =====================
# 业务固定compid，与前端提交完全一致
COMP_ID = 10046
# 接口地址，与前端完全一致
API_URL = "https://surveysc.iearthtime.com:5088/surveysc/api/restUploadFieldscreenForTB?version=36&hasError=false&skipResponse=true"
# 字段类型映射，与前端接口要求完全对齐
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
# 1:1 复刻你提供的前端请求头，与真人浏览器提交完全一致
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


# ===================== 1. Cookie加载与有效性校验 =====================
def load_and_validate_cookie():
    """加载Cookie，清洗特殊字符，校验格式有效性"""
    try:
        with open("config.txt", "r", encoding="utf-8") as f:
            # 清洗Cookie：去除所有换行、回车、首尾空格，确保单行有效
            raw_cookie = f.read()
            cleaned_cookie = raw_cookie.replace("\n", "").replace("\r", "").strip()
        if not cleaned_cookie:
            print("❌ config.txt 中Cookie为空，请填写有效Cookie")
            sys.exit(1)
        # 校验Cookie格式，必须包含JSESSIONID
        if "JSESSIONID" not in cleaned_cookie:
            print("⚠️  警告：Cookie中未找到JSESSIONID，可能无法正常提交")
        print("✅ Cookie加载并清洗完成")
        return cleaned_cookie
    except FileNotFoundError:
        print("❌ 未找到 config.txt 文件，请在当前目录创建该文件并填写Cookie")
        sys.exit(1)
    except Exception as e:
        print(f"❌ 加载Cookie失败：{str(e)}")
        sys.exit(1)


# ===================== 2. 读取STD_DATA.csv 填报标准数据 =====================
def load_std_data():
    """读取填报标准数据，做格式清洗、合法性校验"""
    # 兼容多种编码，优先utf-8，其次gbk、gb2312
    encodings = ["utf-8", "gbk", "gb2312", "utf-8-sig"]
    df = None
    for enc in encodings:
        try:
            df = pd.read_csv("STD_DATA.csv", dtype=str, encoding=enc)
            print(f"✅ STD_DATA.csv 读取成功，编码：{enc}")
            break
        except Exception as e:
            continue
    if df is None:
        print("❌ STD_DATA.csv 读取失败，请检查文件路径、编码、格式是否正确")
        sys.exit(1)

    # 数据清洗与校验
    std_data = {}
    invalid_fields = []
    for _, row in df.iterrows():
        # 清洗字段号：去除空格、换行
        field_code = str(row.iloc[0]).strip()
        # 清洗值：去除空格、换行，空值/纯空格转为None
        raw_value = row.iloc[1]
        cleaned_value = (
            raw_value.strip() if pd.notna(raw_value) and raw_value.strip() else None
        )
        # 自动转换数字类型
        if cleaned_value and cleaned_value.isdigit():
            cleaned_value = int(cleaned_value)
        # 校验字段是否在接口允许的范围内
        if field_code not in FIELD_TYPE_MAP:
            invalid_fields.append(field_code)
            continue
        std_data[field_code] = cleaned_value

    # 校验结果
    if not std_data:
        print("❌ 未读取到有效填报字段，请检查STD_DATA.csv格式是否正确")
        sys.exit(1)
    if invalid_fields:
        print(f"⚠️  警告：以下字段不在接口允许范围内，已自动跳过：{invalid_fields}")
    print(f"✅ 有效填报字段加载完成，共 {len(std_data)} 个")
    return std_data


# ===================== 3. 读取TB_LIST.csv 地块列表（适配无表头CSV） =====================
def load_tb_list():
    """读取地块列表，适配无表头CSV，做数字校验、格式清洗，过滤无效数据"""
    # 兼容多种编码
    encodings = ["utf-8", "gbk", "gb2312", "utf-8-sig"]
    df = None
    for enc in encodings:
        try:
            # 核心适配：header=None 告诉pandas没有表头，所有行都是数据
            df = pd.read_csv("TB_LIST.csv", encoding=enc, header=None)
            print(f"✅ TB_LIST.csv 读取成功，编码：{enc}")
            break
        except Exception as e:
            continue
    if df is None:
        print("❌ TB_LIST.csv 读取失败，请检查文件路径、编码、格式是否正确")
        sys.exit(1)

    # 数据清洗与校验
    tb_list = []
    invalid_rows = []
    for idx, row in df.iterrows():
        try:
            # 强制转换为整数，校验是否为有效数字
            tbid = int(row.iloc[0])
            taskid = int(row.iloc[1])
            tb_list.append((tbid, taskid))
        except (ValueError, TypeError):
            invalid_rows.append(idx + 1)  # 行号从1开始，适配无表头CSV
            continue

    # 校验结果
    if not tb_list:
        print("❌ 未读取到有效地块，请检查TB_LIST.csv中tbid和taskid是否为有效数字")
        sys.exit(1)
    if invalid_rows:
        print(f"⚠️  警告：以下行数据无效，已自动跳过：行号 {invalid_rows}")
    print(f"✅ 有效地块加载完成，共 {len(tb_list)} 个")
    return tb_list


# ===================== 4. 核心安全环节：预览+手动二次确认 =====================
def security_confirm(tb_list, std_data):
    """安全确认环节，预览所有内容，手动确认后才允许提交"""
    print("\n" + "=" * 60)
    print("🔒 【安全确认环节】请仔细核对以下所有信息，无误后再确认提交")
    print("=" * 60)

    # 打印地块清单，超过10个仅打印前10个，避免控制台溢出
    print(f"\n📋 要提交的地块清单（共 {len(tb_list)} 个）：")
    if len(tb_list) <= 10:
        for idx, (tbid, taskid) in enumerate(tb_list, 1):
            print(f"  {idx}. tbid={tbid}, taskid={taskid}")
    else:
        for idx, (tbid, taskid) in enumerate(tb_list[:10], 1):
            print(f"  {idx}. tbid={tbid}, taskid={taskid}")
        print(f"  ... 剩余 {len(tb_list) - 10} 个地块，已全部加载完成")

    # 打印填报内容
    print("\n📝 要填报的标准内容：")
    for field_code, value in std_data.items():
        print(f"  字段{field_code}: {value}")

    # 强制二次确认
    print("\n" + "=" * 60)
    confirm_input = (
        input("⚠️  请确认以上信息无误，输入 y/yes 开始提交，其他输入直接退出：")
        .strip()
        .lower()
    )
    if confirm_input not in ["y", "yes"]:
        print("❌ 已取消提交，程序正常退出")
        sys.exit(0)
    print("\n✅ 已确认，开始批量提交...\n")


# ===================== 5. 构造请求体，与前端格式完全一致 =====================
def build_request_body(tbid, taskid, std_data):
    """构造符合接口要求的请求体，与前端提交格式100%一致"""
    body = []
    base_params = {"compid": COMP_ID, "taskid": taskid, "tbid": tbid}
    for field_code, value_type in FIELD_TYPE_MAP.items():
        # 仅处理CSV中配置的字段
        if field_code not in std_data:
            continue
        item = {"customfield": int(field_code), "valuetype": value_type, **base_params}
        # 按值类型赋值
        if value_type == "string":
            item["stringvalue"] = std_data[field_code]
        elif value_type == "number":
            item["numbervalue"] = std_data[field_code]
        body.append(item)
    return body


# ===================== 6. 批量提交，带重试机制、异常处理、日志留存 =====================
def batch_submit(tb_list, std_data, cookie):
    """批量提交，带重试机制、异常处理、结果统计"""
    # 初始化Session，保持长连接，与请求头keep-alive匹配
    session = requests.Session()
    # 注入Cookie到请求头
    request_headers = REQUEST_HEADERS.copy()
    request_headers["Cookie"] = cookie

    # 结果统计
    success_list = []
    fail_list = []
    total_count = len(tb_list)

    for idx, (tbid, taskid) in enumerate(tb_list, 1):
        print(f"[{idx}/{total_count}] 正在提交 tbid={tbid}, taskid={taskid}")
        # 构造请求体
        request_body = build_request_body(tbid, taskid, std_data)
        if not request_body:
            print(f"❌ 无有效填报数据，跳过该地块\n")
            fail_list.append((tbid, taskid, "无有效填报数据"))
            continue

        # 提交逻辑，带1次重试，应对网络波动
        retry_count = 0
        max_retry = 1
        submit_success = False
        error_msg = ""
        while retry_count <= max_retry and not submit_success:
            try:
                response = session.post(
                    API_URL,
                    json=request_body,
                    headers=request_headers,
                    verify=False,
                    timeout=25,  # 延长超时时间，应对网络波动
                )
                # 解析响应结果
                if response.status_code == 200:
                    try:
                        response_json = response.json()
                        if response_json.get("sucessful"):
                            submit_success = True
                        else:
                            error_msg = f"接口返回失败：{response_json.get('message', '无错误信息')}"
                    except:
                        # 非JSON响应，按成功处理（符合接口skipResponse=true的设计）
                        submit_success = True
                else:
                    error_msg = f"HTTP状态码错误：{response.status_code}，响应内容：{response.text[:200]}"
            except Timeout:
                error_msg = "请求超时，网络波动"
                retry_count += 1
                if retry_count <= max_retry:
                    print(f"⚠️  超时，第{retry_count}次重试...")
                    time.sleep(2)
            except RequestException as e:
                error_msg = f"请求异常：{str(e)}"
                break
            except Exception as e:
                error_msg = f"未知异常：{str(e)}"
                break

        # 结果处理
        if submit_success:
            print(f"✅ 提交成功\n")
            success_list.append((tbid, taskid))
        else:
            print(f"❌ 提交失败：{error_msg}\n")
            fail_list.append((tbid, taskid, error_msg))

        # 核心修改：5-20秒随机延迟，完全模拟真人操作节奏
        wait_seconds = random.randint(5, 20)
        print(f"⏳ 随机等待 {wait_seconds} 秒...\n")
        time.sleep(wait_seconds)

    # 最终结果统计
    print("=" * 60)
    print("📊 批量提交最终结果")
    print("=" * 60)
    print(f"总地块数：{total_count}")
    print(f"✅ 成功：{len(success_list)} 个")
    print(f"❌ 失败：{len(fail_list)} 个")
    if fail_list:
        print("\n失败地块清单：")
        for tbid, taskid, msg in fail_list:
            print(f"  tbid={tbid}, taskid={taskid}：{msg}")
    print("=" * 60)

    # 留存提交日志到文件
    try:
        with open("submit_log.txt", "a", encoding="utf-8") as f:
            f.write(f"\n===== 提交日志 {time.strftime('%Y-%m-%d %H:%M:%S')} =====\n")
            f.write(
                f"总地块数：{total_count}，成功：{len(success_list)}，失败：{len(fail_list)}\n"
            )
            if success_list:
                f.write("成功地块：\n")
                for tbid, taskid in success_list:
                    f.write(f"  tbid={tbid}, taskid={taskid}\n")
            if fail_list:
                f.write("失败地块：\n")
                for tbid, taskid, msg in fail_list:
                    f.write(f"  tbid={tbid}, taskid={taskid}：{msg}\n")
        print(f"✅ 提交日志已保存到 submit_log.txt")
    except Exception as e:
        print(f"⚠️  保存日志失败：{str(e)}")

    return success_list, fail_list


# ===================== 主程序入口 =====================
if __name__ == "__main__":
    print("🔹 地块批量填报程序启动，已加载全量安全校验机制")
    # 1. 加载并校验Cookie
    valid_cookie = load_and_validate_cookie()
    # 2. 加载填报标准数据
    standard_data = load_std_data()
    # 3. 加载地块列表（已适配无表头CSV）
    target_tb_list = load_tb_list()
    # 4. 安全确认环节
    security_confirm(target_tb_list, standard_data)
    # 5. 批量提交
    batch_submit(target_tb_list, standard_data, valid_cookie)
    # 6. 程序结束
    print("\n🔹 程序执行完成，正常退出")
    sys.exit(0)
