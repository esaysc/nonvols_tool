# 导入依赖
import json
import pandas as pd
import os

# ====================== 【只需改这里】 ======================
# 你的 JSON 文件路径（改成你自己的）
json_file_path = r"data.json"
# 输出 CSV 路径
output_csv_path = r"STD_DATA.csv"
# ============================================================

# 从文件读取 JSON
with open(json_file_path, "r", encoding="utf-8") as f:
    json_data = json.load(f)

# 按规则转换为 field + value 格式
result = []
for item in json_data:
    field = item["customfield"]

    # 根据类型取值
    if item["valuetype"] == "string":
        value = item["stringvalue"]
    elif item["valuetype"] == "number":
        value = item["numbervalue"]
    else:
        value = None

    result.append({"field": field, "value": value})

# 生成 DataFrame 并保存 CSV
df = pd.DataFrame(result)
df.to_csv(output_csv_path, index=False, encoding="utf-8-sig")

# 预览
print("✅ 转换成功！CSV 内容预览：")
print("-" * 40)
print(df.to_string(index=False))
print("-" * 40)
print(f"📁 文件已保存到：{os.path.abspath(output_csv_path)}")
