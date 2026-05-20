import pandas as pd

# 读取主表和映射表
df_main = pd.read_excel("input.xlsx")
df_map = pd.read_csv("taskid_map.csv")

# 生成映射字典
taskid_dict = dict(zip(df_map["序号"], df_map["taskid"]))

# 回填 taskid 到 E 列
df_main["taskid"] = df_main["A列的序号列名"].map(taskid_dict)

# 保存结果
df_main.to_excel("output.xlsx", index=False)
print("✅ taskid 已全部回填完成！")