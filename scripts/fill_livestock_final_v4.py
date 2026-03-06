import re
import os
import xlwings as xw


def parse_feedback_data(text):
    lines = text.strip().split("\n")
    final_data = []

    for line in lines:
        if "：" in line:
            name, info = line.split("：", 1)
            name = name.strip()

            def extract_count(patterns, text):
                total = 0
                for p in patterns:
                    matches = re.findall(p, text)
                    for m in matches:
                        if isinstance(m, tuple):
                            for group in m:
                                if group:
                                    total += int(group)
                        elif m:
                            total += int(m)
                return total

            # 针对最新反馈格式进行匹配：鸡 10 只、鸭 10 只
            ji = extract_count([r"鸡\s*(\d+)\s*只", r"(\d+)\s*只鸡"], info)
            ya = extract_count(
                [r"鸭\s*(\d+)\s*只", r"(\d+)\s*只鸭", r"甲鸭\s*(\d+)\s*只"], info
            )
            e = extract_count([r"鹅\s*(\d+)\s*只", r"(\d+)\s*只鹅"], info)
            zhu = extract_count([r"猪\s*(\d+)\s*[只头]", r"(\d+)\s*[只头]猪"], info)
            tu = extract_count([r"兔\s*(\d+)\s*只", r"(\d+)\s*只兔"], info)

            final_data.append(
                {"负责人": name, "鸡": ji, "鸭": ya, "鹅": e, "猪": zhu, "兔": tu}
            )

    return final_data


def fill_excel_xlwings(data, input_file, output_file):
    abs_input = os.path.abspath(input_file)
    abs_output = os.path.abspath(output_file)

    app = xw.App(visible=False)
    try:
        wb = app.books.open(abs_input)
        sheet = wb.sheets[0]

        # 数据起始行
        start_row = 8

        for i, entry in enumerate(data):
            row = start_row + i
            # 负责人 D列
            sheet.range(f"D{row}").value = entry["负责人"]
            # 猪存栏 I列 (假设)
            sheet.range(f"I{row}").value = entry["猪"] if entry["猪"] > 0 else ""
            # 鸡存栏 M列
            sheet.range(f"M{row}").value = entry["鸡"] if entry["鸡"] > 0 else ""
            # 鸭存栏 O列
            sheet.range(f"O{row}").value = entry["鸭"] if entry["鸭"] > 0 else ""
            # 鹅存栏 Q列
            sheet.range(f"Q{row}").value = entry["鹅"] if entry["鹅"] > 0 else ""
            # 兔存栏 S列 (假设)
            sheet.range(f"S{row}").value = entry["兔"] if entry["兔"] > 0 else ""

        if os.path.exists(abs_output):
            os.remove(abs_output)
        wb.save(abs_output)
        wb.close()
        print(f"成功生成: {output_file}")
    finally:
        app.quit()


feedback_text = """
王九彬：鸡 6 只
邓花云：鸡 10 只
胡泽均：鸡 6 只，鸭 4 只
徐定兵：鸡 100 只
周勇：鸡 10 只，猪 1 头
孙学明：鸡 10 只，鸭 5 只
王如贵：鸡 10 只
王长江：鸡 8 只，鸭 6 只
王如明：鸡 10 只
曾红英：鸡 6 只
汤淑清：鸡 8 只
冯子俊：鸡 8 只
费兴明：鸡 10 只
汪子均：鸡 8 只
汪子明：鸡 8 只
雷平华：鸡 6 只
刘子文：鸡 6 只
夏尚方：鸡 8 只
曹有战：鸡 6 只
王子华：鸡 6 只
赵忠明：鸡 6 只
王明钦：鸡 10 只
徐世军：鸡 8 只
王文华：鸡 6 只
林贵志：鸡 4 只
王元中：鸡 4 只
王乾树：鸡 6 只
魏芳：鸡 6 只
王立均：鸡 6 只
王乾云：鸡 8 只
王小文：鸡 12 只
"""

if __name__ == "__main__":
    parsed_data = parse_feedback_data(feedback_text)
    for item in parsed_data:
        print(
            f"{item['负责人']}: 猪{item['猪']} 鸡{item['鸡']} 鸭{item['鸭']} 鹅{item['鹅']} 兔{item['兔']}"
        )
    fill_excel_xlwings(
        parsed_data, "畜禽生产台账空表.xls", "畜禽生产台账_最终版_格式严格.xls"
    )
