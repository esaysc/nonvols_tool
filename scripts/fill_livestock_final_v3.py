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

            # 严格匹配数字
            ji = extract_count([r"(\d+)\s*只鸡", r"鸡\s*(\d+)\s*只"], info)
            ya = extract_count([r"(\d+)\s*只鸭", r"鸭\s*(\d+)\s*只", r"甲鸭\s*(\d+)\s*只"], info)
            e = extract_count([r"(\d+)\s*只鹅", r"鹅\s*(\d+)\s*只"], info)

            final_data.append({"负责人": name, "鸡": ji, "鸭": ya, "鹅": e})

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
        
        # 先清空可能存在的旧数据区域（D8:Q50 左右，保守起见）
        # sheet.range('D8:Q50').clear_contents() 
        
        for i, entry in enumerate(data):
            row = start_row + i
            # 负责人 D列
            sheet.range(f'D{row}').value = entry["负责人"]
            # 鸡存栏 M列 (鸡存栏(只))
            sheet.range(f'M{row}').value = entry["鸡"] if entry["鸡"] > 0 else ""
            # 鸭存栏 O列
            sheet.range(f'O{row}').value = entry["鸭"] if entry["鸭"] > 0 else ""
            # 鹅存栏 Q列
            sheet.range(f'Q{row}').value = entry["鹅"] if entry["鹅"] > 0 else ""
        
        if os.path.exists(abs_output):
            os.remove(abs_output)
        wb.save(abs_output)
        wb.close()
        print(f"成功生成: {output_file}")
    finally:
        app.quit()

feedback_text = """
曹世明：10 只鸡 10 只鸭
刘柱明：21 只鸡
刘俊华：14 只鸡 4 只鸭
程金本：14 只鸡
周春华：15 只鸡 鸭 15 只
程文书：10 只鸡 鸭 10 只 鹅 2 只
程仲良：3 只鸡
曹永根：13 只鸡 鸭 10 只 鹅 3 只
刘华良：16 只鸡 鸭 3 只 鹅 3 只
夏儒明：鹅 3 只
刘黄根：13 只鸡 鸭 7 只
刘小全：6 只鸡 鸭 3 只
刘双全：46 只鸡
刘金明：40 只鸡 鸭 10 只 鹅 6 只
刘现文：5 只鸡
郑勇军：2 只鸡
刘水文：20 只鸡
刘柱根：30 只鸡 鸭 25 只 鹅 8 只
王学文：5 只鸡
黄建文：20 只鸡 鸭 9 只 鹅 2 只
刘文成：60 只鸡 鸭 20 只 鹅 7 只
王元根：26 只鸡 鸭 10 只 鹅 2 只
李光全：12 只鸡 鸭 5 只
刘吉元：鸡 17 只，鸭 1 只，鹅 4 只
程澤群：鸡 4 只、鸭 10 只、鹅 2 只
李树军：4 只鸡
"""

if __name__ == "__main__":
    parsed_data = parse_feedback_data(feedback_text)
    # 打印解析结果进行自我核对
    for item in parsed_data:
        print(f"{item['负责人']}: 鸡{item['鸡']} 鸭{item['鸭']} 鹅{item['鹅']}")
    fill_excel_xlwings(parsed_data, "畜禽生产台账空表.xls", "畜禽生产台账_最终确认版.xls")
