import os
import pandas as pd
from bs4 import BeautifulSoup
import re

def parse_html_to_excel(html_path, excel_path):
    print(f"正在处理: {html_path}")
    
    with open(html_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    soup = BeautifulSoup(content, 'lxml')
    
    # 提取基本信息
    account_set = ""
    time_range = ""
    
    text_boxes = soup.find_all('div', class_='text-box')
    for box in text_boxes:
        text = box.get_text(strip=True)
        if "当前账套：" in text:
            account_set = text.replace("当前账套：", "")
        elif "时间范围：" in text:
            time_range = text.replace("时间范围：", "")
            
    print(f"账套: {account_set}")
    print(f"时间范围: {time_range}")

    # 提取表格数据
    # 资金流向明细(入) 和 资金流向明细(出)
    all_data = []
    
    table_boxes = soup.find_all('div', class_='table-box')
    for box in table_boxes:
        title_div = box.find('div', class_='speacl-text')
        if not title_div:
            continue
            
        table_type = title_div.get_text(strip=True)
        print(f"发现表格: {table_type}")
        
        # 找到对应的数据表
        # 数据在 el-table__body 类的 table 中
        body_table = box.find('table', class_='el-table__body')
        if not body_table:
            continue
            
        rows = body_table.find_all('tr', class_='el-table__row')
        for row in rows:
            cols = row.find_all('td')
            if len(cols) >= 4:
                idx = cols[0].get_text(strip=True)
                date = cols[1].get_text(strip=True)
                summary = cols[2].get_text(strip=True)
                amount = cols[3].get_text(strip=True).replace(',', '')
                
                all_data.append({
                    '账套': account_set,
                    '时间范围': time_range,
                    '类型': table_type,
                    '序号': idx,
                    '日期': date,
                    '摘要': summary,
                    '金额': float(amount) if amount else 0.0
                })

    if all_data:
        df = pd.DataFrame(all_data)
        df.to_excel(excel_path, index=False)
        print(f"成功保存到: {excel_path}")
    else:
        print("未发现有效数据")

if __name__ == "__main__":
    # 自动处理 data/html 目录下的所有 html 文件
    html_dir = os.path.join(os.path.dirname(__file__), '..', 'data', 'html')
    output_dir = os.path.join(os.path.dirname(__file__), '..', 'data', 'excel')
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    for file in os.listdir(html_dir):
        if file.endswith('.html'):
            html_path = os.path.join(html_dir, file)
            excel_name = file.replace('.html', '.xlsx')
            excel_path = os.path.join(output_dir, excel_name)
            parse_html_to_excel(html_path, excel_path)
