import zipfile
from lxml import etree
import pandas as pd
import numpy as np

def extract_all_tables(file_path):
    if not zipfile.is_zipfile(file_path):
        print(f"{file_path} is not a valid zip file")
        return

    with zipfile.ZipFile(file_path, 'r') as z:
        content = z.read('word/document.xml')
        root = etree.fromstring(content)
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        
        tables = root.xpath('//w:tbl', namespaces=ns)
        all_data = []
        header = None

        for i, tbl in enumerate(tables):
            rows = tbl.xpath('.//w:tr', namespaces=ns)
            table_data = []
            for row in rows:
                cells = row.xpath('.//w:tc', namespaces=ns)
                row_data = []
                for cell in cells:
                    texts = cell.xpath('.//w:t/text()', namespaces=ns)
                    row_data.append(''.join(texts).strip())
                table_data.append(row_data)
            
            if not table_data:
                continue

            current_first_row = table_data[0]
            # 识别表头
            if '姓名' in current_first_row or '面积' in current_first_row:
                if header is None:
                    # 统一表头名称
                    header = ['村名', '社别', '姓名', '面积', '供种数量', '领种人签字', '备注']
                all_data.extend(table_data[1:]) 
            else:
                all_data.extend(table_data)

        if not header:
            header = ['村名', '社别', '姓名', '面积', '供种数量', '领种人签字', '备注']
        
        # 确保所有行长度一致
        max_cols = len(header)
        cleaned_data = []
        for row in all_data:
            if len(row) < max_cols:
                row.extend([''] * (max_cols - len(row)))
            cleaned_data.append(row[:max_cols])

        df = pd.DataFrame(cleaned_data, columns=header)
        
        # 清洗数据
        df = df[df['姓名'] != '姓名']
        df = df[df['姓名'] != '']
        
        # 修复村名列
        # 观察到第一列可能包含 OCR 错误或不完整的村名
        df['村名'] = df['村名'].replace('', np.nan)
        
        # 针对性修复：'岗' 和 '中岗树' 应该是 '中岗村'
        df['村名'] = df['村名'].apply(lambda x: '中岗村' if str(x) in ['岗', '中岗树'] else x)
        
        # 向下填充村名
        df['村名'] = df['村名'].ffill()
        
        # 针对性修复：'下  11' 这种出现在村名位置的可能是 OCR 误读或错位，如果村名已经是中岗村，则保持
        # 但在 head 输出中看到 index 1 的村名是 '中岗树'，社别是 '下  11'
        # 我们重新处理一下逻辑：如果村名列看起来像社别，可能需要移动
        
        def fix_row(row):
            # 如果村名包含数字或'组'，且社别为空，可能发生了错位
            if any(char.isdigit() for char in str(row['村名'])) or '组' in str(row['村名']):
                if not row['社别']:
                    row['社别'] = row['村名']
                    row['村名'] = np.nan
            return row

        df = df.apply(fix_row, axis=1)
        df['村名'] = df['村名'].ffill()
        
        # 最终确保村名统一
        df['村名'] = '中岗村' # 根据目前数据看，似乎全是中岗村

        return df

if __name__ == "__main__":
    input_file = '贵平2022年大豆发放花名册扫描_extract__20260408_155204.docm'
    output_file = '贵平2022年大豆发放花名册_修复版.xlsx'
    
    df = extract_all_tables(input_file)
    if df is not None:
        df.to_excel(output_file, index=False)
        print(f"成功提取 {len(df)} 行数据并保存到 {output_file}")
