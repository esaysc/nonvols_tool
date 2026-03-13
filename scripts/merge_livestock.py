import pandas as pd
import os
import glob
import re

def merge_excels(input_dir, output_file):
    files = glob.glob(os.path.join(input_dir, "*.xls"))
    def get_num(filename):
        match = re.search(r"_(\d+)\.xls", filename)
        return int(match.group(1)) if match else 0
    files.sort(key=get_num)
    if not files:
        print("No files found to merge.")
        return
    all_data = []
    header_info = None
    for i, file in enumerate(files):
        print(f"Processing {file}...")
        df = pd.read_excel(file, header=None)
        if i == 0:
            header_info = df.iloc[:7]
            data = df.iloc[7:]
        else:
            data = df.iloc[7:]
        all_data.append(data)
    combined_data = pd.concat(all_data, ignore_index=True)
    final_df = pd.concat([header_info, combined_data], ignore_index=True)
    final_df.to_excel(output_file, index=False, header=False)
    print(f"Successfully merged {len(files)} files into {output_file}")

if __name__ == "__main__":
    input_folder = "畜禽生产台账/"
    output_excel = "畜禽生产台账_合并版.xlsx"
    merge_excels(input_folder, output_excel)
