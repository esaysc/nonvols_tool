import xlrd
from xlutils.copy import copy


def adjust_xls_keep_style(input_path, output_path):
    # 打开原始文件，formatting_info=True 保留格式信息
    rb = xlrd.open_workbook(input_path, formatting_info=True)
    sheet_read = rb.sheet_by_index(0)

    # 使用 xlutils.copy 复制工作簿，这会保留大部分样式（字体、边框、对齐等）
    wb = copy(rb)
    new_sheet = wb.get_sheet(0)

    # 调整行列均匀
    # xlwt 的宽度单位是 1/256 个字符宽度，高度单位是 twips (1/20 point)
    default_width = 256 * 10  # 约10个字符宽
    default_height = 20 * 25  # 约25磅高

    # 设置列宽
    for i in range(sheet_read.ncols):
        new_sheet.col(i).width = default_width

    # 设置行高
    for i in range(sheet_read.nrows):
        new_sheet.row(i).height_mismatch = True
        new_sheet.row(i).height = default_height

    wb.save(output_path)
    print(f"已保存保留样式的调整后文件到: {output_path}")


if __name__ == "__main__":
    input_file = "村居委换届选举场地示意图(贵平镇).xls"
    output_file = "村居委换届选举场地示意图(贵平镇)_调整版_保留样式.xls"

    try:
        adjust_xls_keep_style(input_file, output_file)
    except Exception as e:
        print(f"处理失败: {e}")
