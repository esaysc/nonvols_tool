import xlrd
import xlwt


def adjust_xls(input_path, output_path):
    # 打开原始文件
    # formatting_info=True 保留格式信息（如合并单元格）
    rb = xlrd.open_workbook(input_path, formatting_info=True)
    sheet = rb.sheet_by_index(0)

    # 创建一个新的工作簿
    wb = xlwt.Workbook()
    new_sheet = wb.add_sheet(sheet.name)

    # 复制内容和合并单元格
    # xlrd 的 merged_cells 格式为 (rlo, rhi, clo, chi)
    merged_cells = sheet.merged_cells

    # 记录合并单元格的范围，避免重复写入
    merged_map = {}
    for rlo, rhi, clo, chi in merged_cells:
        for r in range(rlo, rhi):
            for c in range(clo, chi):
                merged_map[(r, c)] = (rlo, rhi, clo, chi)

    # 写入数据
    for r in range(sheet.nrows):
        for c in range(sheet.ncols):
            if (r, c) in merged_map:
                rlo, rhi, clo, chi = merged_map[(r, c)]
                # 只在合并区域的左上角单元格调用 write_merge
                if r == rlo and c == clo:
                    new_sheet.write_merge(
                        rlo, rhi - 1, clo, chi - 1, sheet.cell_value(r, c)
                    )
            else:
                new_sheet.write(r, c, sheet.cell_value(r, c))

    # 调整行列均匀
    # 假设“最小单元格”是指未合并的单元格
    # 我们将所有行高和列宽设为一致的值
    # xlwt 的宽度单位是 1/256 个字符宽度，高度单位是 twips (1/20 point)

    default_width = 256 * 10  # 约10个字符宽
    default_height = 20 * 25  # 约25磅高

    for i in range(sheet.ncols):
        new_sheet.col(i).width = default_width

    for i in range(sheet.nrows):
        new_sheet.row(i).height_mismatch = True
        new_sheet.row(i).height = default_height

    wb.save(output_path)
    print(f"已保存调整后的文件到: {output_path}")


if __name__ == "__main__":
    input_file = "村居委换届选举场地示意图(贵平镇).xls"
    output_file = "村居委换届选举场地示意图(贵平镇)_调整版.xls"
    adjust_xls(input_file, output_file)
