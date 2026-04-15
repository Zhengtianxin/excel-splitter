#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 拆分工具 v1.1 (Linux + Python 3.7)
按指定列拆分xlsx文件，保留原表所有格式。
用法：python3 dooo.txt
"""

import os
import sys
import glob
import copy
import traceback
from collections import defaultdict

try:
    from openpyxl import load_workbook, Workbook
    from openpyxl.utils import get_column_letter
except ImportError:
    print("=" * 60)
    print("  错误：缺少 openpyxl 库！")
    print("  请执行以下命令安装：")
    print("    pip3 install openpyxl")
    print("=" * 60)
    input("\n按回车键退出...")
    sys.exit(1)


# ==================== 样式复制工具函数 ====================

def copy_cell_style(src_cell, dst_cell):
    """完整复制单元格样式（字体/填充/边框/对齐/保护/数字格式）"""
    try:
        if src_cell.font:
            dst_cell.font = copy.copy(src_cell.font)
        if src_cell.fill:
            dst_cell.fill = copy.copy(src_cell.fill)
        if src_cell.border:
            dst_cell.border = copy.copy(src_cell.border)
        if src_cell.alignment:
            dst_cell.alignment = copy.copy(src_cell.alignment)
        if src_cell.protection:
            dst_cell.protection = copy.copy(src_cell.protection)
        if src_cell.number_format:
            dst_cell.number_format = src_cell.number_format
    except Exception:
        pass


def copy_cell(src_cell, dst_cell):
    """完整复制单元格（值 + 样式 + 超链接 + 批注）"""
    dst_cell.value = src_cell.value
    copy_cell_style(src_cell, dst_cell)
    try:
        if src_cell.hyperlink:
            dst_cell.hyperlink = copy.copy(src_cell.hyperlink)
    except Exception:
        pass
    try:
        if src_cell.comment:
            dst_cell.comment = copy.copy(src_cell.comment)
    except Exception:
        pass


def copy_sheet_properties(src_ws, dst_ws):
    """复制工作表级别属性（列宽/冻结窗格/页面设置等）"""
    for col_letter, col_dim in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[col_letter].width = col_dim.width
        dst_ws.column_dimensions[col_letter].hidden = col_dim.hidden
        try:
            if col_dim.font:
                dst_ws.column_dimensions[col_letter].font = copy.copy(col_dim.font)
            if col_dim.fill:
                dst_ws.column_dimensions[col_letter].fill = copy.copy(col_dim.fill)
            if col_dim.border:
                dst_ws.column_dimensions[col_letter].border = copy.copy(col_dim.border)
            if col_dim.alignment:
                dst_ws.column_dimensions[col_letter].alignment = copy.copy(col_dim.alignment)
        except Exception:
            pass
    try:
        dst_ws.freeze_panes = src_ws.freeze_panes
    except Exception:
        pass
    try:
        dst_ws.page_setup.orientation = src_ws.page_setup.orientation
        dst_ws.page_setup.paperSize = src_ws.page_setup.paperSize
        dst_ws.page_setup.fitToHeight = src_ws.page_setup.fitToHeight
        dst_ws.page_setup.fitToWidth = src_ws.page_setup.fitToWidth
    except Exception:
        pass
    try:
        if src_ws.print_title_rows:
            dst_ws.print_title_rows = src_ws.print_title_rows
        if src_ws.print_title_cols:
            dst_ws.print_title_cols = src_ws.print_title_cols
    except Exception:
        pass
    try:
        if src_ws.auto_filter and src_ws.auto_filter.ref:
            dst_ws.auto_filter.ref = src_ws.auto_filter.ref
    except Exception:
        pass
    try:
        if src_ws.sheet_properties.tabColor:
            dst_ws.sheet_properties.tabColor = copy.copy(src_ws.sheet_properties.tabColor)
    except Exception:
        pass


# ==================== 显示工具函数 ====================

def truncate(text, max_len=12):
    if text is None:
        return "(空)"
    text = str(text).replace('\n', ' ').replace('\r', '')
    if len(text) > max_len:
        return text[:max_len] + "..."
    return text


def format_row_preview(ws, row_idx, max_col):
    parts = []
    display_cols = min(max_col, 10)
    for col_idx in range(1, display_cols + 1):
        val = ws.cell(row=row_idx, column=col_idx).value
        parts.append(truncate(val))
    preview = " | ".join(parts)
    if max_col > 10:
        preview += " | ...(共%d列)" % max_col
    return preview


def safe_filename(name):
    if not name:
        return "未分类"
    # 同时过滤 Windows 和 Linux 的非法字符
    for ch in '\\/:*?"<>|\0':
        name = name.replace(ch, '_')
    name = name.strip().rstrip('.')
    if not name:
        return "未分类"
    if len(name) > 200:
        name = name[:200]
    return name


# ==================== 嵌套检测 ====================

def detect_nested_values(group_names):
    names = sorted(group_names, key=len)
    nested_pairs = []
    for i in range(len(names)):
        for j in range(i + 1, len(names)):
            short_name, long_name = names[i], names[j]
            if short_name == long_name:
                continue
            if short_name in long_name:
                nested_pairs.append((short_name, long_name))
    return nested_pairs


# ==================== 主逻辑 ====================

def find_xlsx_file(work_dir):
    xlsx_files = glob.glob(os.path.join(work_dir, '*.xlsx'))
    if not xlsx_files:
        return None, "当前目录下没有找到任何 .xlsx 文件！"
    if len(xlsx_files) == 1:
        return xlsx_files[0], None

    print("\n当前目录下有多个 .xlsx 文件，请选择要拆分的文件：")
    print("-" * 50)
    for i, f in enumerate(xlsx_files):
        size = os.path.getsize(f)
        if size > 1024 * 1024:
            size_str = "%.1f MB" % (size / 1024 / 1024)
        else:
            size_str = "%.1f KB" % (size / 1024)
        print("  %d. %s  (%s)" % (i + 1, os.path.basename(f), size_str))
    print("-" * 50)

    while True:
        choice = input("请输入文件编号 (1-%d): " % len(xlsx_files)).strip()
        try:
            idx = int(choice) - 1
            if 0 <= idx < len(xlsx_files):
                return xlsx_files[idx], None
            else:
                print("  x 请输入 1 到 %d 之间的数字" % len(xlsx_files))
        except ValueError:
            print("  x 请输入数字，不要输入其他字符")


def ask_header_row(ws, max_col):
    max_row = ws.max_row
    preview_rows = min(max_row, 5)

    print("\n" + "=" * 60)
    print("  第一步：请选择【标题行】是哪一行")
    print("  (数据行将从标题行的下一行自动开始)")
    print("=" * 60)
    print()

    for r in range(1, preview_rows + 1):
        preview = format_row_preview(ws, r, max_col)
        print("  %d. 第%d行：%s" % (r, r, preview))

    if max_row > preview_rows:
        print("\n  (共 %d 行，以上仅展示前 %d 行)" % (max_row, preview_rows))

    print("\n  0. 其他 -> 手动输入行号")
    print()

    while True:
        choice = input("请选择标题行编号: ").strip()
        try:
            num = int(choice)
            if num == 0:
                custom = input("请输入标题行的行号 (1-%d): " % max_row).strip()
                try:
                    custom_num = int(custom)
                    if 1 <= custom_num <= max_row:
                        return custom_num
                    else:
                        print("  x 行号超出范围，表格共 %d 行，请输入 1 到 %d" % (max_row, max_row))
                except ValueError:
                    print("  x 请输入数字")
            elif 1 <= num <= preview_rows:
                return num
            else:
                print("  x 请输入 0 到 %d 之间的数字" % preview_rows)
        except ValueError:
            print("  x 请输入数字，不要输入其他字符")


def ask_split_column(ws, header_row, max_col):
    print("\n" + "=" * 60)
    print("  第二步：请选择按哪一列拆分")
    print("  (以下是第 %d 行标题行的各列内容)" % header_row)
    print("=" * 60)
    print()

    for col_idx in range(1, max_col + 1):
        col_letter = get_column_letter(col_idx)
        val = ws.cell(row=header_row, column=col_idx).value
        display = str(val) if val is not None else "(空)"
        if len(display) > 30:
            display = display[:30] + "..."
        print("  %d. %s列：%s" % (col_idx, col_letter, display))

    print()

    while True:
        choice = input("请选择要拆分的列号 (1-%d): " % max_col).strip()
        try:
            num = int(choice)
            if 1 <= num <= max_col:
                col_letter = get_column_letter(num)
                val = ws.cell(row=header_row, column=num).value
                display = str(val) if val is not None else "(空)"
                print("\n  OK 已选择：%s列 -[%s]" % (col_letter, display))
                return num
            else:
                print("  x 请输入 1 到 %d 之间的数字，共 %d 列" % (max_col, max_col))
        except ValueError:
            print("  x 请输入数字，不要输入其他字符")


def do_split(src_file, ws, header_row, split_col, max_row, max_col):
    work_dir = os.path.dirname(os.path.abspath(src_file))
    data_start_row = header_row + 1

    if data_start_row > max_row:
        print("\n  x 错误：标题行是第 %d 行，但表格只有 %d 行，没有数据行可拆分！" % (header_row, max_row))
        return False

    col_letter = get_column_letter(split_col)
    col_name = ws.cell(row=header_row, column=split_col).value or "(空)"

    print("\n" + "=" * 60)
    print("  开始拆分")
    print("=" * 60)
    print("  源文件：%s" % os.path.basename(src_file))
    print("  标题行：第 %d 行" % header_row)
    print("  数据行：第 %d 行 ~ 第 %d 行（共 %d 行数据）" % (data_start_row, max_row, max_row - data_start_row + 1))
    print("  拆分列：%s列 -[%s]" % (col_letter, col_name))
    print("-" * 60)

    # 按指定列分组
    groups = defaultdict(list)
    empty_count = 0
    for row_idx in range(data_start_row, max_row + 1):
        cell_value = ws.cell(row=row_idx, column=split_col).value
        if cell_value is None or str(cell_value).strip() == '':
            key = "未分类"
            empty_count += 1
        else:
            key = str(cell_value).strip()
        groups[key].append(row_idx)

    if empty_count > 0:
        print("\n  !! 注意：有 %d 行数据的%s列为空，将归入[未分类.xlsx]" % (empty_count, col_letter))

    print("\n  共将拆分为 %d 个文件：" % len(groups))
    for name, rows in groups.items():
        print("    * %s -> %d 行数据" % (name, len(rows)))

    # 嵌套检测
    real_names = [n for n in groups.keys() if n != "未分类"]
    nested_pairs = detect_nested_values(real_names)
    if nested_pairs:
        print()
        print("  " + "!" * 56)
        print("  !! 警告：以下分组名称之间存在【包含/嵌套】关系：")
        print("  （它们会被拆分成不同的文件，请确认这是您期望的结果）")
        print("  " + "-" * 56)
        for short_name, long_name in nested_pairs:
            s_count = len(groups[short_name])
            l_count = len(groups[long_name])
            print("    [%s](%d行) 被包含于 [%s](%d行)" % (short_name, s_count, long_name, l_count))
        print("  " + "-" * 56)
        print("  如果[%s]和[%s]本应是同一类，" % (nested_pairs[0][0], nested_pairs[0][1]))
        print("  请先在原表中统一%s列的值，再重新运行此工具。" % col_letter)
        print("  " + "!" * 56)
        print()
        confirm = input("  是否继续拆分？(Y/n): ").strip().lower()
        if confirm == 'n':
            print("\n  已取消拆分。")
            return False

    # 收集合并单元格
    header_merged = []
    for mg in ws.merged_cells.ranges:
        if mg.min_row >= 1 and mg.max_row <= header_row:
            header_merged.append(mg)

    data_row_merges = defaultdict(list)
    for mg in ws.merged_cells.ranges:
        if mg.min_row >= data_start_row and mg.min_row == mg.max_row:
            data_row_merges[mg.min_row].append((mg.min_col, mg.max_col))

    # 行高
    row_heights = {}
    for row_idx in range(1, max_row + 1):
        rd = ws.row_dimensions.get(row_idx)
        if rd and rd.height:
            row_heights[row_idx] = rd.height

    # 逐组生成文件
    print("\n  正在生成文件...")
    success_count = 0
    fail_count = 0

    for group_name, src_rows in groups.items():
        fname = safe_filename(group_name)
        dst_file = os.path.join(work_dir, "%s.xlsx" % fname)

        try:
            dst_wb = Workbook()
            dst_ws = dst_wb.active
            dst_ws.title = ws.title

            copy_sheet_properties(ws, dst_ws)

            # 复制标题区域
            for h_row in range(1, header_row + 1):
                for col_idx in range(1, max_col + 1):
                    src_cell = ws.cell(row=h_row, column=col_idx)
                    dst_cell = dst_ws.cell(row=h_row, column=col_idx)
                    copy_cell(src_cell, dst_cell)
                if h_row in row_heights:
                    dst_ws.row_dimensions[h_row].height = row_heights[h_row]

            for mg in header_merged:
                try:
                    dst_ws.merge_cells(
                        start_row=mg.min_row, start_column=mg.min_col,
                        end_row=mg.max_row, end_column=mg.max_col
                    )
                except Exception:
                    pass

            # 复制数据行
            for offset, src_row_idx in enumerate(src_rows):
                dst_row_idx = header_row + 1 + offset

                for col_idx in range(1, max_col + 1):
                    src_cell = ws.cell(row=src_row_idx, column=col_idx)
                    dst_cell = dst_ws.cell(row=dst_row_idx, column=col_idx)
                    copy_cell(src_cell, dst_cell)

                if src_row_idx in row_heights:
                    dst_ws.row_dimensions[dst_row_idx].height = row_heights[src_row_idx]

                if src_row_idx in data_row_merges:
                    for min_c, max_c in data_row_merges[src_row_idx]:
                        try:
                            dst_ws.merge_cells(
                                start_row=dst_row_idx, start_column=min_c,
                                end_row=dst_row_idx, end_column=max_c
                            )
                        except Exception:
                            pass

            # 列宽
            for col_idx in range(1, max_col + 1):
                cl = get_column_letter(col_idx)
                src_dim = ws.column_dimensions.get(cl)
                if src_dim:
                    dst_ws.column_dimensions[cl].width = src_dim.width

            # 数据验证
            try:
                for dv in ws.data_validations.dataValidation:
                    dst_ws.add_data_validation(copy.copy(dv))
            except Exception:
                pass

            dst_wb.save(dst_file)
            print("    OK %s.xlsx（%d 行）" % (fname, len(src_rows)))
            success_count += 1

        except PermissionError:
            print("    FAIL %s.xlsx -> 失败！文件被占用或无写入权限" % fname)
            print("      解决：检查目录写入权限，或确认文件未被其他程序打开")
            fail_count += 1
        except OSError as e:
            print("    FAIL %s.xlsx -> 失败！系统错误" % fname)
            print("      错误代码：%s" % e.errno)
            print("      错误详情：%s" % e)
            if 'name too long' in str(e).lower() or e.errno == 36:
                print("      原因分析：文件名过长。原始值[%s]超出系统限制（Linux一般255字节）。" % group_name)
            elif 'no space' in str(e).lower() or e.errno == 28:
                print("      原因分析：磁盘空间不足，请清理磁盘后重试。")
            elif e.errno == 13:
                print("      原因分析：权限不足，请检查当前目录的写入权限（试试 chmod 777）。")
            elif e.errno == 30:
                print("      原因分析：只读文件系统，无法写入。请换一个可写的目录。")
            else:
                print("      原因分析：可能是磁盘权限问题或路径问题。")
            fail_count += 1
        except Exception as e:
            print("    FAIL %s.xlsx -> 失败！" % fname)
            print("      错误类型：%s" % type(e).__name__)
            print("      错误详情：%s" % e)
            fail_count += 1

    # 结果汇总
    print()
    print("=" * 60)
    if fail_count == 0:
        print("  OK 拆分完成！成功生成 %d 个文件" % success_count)
    else:
        print("  !! 拆分完成：成功 %d 个，失败 %d 个" % (success_count, fail_count))
    print("  文件保存在：%s" % work_dir)
    print("=" * 60)

    if nested_pairs:
        print()
        print("  !! 再次提醒：以下名称存在包含关系，已拆分为不同文件：")
        for short_name, long_name in nested_pairs:
            print("    [%s] 和 [%s] -> 分别生成了独立文件" % (short_name, long_name))
        print("  如果不符合预期，请在原表中统一名称后重新拆分。")

    return fail_count == 0


def main():
    print()
    print("+" + "=" * 46 + "+")
    print("|        Excel 拆分工具  v1.1 (Linux)         |")
    print("|  按指定列拆分xlsx，保留原表所有格式         |")
    print("+" + "=" * 46 + "+")
    print()

    # PyInstaller 打包后 __file__ 指向临时解压目录，需要用 sys.executable 的目录
    if getattr(sys, 'frozen', False):
        work_dir = os.path.dirname(os.path.abspath(sys.executable))
    else:
        work_dir = os.path.dirname(os.path.abspath(__file__))
    print("工作目录：%s" % work_dir)

    # 查找xlsx文件
    src_file, err = find_xlsx_file(work_dir)
    if err:
        print("\n  x %s" % err)
        print("  请将此脚本放到 .xlsx 文件所在的文件夹中，再运行。")
        input("\n按回车键退出...")
        sys.exit(1)

    print("已找到文件：%s" % os.path.basename(src_file))

    # 打开文件
    print("正在读取，请稍候...")
    try:
        wb = load_workbook(src_file)
    except PermissionError:
        print("\n  x 无法打开文件！没有读取权限。")
        print("  请检查文件权限：ls -la %s" % src_file)
        print("  或尝试：chmod 644 %s" % src_file)
        input("\n按回车键退出...")
        sys.exit(1)
    except Exception as e:
        print("\n  x 读取文件失败！")
        print("  文件路径：%s" % src_file)
        print("  错误类型：%s" % type(e).__name__)
        print("  错误详情：%s" % e)
        print("\n  可能原因：")
        print("    1. 文件不是有效的 .xlsx 格式（可能是 .xls 旧格式，需另存为 .xlsx）")
        print("    2. 文件已损坏（尝试重新从源头获取文件）")
        print("    3. 文件有密码保护（请先去除密码）")
        print("    4. openpyxl版本过低（当前版本可通过 pip3 show openpyxl 查看）")
        print("\n  完整错误堆栈：")
        print("-" * 60)
        traceback.print_exc()
        print("-" * 60)
        input("\n按回车键退出...")
        sys.exit(1)

    # 选择Sheet
    sheet_names = wb.sheetnames
    if len(sheet_names) == 1:
        ws = wb.active
        print("Sheet：[%s]" % ws.title)
    else:
        print("\n" + "=" * 60)
        print("  该文件包含 %d 个Sheet，请选择要拆分的Sheet：" % len(sheet_names))
        print("=" * 60)
        for i, name in enumerate(sheet_names):
            tmp_ws = wb[name]
            r = tmp_ws.max_row if tmp_ws.max_row else 0
            c = tmp_ws.max_column if tmp_ws.max_column else 0
            print("  %d. [%s]  (%d行 x %d列)" % (i + 1, name, r, c))
        print()
        while True:
            choice = input("请选择Sheet编号 (1-%d): " % len(sheet_names)).strip()
            try:
                idx = int(choice)
                if 1 <= idx <= len(sheet_names):
                    ws = wb[sheet_names[idx - 1]]
                    print("\n  OK 已选择Sheet：[%s]" % ws.title)
                    break
                else:
                    print("  x 请输入 1 到 %d 之间的数字" % len(sheet_names))
            except ValueError:
                print("  x 请输入数字，不要输入其他字符")

    max_row = ws.max_row
    max_col = ws.max_column

    if max_row is None or max_row < 1:
        print("\n  x 表格为空，没有任何数据！")
        input("\n按回车键退出...")
        sys.exit(1)

    if max_col is None or max_col < 1:
        print("\n  x 表格没有任何列！")
        input("\n按回车键退出...")
        sys.exit(1)

    print("读取成功！表格共 %d 行，%d 列，Sheet名称：[%s]\n" % (max_row, max_col, ws.title))

    if max_row < 2:
        print("\n  x 表格只有 %d 行，至少需要1行标题 + 1行数据才能拆分！" % max_row)
        input("\n按回车键退出...")
        sys.exit(1)

    # 交互问答
    header_row = ask_header_row(ws, max_col)
    print("\n  OK 标题行：第 %d 行" % header_row)
    print("  OK 数据行：从第 %d 行开始" % (header_row + 1))

    split_col = ask_split_column(ws, header_row, max_col)

    # 执行拆分
    try:
        do_split(src_file, ws, header_row, split_col, max_row, max_col)
    except Exception as e:
        print("\n  x 拆分过程中发生未预期的错误！")
        print("  错误类型：%s" % type(e).__name__)
        print("  错误详情：%s" % e)
        print("\n  完整错误堆栈（可截图发给开发者排查）：")
        print("-" * 60)
        traceback.print_exc()
        print("-" * 60)

    input("\n按回车键退出...")


if __name__ == '__main__':
    main()
