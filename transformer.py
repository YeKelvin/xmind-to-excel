#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @File    : transformer.py
# @Time    : 2021-11-08 14:06:05
# @Author  : Kelvin.Ye
import os
import platform
import shutil
import sys
from datetime import datetime

import xlwings as xw
from xmindparser import xmind_to_dict


# 添加项目路径到 system-path
sys.path.append(os.path.dirname(sys.path[0]))

# 项目路径
PROJECT_PATH = os.path.abspath(os.path.dirname(__file__))


def copy_file_to_output(source: str, target_name: str = None) -> str:
    """复制文件至 output 目录

    Args:
        source (str): 源文件路径
        target_name (str, optional): 目标文件名称（需要文件后缀）. Defaults to None.

    Raises:
        Exception:
            1、源文件不是文件
            2、目标文件已存在

    Returns:
        str: 复制后文件的路径
    """
    # 判断是否为文件
    if not os.path.isfile(source):
        raise Exception(f'{source} 非文件')

    # 判断 output 目录是否存在，不存在则新建
    output_path = os.path.join(PROJECT_PATH, 'output')
    if not os.path.exists(output_path):
        os.mkdir(output_path)

    # 存在 target_name 时修改复制后的文件名为 target_name
    file_name = target_name
    if not file_name:
        file_name = os.path.split(source)[1]
    name, ext = os.path.splitext(file_name)
    name = datetime.now().strftime(r'[%Y-%m-%d_%H.%M.%S]') + name
    file_name = name + '.xlsx'
    target_file_path = os.path.join(output_path, file_name)

    # 判断目标文件是否存在
    if os.path.exists(target_file_path):
        raise Exception(f'{target_file_path} 文件已存在')

    # 复制文件
    shutil.copyfile(source, target_file_path)
    return target_file_path


def parse_xmind_by_sheet(file_path: str, sheet_name: str) -> dict:
    """解析 xmind 文件里指定的 sheet 页

    Args:
        file_path (str): xmind 文件路径
        sheet_name (str): xmind sheet 名称

    Raises:
        Exception: sheet 页不存在

    Returns:
        dict: xmind-dict
    """
    sheets = xmind_to_dict(file_path)
    specified_sheet = [sheet for sheet in sheets if sheet['title'] == sheet_name]
    if not specified_sheet:
        raise Exception(f'sheet页:[ {sheet_name} ] 不存在')
    sheet = specified_sheet[0]
    return sheet


def check_topics_format(topics: list):
    for topic in topics:
        # 中文冒号替换为英文冒号
        title = topic['title'].replace('：', ':')
        if ':' not in title and (
            title.startswith('module') or  # noqa
            title.startswith('path') or  # noqa
            title.startswith('func') or  # noqa
            title.startswith('title') or  # noqa
            title.startswith('pre') or  # noqa
            title.startswith('step') or  # noqa
            title.startswith('exp')  # noqa
        ):
            raise Exception(f'topic:[ {title} ] 格式不正确')
        if 'topics' in topic:
            check_topics_format(topic['topics'])


TAGS = ['module', 'path', 'func', 'title', 'pre', 'step', 'exp']


def parse_topic(topic):
    # 中文冒号替换为英文冒号
    data = topic['title'].replace('：', ':')
    # 分割标签和内容
    splits = data.split(':')
    has_tag = False
    tag = ''
    text = ''
    if len(splits) >= 2:
        tag = splits[0]
        tag = tag.strip()  # 移除首尾空格
        if tag in TAGS:
            has_tag = True
            text = ':'.join(splits[1:])
            text = text.strip()  # 移除首尾空格
    return has_tag, tag, text


def topics_to_rows(topics: list, rows: list, metadata: dict, classified_data: dict = None) -> None:
    for topic in topics:
        # 解析主题
        has_tag, tag, text = parse_topic(topic)
        # 添加用例原始数据
        has_tag and metadata[tag].append(text)
        # 存在子 topic 时，递归解析
        if 'topics' in topic:
            topics_to_rows(topic['topics'], rows, metadata, classified_data)
        # 遍历至 topic 路径末端时，组装数据并添加至用例集
        else:
            # topic 路径上存在 title 才识别为一条用例
            if metadata['title']:
                module = '-'.join(metadata['module'])
                full_path = metadata['root'] + '-' + module + '-' + '-'.join(metadata['path'])
                func = '-'.join(metadata['func'])
                title = '-'.join(metadata['title'])

                # 抵达 topic 路径末端时，判断用例是否已存在，存在则拼接预期结果，不存在则添加用例
                # path、func 和 title 相同代表末端有多个 exp （预期结果）
                match = [row for row in rows if row['path'] == full_path and row['func'] == func and row['title'] == title]
                if match:
                    existed_row = match[0]
                    existed_row['exp'] = existed_row['exp'] + '\n' + '-'.join(metadata['exp'])
                else:
                    rows.append({
                        'path': full_path,
                        'func': func,
                        'title': title,
                        'pre': '-'.join(metadata['pre']),
                        'step': '-'.join(metadata['step']),
                        'exp': '-'.join(metadata['exp'])
                    })
                    # 分类 module 到不同的 sheet 页
                    if classified_data is not None:
                        sheet_rows = classified_data.get(module, [])
                        sheet_rows.append({
                            'path': '-'.join(metadata['path']),
                            'func': func,
                            'title': title,
                            'pre': '-'.join(metadata['pre']),
                            'step': '-'.join(metadata['step']),
                            'exp': '-'.join(metadata['exp'])
                        })
                        classified_data[module] = sheet_rows
        # 回溯时删除数据
        has_tag and metadata[tag].pop()


def open_excel(file_path, spec=None):
    if spec and platform.system().lower() == 'darwin':
        app = xw.App(spec=spec, add_book=False)
        wb = app.books.open(file_path)
        # 禁用提示和屏幕刷新可以提升速度
        # app.display_alerts = False
        # app.screen_updating = False
    else:
        wb = xw.Book(file_path)

    return wb


def add_used_range_borders(sheet):
    sheet.used_range.api.Borders(8).LineStyle = 1  # 上边框
    sheet.used_range.api.Borders(9).LineStyle = 1  # 下边框
    sheet.used_range.api.Borders(7).LineStyle = 1  # 左边框
    sheet.used_range.api.Borders(10).LineStyle = 1  # 右边框
    sheet.used_range.api.Borders(12).LineStyle = 1  # 内横边框
    sheet.used_range.api.Borders(11).LineStyle = 1  # 内纵边框


def write_to_excel_by_testcase(file_path, sheet_name, rows: list):
    """写入excel

    Args:
        file_path (str): excel 文件路径
        sheet_name (str): sheet 页名称
        rows (list): 测试用例数据
        spec (str, optional): MacOS下 excelApp 的名称. e.g.: wpsoffice
    """
    wb = open_excel(file_path)
    sheet = wb.sheets[sheet_name]
    # 遍历写入数据
    for rownum, testcase in enumerate(rows):
        rownum = rownum + 2
        print(f'No.{rownum} Testcas: {testcase}')
        # 用例类型
        sheet.range(f'A{rownum}').value = '功能测试'
        # 目录
        sheet.range(f'B{rownum}').value = testcase['path']
        # 功能点
        sheet.range(f'C{rownum}').value = testcase['func']
        # 用例名称
        sheet.range(f'D{rownum}').value = testcase['title']
        # 前置条件
        sheet.range(f'E{rownum}').value = testcase['pre']
        # 用例步骤
        sheet.range(f'F{rownum}').value = testcase['step']
        # 预期结果
        sheet.range(f'G{rownum}').value = testcase['exp']

    # 自动调整单元格大小
    sheet.autofit()
    # 保存
    wb.save()


def classify_testcase_to_excel(file_path: str, classified_data: dict):
    wb = open_excel(file_path)
    template_sheet = wb.sheets['模板']
    # 遍历写入不同模块的测试用例
    for module, rows in classified_data.items():
        # 从模板复制一个 sheet 页并修改为模块的名称
        template_sheet.copy(before=template_sheet, name=module)
        # 打开复制后的 sheet 页
        sheet = wb.sheets[module]
        # 遍历写入测试用例
        for rownum, testcase in enumerate(rows):
            rownum = rownum + 2
            print(f'module:[{module}] No.{rownum} Testcas: {testcase}')
            # 用例类型
            sheet.range(f'A{rownum}').value = '功能测试'
            # 目录
            sheet.range(f'B{rownum}').value = testcase['path']
            # 功能点
            sheet.range(f'C{rownum}').value = testcase['func']
            # 用例名称
            sheet.range(f'D{rownum}').value = testcase['title']
            # 前置条件
            sheet.range(f'E{rownum}').value = testcase['pre']
            # 用例步骤
            sheet.range(f'F{rownum}').value = testcase['step']
            # 预期结果
            sheet.range(f'G{rownum}').value = testcase['exp']
        # 删除不需要的实际结果列
        delete_actual_results_column_by_module(sheet)
        # 添加边框
        add_used_range_borders(sheet)
        # 自动调整单元格大小
        sheet.autofit()
    # 所有模块写入完成后删除模板页
    template_sheet.delete()
    # 保存
    wb.save()


ACTUAL_RESULTS_COLUMNS = {
    'DEFAULT': 'H:H',
    # 如果需要统计其他端时，默认（DEFAULT）实际结果列会被删掉，所以其他端的列还是从 H 列开始
    'ANDROID': 'H:H',
    'IOS': 'I:I',
    'H5': 'J:J'
}


def delete_actual_results_column_by_module(sheet):
    terminal_total = 0
    del_android = True
    del_ios = True
    del_h5 = True

    if 'APP' in sheet.name:
        terminal_total += 2
        del_android = False
        del_ios = False
    if 'H5' in sheet.name:
        terminal_total += 1
        del_h5 = False

    # 删除不需要的列（要从后面开始删，不然会报错）
    if del_h5:
        sheet.range(ACTUAL_RESULTS_COLUMNS['H5']).api.EntireColumn.Delete()
    if del_ios:
        sheet.range(ACTUAL_RESULTS_COLUMNS['IOS']).api.EntireColumn.Delete()
    if del_android:
        sheet.range(ACTUAL_RESULTS_COLUMNS['ANDROID']).api.EntireColumn.Delete()
    if terminal_total > 0:
        sheet.range(ACTUAL_RESULTS_COLUMNS['DEFAULT']).api.EntireColumn.Delete()


def analysis_testcase_to_excel(file_path: str, classified_data: dict):
    wb = open_excel(file_path)
    analysis_sheet = wb.sheets['数据统计']
    default_column = ACTUAL_RESULTS_COLUMNS['DEFAULT']
    android_column = ACTUAL_RESULTS_COLUMNS['ANDROID']
    ios_column = ACTUAL_RESULTS_COLUMNS['IOS']
    h5_column = ACTUAL_RESULTS_COLUMNS['H5']

    for rownum, module_name in enumerate(classified_data.keys()):
        rownum = rownum + 3
        terminal_total = 0
        count_android = False
        count_ios = False
        count_h5 = False

        if 'APP' in module_name:
            terminal_total += 2
            count_android = True
            count_ios = True
        if 'H5' in module_name:
            terminal_total += 1
            count_h5 = True

        # 案例名称
        analysis_sheet.range(f'A{rownum}').value = module_name
        analysis_sheet.range(f'A{rownum}').api.Font.Bold = True  # 字体加粗
        # 总编写用例数
        analysis_sheet.range(f'B{rownum}').value = f'=IFERROR(COUNTIF(INDIRECT("\'{module_name}\'!D:D"), "*") - 1, 0)'
        # 需执行用例数
        analysis_sheet.range(f'C{rownum}').value = (
            f'=IFERROR(B{rownum} * {terminal_total if terminal_total >0 else 1} - G{rownum}, 0)'
        )
        # 通过、失败、阻塞、不适用
        if terminal_total == 0:
            # 通过
            analysis_sheet.range(f'D{rownum}').value = f'=COUNTIF(INDIRECT("\'{module_name}\'!{default_column}"), "通过")'
            # 失败
            analysis_sheet.range(f'E{rownum}').value = f'=COUNTIF(INDIRECT("\'{module_name}\'!{default_column}"), "失败")'
            # 阻塞
            analysis_sheet.range(f'F{rownum}').value = f'=COUNTIF(INDIRECT("\'{module_name}\'!{default_column}"), "阻塞")'
            # 不适用
            analysis_sheet.range(f'G{rownum}').value = f'=COUNTIF(INDIRECT("\'{module_name}\'!{default_column}"), "不适用")'
        else:
            # 组装公式
            pass_formula = '='
            fail_formula = '='
            block_formula = '='
            invalid_formula = '='
            if count_android:
                pass_formula = pass_formula + '+' + f'COUNTIF(INDIRECT("\'{module_name}\'!{android_column}"), "通过")'
                fail_formula = fail_formula + '+' + f'COUNTIF(INDIRECT("\'{module_name}\'!{android_column}"), "失败")'
                block_formula = block_formula + '+' + f'COUNTIF(INDIRECT("\'{module_name}\'!{android_column}"), "阻塞")'
                invalid_formula = invalid_formula + '+' + f'COUNTIF(INDIRECT("\'{module_name}\'!{android_column}"), "不适用")'
            if count_ios:
                pass_formula = pass_formula + '+' + f'COUNTIF(INDIRECT("\'{module_name}\'!{ios_column}"), "通过")'
                fail_formula = fail_formula + '+' + f'COUNTIF(INDIRECT("\'{module_name}\'!{ios_column}"), "失败")'
                block_formula = block_formula + '+' + f'COUNTIF(INDIRECT("\'{module_name}\'!{ios_column}"), "阻塞")'
                invalid_formula = invalid_formula + '+' + f'COUNTIF(INDIRECT("\'{module_name}\'!{ios_column}"), "不适用")'
            if count_h5:
                pass_formula = pass_formula + '+' + f'COUNTIF(INDIRECT("\'{module_name}\'!{h5_column}"), "通过")'
                fail_formula = fail_formula + '+' + f'COUNTIF(INDIRECT("\'{module_name}\'!{h5_column}"), "失败")'
                block_formula = block_formula + '+' + f'COUNTIF(INDIRECT("\'{module_name}\'!{h5_column}"), "阻塞")'
                invalid_formula = invalid_formula + '+' + f'COUNTIF(INDIRECT("\'{module_name}\'!{h5_column}"), "不适用")'
            # 通过
            analysis_sheet.range(f'D{rownum}').value = pass_formula
            # 失败
            analysis_sheet.range(f'E{rownum}').value = fail_formula
            # 阻塞
            analysis_sheet.range(f'F{rownum}').value = block_formula
            # 不适用
            analysis_sheet.range(f'G{rownum}').value = invalid_formula

        # 未执行
        analysis_sheet.range(f'H{rownum}').value = f'=IFERROR(C{rownum} - (D{rownum} + E{rownum}), 0)'
        # 总完成率
        analysis_sheet.range(f'I{rownum}').value = f'=IFERROR((D{rownum} + E{rownum}) / C{rownum}, 0)'
        analysis_sheet.range(f'I{rownum}').api.NumberFormat = "0%"
        # 总通过率
        analysis_sheet.range(f'J{rownum}').value = f'=IFERROR(D{rownum} / C{rownum}, 0)'
        analysis_sheet.range(f'J{rownum}').api.NumberFormat = "0%"
        # Android通过率
        if count_android:
            analysis_sheet.range(f'K{rownum}').value = (
                f'=IFERROR(COUNTIF(INDIRECT("\'{module_name}\'!{android_column}"), "通过") / (B{rownum}-COUNTIF(INDIRECT("\'{module_name}\'!{android_column}"), "不适用")), 0)'
            )
            analysis_sheet.range(f'K{rownum}').api.NumberFormat = "0%"
        else:
            analysis_sheet.range(f'K{rownum}').value = 'X'
        # IOS通过率
        if count_ios:
            analysis_sheet.range(f'L{rownum}').value = (
                f'=IFERROR(COUNTIF(INDIRECT("\'{module_name}\'!{ios_column}"), "通过") / (B{rownum}-COUNTIF(INDIRECT("\'{module_name}\'!{ios_column}"), "不适用")), 0)'
            )
            analysis_sheet.range(f'L{rownum}').api.NumberFormat = "0%"
        else:
            analysis_sheet.range(f'L{rownum}').value = 'X'
        # H5通过率
        if count_h5:
            analysis_sheet.range(f'M{rownum}').value = (
                f'=IFERROR(COUNTIF(INDIRECT("\'{module_name}\'!{h5_column}"), "通过") / (B{rownum}-COUNTIF(INDIRECT("\'{module_name}\'!{h5_column}"), "不适用")), 0)'
            )
            analysis_sheet.range(f'M{rownum}').api.NumberFormat = "0%"
        else:
            analysis_sheet.range(f'M{rownum}').value = 'X'

    # 总计
    last_rownum = analysis_sheet.used_range.last_cell.row
    total_rownum = last_rownum + 1
    # 案例名称
    analysis_sheet.range(f'A{total_rownum}').value = '总计'
    analysis_sheet.range(f'A{total_rownum}').api.Font.Bold = True  # 字体加粗
    # 总编写用例数
    analysis_sheet.range(f'B{total_rownum}').value = f'=SUM(B3:B{total_rownum - 1})'
    # 需执行用例数
    analysis_sheet.range(f'C{total_rownum}').value = f'=SUM(C3:C{total_rownum - 1})'
    # 通过
    analysis_sheet.range(f'D{total_rownum}').value = f'=SUM(D3:D{total_rownum - 1})'
    # 失败
    analysis_sheet.range(f'E{total_rownum}').value = f'=SUM(E3:E{total_rownum - 1})'
    # 阻塞
    analysis_sheet.range(f'F{total_rownum}').value = f'=SUM(F3:F{total_rownum - 1})'
    # 不适用
    analysis_sheet.range(f'G{total_rownum}').value = f'=SUM(G3:G{total_rownum - 1})'
    # 未执行
    analysis_sheet.range(f'H{total_rownum}').value = f'=SUM(H3:H{total_rownum - 1})'
    # 总完成率
    analysis_sheet.range(f'I{total_rownum}').value = f'=IFERROR((D{total_rownum} + E{total_rownum}) / C{total_rownum}, 0)'
    analysis_sheet.range(f'I{total_rownum}').api.NumberFormat = "0%"
    # 总通过率
    analysis_sheet.range(f'J{total_rownum}').value = f'=IFERROR(D{total_rownum} / C{total_rownum}, 0)'
    analysis_sheet.range(f'J{total_rownum}').api.NumberFormat = "0%"
    # Android通过率
    analysis_sheet.range(f'K{total_rownum}').value = 'X'
    # IOS通过率
    analysis_sheet.range(f'L{total_rownum}').value = 'X'
    # H5通过率
    analysis_sheet.range(f'M{total_rownum}').value = 'X'
    # 添加边框
    add_used_range_borders(analysis_sheet)
    # 保存
    wb.save()


def xmind_to_excel(xmind_file_path: str, xmind_sheet_name: str, classify: bool = False):
    """XMind 转 Excel

    Args:
        xmind_file_path (str): xmind文件路径
        xmind_sheet_name (str): xmind文件sheet页名称
        classify (bool, optional): 是否分类用例
    """
    # 解析 XMind
    xmind_sheet = parse_xmind_by_sheet(xmind_file_path, xmind_sheet_name)
    # 获取根主题下的子节点
    topics = xmind_sheet['topic']['topics']
    root_name = xmind_sheet['topic']['title']
    rows = []
    classified_data = None
    metadata = {
        'root': root_name,
        'module': [],
        'path': [],
        'func': [],
        'title': [],
        'pre': [],
        'step': [],
        'exp': []
    }
    if classify:
        classified_data = {}
    # 校验 XMind 各个主题是否符合用例格式
    check_topics_format(topics)
    # 主题转用例数据
    topics_to_rows(topics, rows, metadata, classified_data)
    print(f'XMind 解析完成，总计 {len(rows)} 条用例')
    # [print(row) for row in rows]  # debug print
    # for module, rows in classified_data.items():  # debug print
    #     print(f'module={module}')
    #     [print(row) for row in rows]
    #     print('\n')
    # 复制测试用例模板文件
    template_file_path = os.path.join(PROJECT_PATH, 'testcase.template.xlsx')
    output_file_path = copy_file_to_output(template_file_path, f'[testcase]{xmind_sheet_name}.xlsx')
    print('写入 Excel 开始')
    write_to_excel_by_testcase(output_file_path, '测试用例', rows)
    if classify:
        classify_testcase_to_excel(output_file_path, classified_data)
        analysis_testcase_to_excel(output_file_path, classified_data)
    print('写入 Excel 完成')
    print(f'Excel路径: {output_file_path}')


if __name__ == '__main__':
    xmind_file_path = r'xxx'
    xmind_sheet_name = 'xxx'
    xmind_to_excel(xmind_file_path, xmind_sheet_name)
