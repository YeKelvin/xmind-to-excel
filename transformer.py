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
    name = datetime.now().strftime(r'[%m-%d][%H.%M.%S]') + name
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


TAGS = ['path', 'func', 'title', 'pre', 'step', 'exp']


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


def topics_to_rows(topics: list, rows: list, metadata: dict) -> None:
    for topic in topics:
        # 解析主题
        has_tag, tag, text = parse_topic(topic)
        # 添加用例原始数据
        has_tag and metadata[tag].append(text)
        # 存在子 topic 时，递归解析
        if 'topics' in topic:
            topics_to_rows(topic['topics'], rows, metadata)
        # 遍历至 topic 路径末端时，组装数据并添加至用例集
        else:
            # topic 路径上存在 title 才识别为一条用例
            if metadata['title']:
                path = '-'.join(metadata['path'])
                func = '-'.join(metadata['func'])
                title = '-'.join(metadata['title'])

                # 抵达 topic 路径末端时，判断用例是否已存在，存在则拼接预期结果，不存在则添加用例
                # path、func 和 title 相同代表末端有多个 exp （预期结果）
                match = [row for row in rows if row['path'] == path and row['func'] == func and row['title'] == title]
                if match:
                    existed_row = match[0]
                    existed_row['exp'] = existed_row['exp'] + '\n' + '-'.join(metadata['exp'])
                else:
                    rows.append({
                        'path': path,
                        'func': func,
                        'title': title,
                        'pre': '-'.join(metadata['pre']),
                        'step': '-'.join(metadata['step']),
                        'exp': '-'.join(metadata['exp'])
                    })
        # 回溯时删除数据
        has_tag and metadata[tag].pop()


def write_to_excel_by_testcase(file_path, sheet_name, rows: list, spec=None):
    """写入excel

    Args:
        file_path (str): excel 文件路径
        sheet_name (str): sheet 页名称
        rows (list): 测试用例数据
        spec (str, optional): MacOS下 excelApp 的名称. e.g.: wpsoffice
    """
    if spec and platform.system().lower() == 'darwin':
        app = xw.App(spec=spec, add_book=False)
        wb = app.books.open(file_path)
        # 禁用提示和屏幕刷新可以提升速度
        # app.display_alerts = False
        # app.screen_updating = False
    else:
        wb = xw.Book(file_path)

    sheet = wb.sheets[sheet_name]
    # 遍历写入数据
    for rownum, testcase in enumerate(rows):
        print(f'写入 Testcas: {testcase}')
        rownum = rownum + 2
        # 目录
        sheet.range(f'A{rownum}').value = testcase['path']
        # 标题
        sheet.range(f'B{rownum}').value = testcase['title']
        # 前置条件
        sheet.range(f'D{rownum}').value = testcase['pre']
        # 步骤
        sheet.range(f'E{rownum}').value = testcase['step']
        # 预期结果
        sheet.range(f'F{rownum}').value = testcase['exp']
        # 用例类型
        sheet.range(f'G{rownum}').value = '功能测试'
        # 用例状态
        sheet.range(f'H{rownum}').value = '正常'
        # 功能点
        sheet.range(f'K{rownum}').value = testcase['func']

    # 自动调整单元格大小
    sheet.autofit()
    # 保存
    wb.save()


def xmind_to_excel_for_tapd(xmind_file_path, xmind_sheet_name):
    # 解析 XMind
    xmind_sheet = parse_xmind_by_sheet(xmind_file_path, xmind_sheet_name)
    # 获取根主题下的子节点
    topics = xmind_sheet['topic']['topics']
    root_path = xmind_sheet['topic']['title']
    rows = []
    metadata = {
        'path': [root_path],
        'func': [],
        'title': [],
        'pre': [],
        'step': [],
        'exp': []
    }
    # 校验 XMind 各个主题是否符合用例格式
    check_topics_format(topics)
    # 主题转用例数据
    topics_to_rows(topics, rows, metadata)
    print(f'XMind 解析完成，总计 {len(rows)} 条用例')
    # [print(row) for row in rows]  # debug print
    # 复制测试用例模板文件
    template_file_path = os.path.join(PROJECT_PATH, 'testcase.template.tapd.xlsx')
    output_file_path = copy_file_to_output(template_file_path, f'[tapd]{xmind_sheet_name}.xlsx')
    print('写入 Excel 开始')
    write_to_excel_by_testcase(output_file_path, '测试用例', rows)
    print('写入 Excel 完成')
    print(f'Excel路径: {output_file_path}')


if __name__ == '__main__':
    xmind_file_path = r'xxx'
    xmind_sheet_name = 'xxx'
    xmind_to_excel_for_tapd(xmind_file_path, xmind_sheet_name)
