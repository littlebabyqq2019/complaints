#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
详细读取生成的批办单内容
"""

from docx import Document
import os

def read_complaint_doc(filepath):
    """读取批办单详细内容"""
    try:
        doc = Document(filepath)

        print('=' * 70)
        print(f'文件: {os.path.basename(filepath)}')
        print('-' * 70)

        if not doc.tables:
            print('[错误] 文档中没有表格')
            return

        table = doc.tables[0]

        # 提取所有字段
        data = {}
        for row in table.rows:
            cells = row.cells
            for i in range(0, len(cells)-1, 2):  # 每两列一组
                if i+1 < len(cells):
                    key = cells[i].text.strip()
                    value = cells[i+1].text.strip()
                    if key and key != '':
                        # 移除末尾的冒号
                        key = key.rstrip('：:')
                        data[key] = value

        # 显示关键字段
        print(f"\n来文标题: {data.get('来文标题', '(空)')}")
        print(f"文件名: {data.get('文件名', '(空)')}")
        print(f"文件标题: {data.get('文件标题', '(空)')}")
        print(f"\n承办意见:\n{data.get('承办意见', '(空)')}")
        print(f"\n办理情况:\n{data.get('办理情况', '(空)')}")

        print()

    except Exception as e:
        print(f'[错误] 读取失败: {e}')
        import traceback
        traceback.print_exc()

def main():
    print("\n读取生成的批办单文件\n")

    output_dir = '待打印'

    if not os.path.exists(output_dir):
        print(f'[错误] 找不到文件夹: {output_dir}')
        return

    files = sorted([f for f in os.listdir(output_dir) if f.endswith('.docx')])

    if not files:
        print(f'[错误] {output_dir}/ 文件夹中没有docx文件')
        return

    for filename in files:
        filepath = os.path.join(output_dir, filename)
        read_complaint_doc(filepath)

if __name__ == "__main__":
    main()
