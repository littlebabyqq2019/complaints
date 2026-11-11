#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
调试规则匹配过程
"""

from rule_engine import RuleEngine

def test_file(filename, title):
    """测试单个文件"""
    print('=' * 70)
    print(f'测试文件: {filename}')
    print(f'文件标题: {title}')
    print('-' * 70)

    engine = RuleEngine()

    # 匹配规则
    result = engine.match_filling_rule(filename, title)

    if result:
        print(f'\n匹配结果:')
        print(f'  拟办意见: {result.get("拟办意见", "(空)")}')
        print(f'  办理结果: {result.get("办理结果", "(空)")}')
    else:
        print('\n未匹配到任何规则')
    print()

def main():
    print('\n调试规则匹配\n')

    # 测试三个文件
    test_file(
        '医政、疾控 热线工单DH6101002026091610366.docx',
        '投诉医院不退费的问题'
    )

    test_file(
        '疾控 热线工单SZZX6101002026090908045.docx',
        '洗脚房'
    )

    test_file(
        '疾控、医政 热线工单JQ6101002026091013068.docx',
        '投诉问题'
    )

if __name__ == "__main__":
    main()
