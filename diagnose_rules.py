#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
规则诊断工具 - 帮助定位规则不生效的原因
"""

import sys
import json
from rule_engine import RuleEngine

def test_multi_dept_rule():
    """测试多科室联合规则"""
    print("\n" + "=" * 70)
    print("诊断1：多科室联合规则")
    print("=" * 70)

    engine = RuleEngine()

    # 读取配置文件中的第一条多科室规则
    multi_rules = engine.rules.get('multi_dept_rules', [])
    if not multi_rules:
        print("[错误] 没有找到多科室联合规则！")
        return

    rule = multi_rules[0]
    print(f"\n规则ID: {rule['id']}")
    print(f"规则名称: {rule['name']}")
    print(f"启用状态: {rule['enabled']}")
    print(f"优先级: {rule['priority']}")

    conditions = rule.get('conditions', {})
    print("\n条件配置：")
    print(f"  filename_contains_all: {conditions.get('filename_contains_all', '未设置')}")
    print(f"  order: {conditions.get('order', '未设置')}")
    print(f"  dept_keywords_order: {conditions.get('dept_keywords_order', '未设置')}")

    # 测试匹配
    print("\n测试文件匹配：")
    test_files = [
        "医政疾控.doc",
        "疾控医政.doc",
    ]

    for filename in test_files:
        print(f"\n  文件名: {filename}")

        # 检查关键词
        matched_keywords = engine.find_keywords_with_position(filename)
        print(f"    找到的关键词: {[k['keyword'] for k in matched_keywords]}")

        # 检查条件
        if engine.check_condition(conditions, filename, "测试标题"):
            print(f"    [成功] 条件匹配")
            result = engine.match_multi_dept_rules(filename, "测试标题")
            if result:
                print(f"    拟办意见: {result.get('拟办意见', '')[:60]}...")
        else:
            print(f"    [失败] 条件不匹配")

            # 详细检查每个条件
            if 'filename_contains_all' in conditions:
                required = conditions['filename_contains_all']
                found = [kw for kw in required if kw in filename]
                print(f"      需要关键词: {required}")
                print(f"      找到关键词: {found}")
                if len(found) < len(required):
                    print(f"      缺少关键词: {set(required) - set(found)}")

            if 'order' in conditions:
                expected_order = conditions['order']
                actual_order = [k['keyword'] for k in matched_keywords if k['keyword'] in expected_order]
                print(f"      期望顺序: {expected_order}")
                print(f"      实际顺序: {actual_order}")
                if actual_order != expected_order:
                    print(f"      顺序不匹配！")

def test_public_place_rule():
    """测试公共场所规则"""
    print("\n" + "=" * 70)
    print("诊断2：公共场所识别码规则")
    print("=" * 70)

    engine = RuleEngine()

    # 查找公共场所规则
    special_rules = engine.rules.get('special_rules', [])
    public_place_rule = None
    for rule in special_rules:
        if rule['id'] == 'special_public_place':
            public_place_rule = rule
            break

    if not public_place_rule:
        print("[错误] 没有找到公共场所规则！")
        return

    print(f"\n规则ID: {public_place_rule['id']}")
    print(f"规则名称: {public_place_rule['name']}")
    print(f"启用状态: {public_place_rule['enabled']}")
    print(f"优先级: {public_place_rule['priority']}")

    conditions = public_place_rule.get('conditions', {})
    print("\n条件配置：")
    print(f"  filename_contains_any: {conditions.get('filename_contains_any', [])}")

    identifiers = conditions.get('file_title_contains_any', [])
    print(f"\n  识别码列表 (共{len(identifiers)}个):")
    for i, kw in enumerate(identifiers, 1):
        print(f"    {i}. {kw}")

    # 检查是否有"洗脚"
    if "洗脚" in identifiers:
        print("\n  [成功] 识别码列表中包含'洗脚'")
    else:
        print("\n  [失败] 识别码列表中不包含'洗脚'")

    # 检查modifier
    modifier = public_place_rule.get('modifier', {})
    print(f"\n修正规则配置：")
    opinion_mods = modifier.get('拟办意见', [])
    for mod in opinion_mods:
        action = mod.get('action', '')
        if action == 'remove':
            print(f"  删除: '{mod.get('text', '')}'")

    # 测试匹配
    print("\n测试文件匹配：")
    test_cases = [
        ("疾控.doc", "某洗脚店卫生投诉"),
        ("监督.doc", "某酒店卫生投诉"),
        ("医政.doc", "某洗脚店卫生投诉"),  # 不应该匹配（文件名不含疾控/监督）
    ]

    for filename, title in test_cases:
        print(f"\n  文件名: {filename}, 标题: {title}")

        # 检查条件
        if engine.check_condition(conditions, filename, title):
            print(f"    [成功] 条件匹配")

            # 先获取基础规则
            base_result = engine.match_single_dept_rules(filename, title)
            if base_result:
                print(f"    初始拟办意见: {base_result.get('拟办意见', '')[:60]}...")

                # 应用修正
                special_result = engine.match_special_rules(filename, title)
                if special_result and '_modifier' in special_result:
                    modified = engine.apply_modifier(base_result, special_result['_modifier'])
                    print(f"    修正后拟办意见: {modified.get('拟办意见', '')[:60]}...")

                    # 检查是否成功删除
                    if '，请法监科亦阅' in base_result.get('拟办意见', ''):
                        if '，请法监科亦阅' not in modified.get('拟办意见', ''):
                            print(f"    [成功] 已删除'，请法监科亦阅'")
                        else:
                            print(f"    [失败] 未能删除'，请法监科亦阅'")
        else:
            print(f"    [失败] 条件不匹配")

            # 详细检查
            filename_ok = any(kw in filename for kw in conditions.get('filename_contains_any', []))
            title_ok = any(kw in title for kw in conditions.get('file_title_contains_any', []))

            print(f"      文件名检查: {'通过' if filename_ok else '失败'}")
            print(f"      标题检查: {'通过' if title_ok else '失败'}")

def main():
    print("\n")
    print("*" * 70)
    print("*" + " " * 68 + "*")
    print("*" + "规则诊断工具".center(66) + "*")
    print("*" + " " * 68 + "*")
    print("*" * 70)

    try:
        test_multi_dept_rule()
        test_public_place_rule()

        print("\n" + "=" * 70)
        print("诊断完成")
        print("=" * 70)
        print("\n如果规则配置正确但程序中仍不生效，请：")
        print("1. 完全关闭程序")
        print("2. 重新启动程序")
        print("3. 确保文件名和标题完全符合规则条件")
        print()

    except Exception as e:
        print(f"\n[错误] 诊断过程中出现异常: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
