#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
修复多科室联合规则和公共场所识别码保存问题
"""

# 问题1：多科室联合规则保存时，需要同步更新规则引擎使用的字段
# 问题2：公共场所识别码添加/删除后，需要立即保存到配置文件

import sys
import os

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

print("修复规则管理功能...")
print("=" * 60)

# 读取 rule_management_ui.py
with open('rule_management_ui.py', 'r', encoding='utf-8') as f:
    content = f.read()

# 修复1：apply_multi_dept_rule 方法 - 同步更新规则引擎字段
old_apply_multi = '''        if dept_keywords_order:
            self.current_rule['conditions']['dept_keywords_order'] = dept_keywords_order

        filename_not_contains = self.multi_dept_not_filename_entry.get().strip()'''

new_apply_multi = '''        if dept_keywords_order:
            self.current_rule['conditions']['dept_keywords_order'] = dept_keywords_order

            # 同步更新规则引擎使用的字段
            # filename_contains_all: 文件名必须包含所有关键词
            all_keywords = []
            for kw_list in dept_keywords_order:
                all_keywords.extend(kw_list)
            self.current_rule['conditions']['filename_contains_all'] = all_keywords

            # order: 关键词在文件名中的出现顺序（使用每个位置的第一个关键词）
            if len(dept_keywords_order) >= 2:
                self.current_rule['conditions']['order'] = [
                    dept_keywords_order[0][0],  # 第一位的第一个关键词
                    dept_keywords_order[1][0]   # 第二位的第一个关键词
                ]

        filename_not_contains = self.multi_dept_not_filename_entry.get().strip()'''

if old_apply_multi in content:
    content = content.replace(old_apply_multi, new_apply_multi)
    print("✓ 已修复多科室联合规则保存逻辑")
else:
    print("⚠ 未找到多科室联合规则保存代码，可能已经修复")

# 修复2：add_public_place_keyword 方法 - 立即保存
old_add_keyword = '''        # 添加到列表
        self.public_place_listbox.insert(tk.END, keyword)
        self.public_place_add_entry.delete(0, tk.END)
        messagebox.showinfo("成功", f"已添加识别码：{keyword}")'''

new_add_keyword = '''        # 添加到列表
        self.public_place_listbox.insert(tk.END, keyword)
        self.public_place_add_entry.delete(0, tk.END)

        # 立即保存到配置文件
        if self.current_rule and self.current_rule.get('id') == 'special_public_place':
            # 获取所有识别码
            all_keywords = list(self.public_place_listbox.get(0, tk.END))

            # 更新规则条件
            conditions = self.current_rule.get('conditions', {})
            conditions['file_title_contains_any'] = all_keywords
            self.current_rule['conditions'] = conditions

            # 保存到配置文件
            self.engine.save_rules()

            # 重新加载规则以确保同步
            self.engine.load_rules()

            messagebox.showinfo("成功", f"已添加识别码：{keyword}\\n配置已自动保存")
        else:
            messagebox.showinfo("成功", f"已添加识别码：{keyword}")'''

if old_add_keyword in content:
    content = content.replace(old_add_keyword, new_add_keyword)
    print("✓ 已修复公共场所识别码添加功能")
else:
    print("⚠ 未找到识别码添加代码，可能已经修复")

# 修复3：remove_public_place_keyword 方法 - 立即保存
old_remove_keyword = '''        keyword = self.public_place_listbox.get(selection[0])
        if messagebox.askyesno("确认", f"确定要删除识别码 \\"{keyword}\\" 吗？"):
            self.public_place_listbox.delete(selection[0])
            messagebox.showinfo("成功", f"已删除识别码：{keyword}")'''

new_remove_keyword = '''        keyword = self.public_place_listbox.get(selection[0])
        if messagebox.askyesno("确认", f"确定要删除识别码 \\"{keyword}\\" 吗？"):
            self.public_place_listbox.delete(selection[0])

            # 立即保存到配置文件
            if self.current_rule and self.current_rule.get('id') == 'special_public_place':
                # 获取所有识别码
                all_keywords = list(self.public_place_listbox.get(0, tk.END))

                # 更新规则条件
                conditions = self.current_rule.get('conditions', {})
                conditions['file_title_contains_any'] = all_keywords
                self.current_rule['conditions'] = conditions

                # 保存到配置文件
                self.engine.save_rules()

                # 重新加载规则以确保同步
                self.engine.load_rules()

                messagebox.showinfo("成功", f"已删除识别码：{keyword}\\n配置已自动保存")
            else:
                messagebox.showinfo("成功", f"已删除识别码：{keyword}")'''

if old_remove_keyword in content:
    content = content.replace(old_remove_keyword, new_remove_keyword)
    print("✓ 已修复公共场所识别码删除功能")
else:
    print("⚠ 未找到识别码删除代码，可能已经修复")

# 保存修改后的文件
with open('rule_management_ui.py', 'w', encoding='utf-8') as f:
    f.write(content)

print("=" * 60)
print("修复完成！")
print("\n修复内容：")
print("1. 多科室联合规则保存时，自动同步更新规则引擎使用的字段：")
print("   - filename_contains_all：包含所有关键词")
print("   - order：关键词顺序")
print("2. 公共场所识别码添加/删除后，立即保存到配置文件")
print("3. 保存后自动重新加载规则，确保同步")
print("\n请重启程序测试！")
