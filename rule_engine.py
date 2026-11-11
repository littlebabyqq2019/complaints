"""
投诉批办单管理系统 - 规则引擎模块
实现可配置的拟办意见和文件标题智能填充规则
"""

import json
import os
import re
import sys
from typing import Dict, List, Optional, Tuple, Any
from datetime import datetime
from platform_compat import get_resource_path, get_user_data_path


class RuleEngine:
    """规则匹配引擎"""

    def __init__(self, config_file: str = None):
        """
        初始化规则引擎

        Args:
            config_file: 规则配置文件路径，默认使用 AppData/filling_rules.json
        """
        if config_file is None:
            # 用户配置文件保存在exe所在目录的AppData下
            config_file = get_user_data_path(os.path.join('AppData', 'filling_rules.json'))

        self.config_file = config_file
        self.rules = None
        self.keywords = None
        self.load_rules()

    def load_rules(self) -> bool:
        """
        加载规则配置文件

        Returns:
            bool: 加载成功返回True，失败返回False
        """
        try:
            # 如果配置文件不存在，复制默认配置
            if not os.path.exists(self.config_file):
                # 默认配置文件从_internal目录加载（只读资源）
                default_file = get_resource_path(os.path.join('AppData', 'filling_rules_default.json'))
                if os.path.exists(default_file):
                    import shutil
                    os.makedirs(os.path.dirname(self.config_file), exist_ok=True)
                    shutil.copy(default_file, self.config_file)
                    print(f"已从默认配置创建规则文件: {self.config_file}")
                else:
                    print(f"警告: 规则配置文件和默认配置均不存在")
                    return False

            with open(self.config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
                self.rules = config.get('rules', {})
                self.keywords = config.get('keywords', {})
                print(f"成功加载规则配置: {self.config_file}")
                return True

        except Exception as e:
            print(f"加载规则配置失败: {e}")
            return False

    def save_rules(self) -> bool:
        """
        保存规则配置到文件

        Returns:
            bool: 保存成功返回True，失败返回False
        """
        try:
            config = {
                'version': '2.0',
                'last_updated': datetime.now().strftime('%Y-%m-%d'),
                'keywords': self.keywords,
                'rules': self.rules
            }

            os.makedirs(os.path.dirname(self.config_file), exist_ok=True)
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)

            print(f"规则配置已保存: {self.config_file}")
            return True

        except Exception as e:
            print(f"保存规则配置失败: {e}")
            return False

    def find_keywords_with_position(self, filename: str) -> List[Dict[str, Any]]:
        """
        查找文件名中所有匹配的科室关键词及其位置

        Args:
            filename: 文件名

        Returns:
            List[Dict]: 匹配的关键词列表，每项包含 keyword, position, length
        """
        if not self.keywords:
            return []

        dept_keywords = self.keywords.get('dept_keywords', [])
        matched = []

        for keyword in dept_keywords:
            pos = filename.find(keyword)
            if pos != -1:
                matched.append({
                    'keyword': keyword,
                    'position': pos,
                    'length': len(keyword)
                })

        # 按位置排序
        matched.sort(key=lambda x: x['position'])
        return matched

    def check_condition(self, condition: Dict[str, Any], filename: str, file_title: str) -> bool:
        """
        检查单个条件是否满足

        Args:
            condition: 条件字典
            filename: 文件名
            file_title: 文件标题

        Returns:
            bool: 条件满足返回True
        """
        # filename_contains: 文件名包含指定关键词
        if 'filename_contains' in condition:
            keywords = condition['filename_contains']
            if not all(kw in filename for kw in keywords):
                return False

        # filename_not_contains: 文件名不包含指定关键词
        if 'filename_not_contains' in condition:
            keywords = condition['filename_not_contains']
            if any(kw in filename for kw in keywords):
                return False

        # filename_contains_any: 文件名包含任一关键词
        if 'filename_contains_any' in condition:
            keywords = condition['filename_contains_any']
            if not any(kw in filename for kw in keywords):
                return False

        # filename_contains_all: 文件名包含所有关键词
        if 'filename_contains_all' in condition:
            keywords = condition['filename_contains_all']
            if not all(kw in filename for kw in keywords):
                return False

        # file_title_contains: 文件标题包含指定关键词
        if 'file_title_contains' in condition:
            keywords = condition['file_title_contains']
            if not all(kw in file_title for kw in keywords):
                return False

        # file_title_contains_any: 文件标题包含任一关键词
        if 'file_title_contains_any' in condition:
            keywords = condition['file_title_contains_any']
            if not any(kw in file_title for kw in keywords):
                return False

        # order: 检查关键词在文件名中的出现顺序
        if 'order' in condition:
            expected_order = condition['order']
            matched_keywords = self.find_keywords_with_position(filename)

            # 提取匹配的关键词名称（按位置顺序）
            actual_order = [item['keyword'] for item in matched_keywords
                          if item['keyword'] in expected_order]

            # 检查是否按预期顺序出现
            if actual_order != expected_order:
                return False

        return True

    def apply_modifier(self, result: Dict[str, str], modifier: Dict[str, Any]) -> Dict[str, str]:
        """
        应用修正规则到结果

        Args:
            result: 当前结果字典
            modifier: 修正规则字典

        Returns:
            Dict: 修正后的结果
        """
        if '拟办意见' in modifier:
            actions = modifier['拟办意见']
            text = result.get('拟办意见', '')

            for action in actions:
                action_type = action.get('action')

                if action_type == 'remove':
                    # 移除指定文本
                    text = text.replace(action['text'], '')

                elif action_type == 'replace':
                    # 替换文本
                    text = text.replace(action['old'], action['new'])

                elif action_type == 'append':
                    # 追加文本
                    text += action['text']

                elif action_type == 'prepend':
                    # 前置文本
                    text = action['text'] + text

            result['拟办意见'] = text

        return result

    def match_special_rules(self, filename: str, file_title: str) -> Optional[Dict[str, str]]:
        """
        匹配特殊情况规则（优先级1-9）

        Args:
            filename: 文件名
            file_title: 文件标题

        Returns:
            Optional[Dict]: 匹配的规则输出，未匹配返回None
        """
        special_rules = self.rules.get('special_rules', [])

        # 按优先级排序
        special_rules = sorted(special_rules, key=lambda x: x.get('priority', 999))

        for rule in special_rules:
            # 跳过禁用的规则
            if not rule.get('enabled', True):
                continue

            conditions = rule.get('conditions', {})

            # 检查条件
            if self.check_condition(conditions, filename, file_title):
                # 如果规则有output，直接返回
                if 'output' in rule:
                    print(f"匹配特殊规则: {rule.get('name', rule.get('id'))}")
                    return rule['output'].copy()

                # 如果规则有modifier，标记需要应用修正（在后续处理）
                if 'modifier' in rule:
                    print(f"匹配特殊修正规则: {rule.get('name', rule.get('id'))}")
                    return {'_modifier': rule['modifier']}

        return None

    def match_multi_dept_rules(self, filename: str, file_title: str) -> Optional[Dict[str, str]]:
        """
        匹配多科室联合规则（优先级10-19）

        Args:
            filename: 文件名
            file_title: 文件标题

        Returns:
            Optional[Dict]: 匹配的规则输出，未匹配返回None
        """
        multi_dept_rules = self.rules.get('multi_dept_rules', [])

        # 按优先级排序
        multi_dept_rules = sorted(multi_dept_rules, key=lambda x: x.get('priority', 999))

        # 获取文件名中所有关键词
        matched_keywords = self.find_keywords_with_position(filename)

        # 至少需要2个关键词才可能匹配多科室规则
        if len(matched_keywords) < 2:
            return None

        for rule in multi_dept_rules:
            # 跳过禁用的规则
            if not rule.get('enabled', True):
                continue

            conditions = rule.get('conditions', {})

            # 检查条件
            if self.check_condition(conditions, filename, file_title):
                print(f"匹配多科室联合规则: {rule.get('name', rule.get('id'))}")
                return rule.get('output', {}).copy()

        return None

    def match_single_dept_rules(self, filename: str, file_title: str) -> Optional[Dict[str, str]]:
        """
        匹配单科室规则（优先级20-29）

        Args:
            filename: 文件名
            file_title: 文件标题

        Returns:
            Optional[Dict]: 匹配的规则输出，未匹配返回None
        """
        single_dept_rules = self.rules.get('single_dept_rules', [])

        # 按优先级排序
        single_dept_rules = sorted(single_dept_rules, key=lambda x: x.get('priority', 999))

        for rule in single_dept_rules:
            # 跳过禁用的规则
            if not rule.get('enabled', True):
                continue

            conditions = rule.get('conditions', {})

            # 检查条件
            if self.check_condition(conditions, filename, file_title):
                print(f"匹配单科室规则: {rule.get('name', rule.get('id'))}")
                return rule.get('output', {}).copy()

        return None

    def apply_title_modifiers(self, filename: str, file_title: str) -> str:
        """
        应用文件标题修正规则

        Args:
            filename: 文件名
            file_title: 文件标题

        Returns:
            str: 修正后的文件标题
        """
        title_modifiers = self.rules.get('title_modifiers', [])

        modified_title = file_title

        for modifier in title_modifiers:
            # 跳过禁用的规则
            if not modifier.get('enabled', True):
                continue

            conditions = modifier.get('conditions', {})

            # 检查条件
            if self.check_condition(conditions, filename, file_title):
                action = modifier.get('action', {})
                action_type = action.get('type')

                if action_type == 'append_if_not_exists':
                    # 如果不存在则追加
                    text = action.get('text', '')
                    if text and text not in modified_title:
                        modified_title += text
                        print(f"应用标题修正: {modifier.get('name', modifier.get('id'))}")

                elif action_type == 'prepend_if_not_exists':
                    # 如果不存在则前置
                    text = action.get('text', '')
                    if text and text not in modified_title:
                        modified_title = text + modified_title
                        print(f"应用标题修正: {modifier.get('name', modifier.get('id'))}")

                elif action_type == 'replace':
                    # 替换
                    old = action.get('old', '')
                    new = action.get('new', '')
                    if old:
                        modified_title = modified_title.replace(old, new)
                        print(f"应用标题修正: {modifier.get('name', modifier.get('id'))}")

        return modified_title

    def match_filling_rule(self, filename: str, file_title: str = '') -> Optional[Dict[str, str]]:
        """
        匹配填充规则（主入口）

        Args:
            filename: 文件名
            file_title: 文件标题

        Returns:
            Optional[Dict]: 包含拟办意见、办理结果等字段的字典，未匹配返回None
        """
        if not self.rules:
            print("警告: 规则配置未加载")
            return None

        result = None
        modifier_to_apply = None

        # 1. 检查特殊情况规则（优先级1-9）
        special_result = self.match_special_rules(filename, file_title)
        if special_result:
            # 如果是修正规则，保存待应用
            if '_modifier' in special_result:
                modifier_to_apply = special_result['_modifier']
            else:
                # 如果是完整规则，直接返回
                result = special_result
                # 检查是否还有其他修正规则需要应用
                for rule in self.rules.get('special_rules', []):
                    if not rule.get('enabled', True):
                        continue
                    if 'modifier' in rule and self.check_condition(rule.get('conditions', {}), filename, file_title):
                        result = self.apply_modifier(result, rule['modifier'])

                # 应用标题修正
                if result:
                    modified_title = self.apply_title_modifiers(filename, file_title)
                    if modified_title != file_title:
                        result['文件标题'] = modified_title

                return result

        # 2. 检查多科室联合规则（优先级10-19）
        if not result:
            result = self.match_multi_dept_rules(filename, file_title)

        # 3. 检查单科室规则（优先级20-29）
        if not result:
            result = self.match_single_dept_rules(filename, file_title)

        # 4. 应用修正规则
        if result and modifier_to_apply:
            result = self.apply_modifier(result, modifier_to_apply)

        # 5. 应用标题修正规则
        if result:
            modified_title = self.apply_title_modifiers(filename, file_title)
            if modified_title != file_title:
                result['文件标题'] = modified_title

        return result

    def test_rule(self, filename: str, file_title: str = '') -> Dict[str, Any]:
        """
        测试规则匹配（用于界面测试功能）

        Args:
            filename: 文件名
            file_title: 文件标题

        Returns:
            Dict: 测试结果，包含匹配状态、规则信息、填充内容等
        """
        result = self.match_filling_rule(filename, file_title)

        if result:
            return {
                'matched': True,
                'output': result,
                'message': '匹配成功'
            }
        else:
            return {
                'matched': False,
                'output': None,
                'message': '未匹配任何规则'
            }


# 向后兼容：提供与原KeywordMatcher相同的接口
class KeywordMatcher:
    """关键词匹配器（向后兼容接口）"""

    _engine = None

    @classmethod
    def _get_engine(cls):
        """获取规则引擎实例（单例）"""
        if cls._engine is None:
            cls._engine = RuleEngine()
        return cls._engine

    @staticmethod
    def match_keywords(filename, file_title=''):
        """
        根据文件名和文件标题匹配关键词，返回对应的填充信息

        Args:
            filename: 文件名
            file_title: 文件标题（从PDF/DOC中提取的）

        返回: dict包含拟办意见和办理结果，如果没有匹配返回None
        """
        engine = KeywordMatcher._get_engine()
        result = engine.match_filling_rule(filename, file_title)

        # 只返回拟办意见和办理结果，不返回文件标题
        if result and '文件标题' in result:
            # 文件标题的修正在其他地方处理
            result_copy = result.copy()
            del result_copy['文件标题']
            return result_copy

        return result

    @staticmethod
    def check_chongban(filename):
        """检查文件名中是否包含"重办"""
        return '重办' in filename


if __name__ == '__main__':
    # 测试代码
    print("=" * 60)
    print("规则引擎测试")
    print("=" * 60)

    engine = RuleEngine()

    # 测试用例
    test_cases = [
        ("热线工单_医政.doc", "关于XX医院收费问题的投诉"),
        ("热线工单_公卫.doc", "群众表扬XX社区卫生服务中心"),
        ("热线工单_医政.doc", "关于XX医院医托欺诈患者的投诉"),
        ("热线工单_医政_疾控.doc", "关于XX医院院内感染防控问题的投诉"),
        ("热线工单_疾控.doc", "关于XX酒店卫生条件差的投诉"),
        ("热线工单_中医_重办.doc", "关于XX中医诊所资质问题的投诉"),
        ("热线工单_中医院.doc", "关于市中医医院服务态度的投诉"),
    ]

    for filename, title in test_cases:
        print(f"\n测试: {filename}")
        print(f"标题: {title}")
        result = engine.match_filling_rule(filename, title)
        if result:
            print("匹配结果:")
            for key, value in result.items():
                print(f"  {key}: {value[:50]}..." if len(value) > 50 else f"  {key}: {value}")
        else:
            print("未匹配任何规则")
        print("-" * 60)
