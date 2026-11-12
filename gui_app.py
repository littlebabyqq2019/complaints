#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
投诉批办单管理系统 - GUI版本
提供可视化界面进行文件处理、编辑和打印
"""

import os
import sys
import json
import tkinter as tk
from tkinter import messagebox, filedialog, scrolledtext, simpledialog
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from datetime import datetime
import threading
from pathlib import Path
from docx import Document
from docx.shared import Pt
import platform_compat as pc
from platform_compat import (IS_WINDOWS, ui_font, get_appdata_path,
                             get_resource_path, get_user_data_path,
                             open_path, print_file, word_to_pdf)
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from PIL import Image, ImageTk
import tempfile
from pypinyin import lazy_pinyin, Style

# 导入主程序模块
from main import DocumentProcessor, WordGenerator, PDFExtractor, WordExtractor, DOCExtractor, CounterManager
from database import ComplaintDatabase

# 导入规则引擎
try:
    from rule_engine import RuleEngine
    USE_NEW_RULE_ENGINE = True
except ImportError:
    USE_NEW_RULE_ENGINE = False
    print("警告: 无法导入规则引擎模块")


class InputHistoryManager:
    """输入历史管理器 - 管理文本框的输入历史和频率统计"""
    
    def __init__(self, history_file='input_history.json'):
        self.history_file = get_appdata_path(history_file)
        self.history_data = {}  # {field_name: {value: count}}
        self.load_history()
    
    def load_history(self):
        """加载历史记录"""
        try:
            if os.path.exists(self.history_file):
                with open(self.history_file, 'r', encoding='utf-8') as f:
                    self.history_data = json.load(f)
        except Exception as e:
            print(f"加载输入历史失败: {e}")
            self.history_data = {}
    
    def save_history(self):
        """保存历史记录"""
        try:
            with open(self.history_file, 'w', encoding='utf-8') as f:
                json.dump(self.history_data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存输入历史失败: {e}")
    
    def add_value(self, field_name, value):
        """添加输入值（增加频率计数）"""
        if not value or not value.strip():
            return
        
        value = value.strip()
        if field_name not in self.history_data:
            self.history_data[field_name] = {}
        
        if value in self.history_data[field_name]:
            self.history_data[field_name][value] += 1
        else:
            self.history_data[field_name][value] = 1
        
        self.save_history()
    
    @staticmethod
    def get_pinyin_initials(text):
        """获取汉字的拼音首字母
        
        Args:
            text: 输入文本
            
        Returns:
            拼音首字母字符串（小写）
        """
        if not text:
            return ''
        
        # 获取每个字符的拼音首字母
        initials = []
        for char in text:
            if '\u4e00' <= char <= '\u9fff':  # 判断是否为汉字
                pinyin_list = lazy_pinyin(char, style=Style.FIRST_LETTER)
                if pinyin_list:
                    initials.append(pinyin_list[0].lower())
            else:
                # 非汉字字符保持原样
                initials.append(char.lower())
        
        return ''.join(initials)
    
    def get_suggestions(self, field_name, prefix='', max_count=10):
        """获取建议列表
        
        Args:
            field_name: 字段名
            prefix: 查询字符串（空字符串表示不过滤）
            max_count: 最大返回数量
            
        Returns:
            按频率排序的建议列表
        """
        if field_name not in self.history_data:
            return []
        
        # 获取所有值及其频率
        items = self.history_data[field_name].items()
        
        # 如果有查询字符串，过滤匹配的值（支持原文匹配和拼音首字母匹配）
        if prefix:
            prefix_lower = prefix.lower()
            # 获取查询字符串的拼音首字母（用于拼音匹配）
            prefix_pinyin = self.get_pinyin_initials(prefix)
            
            matched_items = []
            for value, count in items:
                value_lower = value.lower()
                # 匹配1：原文包含查询字符串
                if prefix_lower in value_lower:
                    matched_items.append((value, count))
                # 匹配2：拼音首字母包含查询字符串的拼音首字母
                elif prefix_pinyin and prefix_pinyin in self.get_pinyin_initials(value):
                    matched_items.append((value, count))
            
            items = matched_items
        else:
            items = list(items)
        
        # 按频率降序排序
        items.sort(key=lambda x: x[1], reverse=True)
        
        # 返回前max_count个值
        return [value for value, count in items[:max_count]]


class AutocompleteEntry(tk.Entry):
    """带自动完成功能的Entry控件（使用tk.Entry并实现撤销功能）"""
    
    def __init__(self, parent, field_name, history_manager, **kwargs):
        # 提取字体参数，因为tk.Entry的font参数格式可能不同
        font = kwargs.pop('font', (ui_font('仿宋_GB2312'), 11))
        width = kwargs.pop('width', 60)
        
        # 初始化tk.Entry（Entry控件不支持undo参数）
        super().__init__(parent, font=font, width=width, **kwargs)
        
        self.field_name = field_name
        self.history_manager = history_manager
        self.suggestion_text = None  # Text控件用于显示建议
        self.listbox_window = None
        self.suggestions_list = []  # 存储当前的建议列表
        self.selected_line = 0  # 当前选中的行（0-indexed）
        
        # 撤销/重做栈
        self.undo_stack = []
        self.redo_stack = []
        self.last_content = ""
        self.is_undoing = False
        
        # 记录初始状态
        self._save_state()
        
        # 绑定事件
        self.bind('<FocusIn>', self.on_focus_in)
        self.bind('<KeyRelease>', self.on_key_release_with_undo)
        self.bind('<Down>', self.on_down)
        self.bind('<Up>', self.on_up)
        self.bind('<Return>', self.on_return)
        self.bind('<Escape>', self.hide_listbox)
        self.bind('<FocusOut>', self.on_focus_out)
    
    def _save_state(self):
        """保存当前状态到撤销栈"""
        if self.is_undoing:
            return
        
        current = self.get()
        if current != self.last_content:
            self.undo_stack.append(self.last_content)
            self.last_content = current
            self.redo_stack.clear()  # 新输入会清空重做栈
            
            # 限制撤销栈大小，避免内存过大
            if len(self.undo_stack) > 100:
                self.undo_stack.pop(0)
    
    def edit_undo(self):
        """撤销操作"""
        if self.undo_stack:
            self.is_undoing = True
            self.redo_stack.append(self.last_content)
            content = self.undo_stack.pop()
            self.delete(0, tk.END)
            self.insert(0, content)
            self.last_content = content
            self.is_undoing = False
    
    def edit_redo(self):
        """重做操作"""
        if self.redo_stack:
            self.is_undoing = True
            self.undo_stack.append(self.last_content)
            content = self.redo_stack.pop()
            self.delete(0, tk.END)
            self.insert(0, content)
            self.last_content = content
            self.is_undoing = False
    
    def on_key_release_with_undo(self, event):
        """键盘释放时更新建议列表并保存状态"""
        # 保存撤销状态
        self._save_state()
        
        # 调用原有的自动完成逻辑
        self.on_key_release(event)
    
    def on_focus_in(self, event):
        """获得焦点时显示历史记录"""
        self.show_suggestions()
    
    def on_key_release(self, event):
        """键盘释放时更新建议列表"""
        # 忽略特殊键
        if event.keysym in ('Down', 'Up', 'Return', 'Escape', 'Left', 'Right', 
                            'Home', 'End', 'Prior', 'Next', 'Tab'):
            return
        
        self.show_suggestions()
    
    def show_suggestions(self):
        """显示建议列表"""
        current_value = self.get()
        suggestions = self.history_manager.get_suggestions(self.field_name, current_value, 10)
        
        if not suggestions:
            self.hide_listbox()
            return
        
        # 创建或更新suggestion_text
        if not self.suggestion_text:
            self.create_listbox()
        
        # 保存建议列表
        self.suggestions_list = suggestions
        self.selected_line = 0  # 重置选中行
        
        # 更新Text内容，并高亮匹配的字符
        self.suggestion_text.config(state='normal')
        self.suggestion_text.delete('1.0', tk.END)
        
        for i, suggestion in enumerate(suggestions):
            line_start = f"{i+1}.0"
            self.suggestion_text.insert(tk.END, suggestion + '\n')
            
            # 查找并高亮匹配的字符
            if current_value:
                self._highlight_matches(suggestion, current_value, i+1)
        
        # 高亮选中的行
        self._update_selection()
        
        self.suggestion_text.config(state='disabled')
        
        # 显示listbox
        if self.listbox_window:
            self.listbox_window.deiconify()
            self.position_listbox()
    
    def create_listbox(self):
        """创建下拉列表框（使用Text控件）"""
        # 创建顶层窗口
        self.listbox_window = tk.Toplevel(self.winfo_toplevel())
        self.listbox_window.wm_overrideredirect(True)
        self.listbox_window.withdraw()  # 初始隐藏
        
        # 创建Text控件（替代Listbox）
        self.suggestion_text = tk.Text(self.listbox_window, 
                                       font=(ui_font('仿宋_GB2312'), 11),
                                       height=10,
                                       wrap=tk.NONE,
                                       cursor='arrow')
        self.suggestion_text.pack(fill=tk.BOTH, expand=True)
        
        # 配置标签样式
        self.suggestion_text.tag_config('selected', background='#0078d7', foreground='white')
        self.suggestion_text.tag_config('highlight', foreground='red', font=(ui_font('仿宋_GB2312'), 11, 'bold'))
        # 提高highlight标签的优先级，确保红色文字在选中行也能显示
        self.suggestion_text.tag_raise('highlight')
        
        # 绑定选择事件
        self.suggestion_text.bind('<Button-1>', self.on_listbox_click)
        self.suggestion_text.bind('<Return>', self.on_listbox_select)
    
    def _highlight_matches(self, text, query, line_number):
        """在Text控件中高亮匹配的字符（支持原文匹配和拼音匹配）"""
        query_lower = query.lower()
        text_lower = text.lower()
        
        # 模式1：原文匹配 - 查找文本中直接包含查询字符串的部分
        start_pos = 0
        while True:
            pos = text_lower.find(query_lower, start_pos)
            if pos == -1:
                break
            
            # 计算匹配位置的索引
            start_index = f"{line_number}.{pos}"
            end_index = f"{line_number}.{pos + len(query)}"
            
            # 添加高亮标签
            self.suggestion_text.tag_add('highlight', start_index, end_index)
            
            start_pos = pos + 1
        
        # 模式2：拼音首字母匹配 - 查找拼音首字母匹配的汉字序列
        # 只有当查询字符串不在原文中时，才进行拼音匹配（避免重复标红）
        if query_lower not in text_lower:
            self._highlight_pinyin_matches(text, query_lower, line_number)
    
    def _highlight_pinyin_matches(self, text, query, line_number):
        """高亮拼音首字母匹配的汉字序列"""
        if not query:
            return
        
        query_len = len(query)
        text_len = len(text)
        
        # 获取整个文本的拼音首字母
        text_pinyin = self.history_manager.get_pinyin_initials(text)
        
        # 在拼音首字母中查找查询字符串
        pinyin_pos = text_pinyin.find(query)
        if pinyin_pos == -1:
            return  # 未找到匹配
        
        # 找到对应的汉字位置
        # 因为每个汉字对应一个拼音首字母（非汉字字符也对应一个字符）
        # 所以拼音位置 = 汉字位置
        char_start = pinyin_pos
        char_end = pinyin_pos + query_len
        
        # 确保位置有效
        if char_end <= text_len:
            # 标红匹配的汉字
            start_index = f"{line_number}.{char_start}"
            end_index = f"{line_number}.{char_end}"
            self.suggestion_text.tag_add('highlight', start_index, end_index)
    
    def _update_selection(self):
        """更新选中行的高亮"""
        if not self.suggestion_text:
            return
        
        # 清除所有选中标签
        self.suggestion_text.tag_remove('selected', '1.0', tk.END)
        
        # 高亮当前选中的行
        if 0 <= self.selected_line < len(self.suggestions_list):
            line_num = self.selected_line + 1
            self.suggestion_text.tag_add('selected', f"{line_num}.0", f"{line_num}.end")
            # 提高highlight标签优先级，确保红色文字显示在蓝色背景之上
            self.suggestion_text.tag_raise('highlight')
            # 确保选中的行可见
            self.suggestion_text.see(f"{line_num}.0")
    
    def position_listbox(self):
        """定位listbox到Entry下方"""
        if not self.listbox_window:
            return
        
        # 获取Entry的位置和尺寸
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height()
        width = self.winfo_width()
        
        # 设置listbox位置和宽度
        self.listbox_window.geometry(f"{width}x200+{x}+{y}")
    
    def hide_listbox(self, event=None):
        """隐藏建议列表"""
        if self.listbox_window:
            self.listbox_window.withdraw()
    
    def on_down(self, event):
        """向下键 - 选择下一个"""
        if self.suggestion_text and self.listbox_window and self.listbox_window.winfo_viewable():
            if self.selected_line < len(self.suggestions_list) - 1:
                self.selected_line += 1
                self.suggestion_text.config(state='normal')
                self._update_selection()
                self.suggestion_text.config(state='disabled')
            return 'break'  # 阻止默认行为
    
    def on_up(self, event):
        """向上键 - 选择上一个"""
        if self.suggestion_text and self.listbox_window and self.listbox_window.winfo_viewable():
            if self.selected_line > 0:
                self.selected_line -= 1
                self.suggestion_text.config(state='normal')
                self._update_selection()
                self.suggestion_text.config(state='disabled')
            return 'break'
    
    def on_return(self, event):
        """回车键 - 选择当前项"""
        if self.suggestion_text and self.listbox_window and self.listbox_window.winfo_viewable():
            self.on_listbox_select(event)
            return 'break'
    
    def on_listbox_click(self, event):
        """鼠标点击listbox"""
        # 获取点击的位置
        index = self.suggestion_text.index(f"@{event.x},{event.y}")
        line_num = int(index.split('.')[0]) - 1  # 转换为0-indexed
        
        if 0 <= line_num < len(self.suggestions_list):
            self.selected_line = line_num
            # 延迟一点时间，确保更新后再选择
            self.after(50, lambda: self.on_listbox_select(event))
    
    def on_listbox_select(self, event):
        """选择listbox中的项"""
        if not self.suggestion_text or not self.suggestions_list:
            return
        
        if 0 <= self.selected_line < len(self.suggestions_list):
            value = self.suggestions_list[self.selected_line]
            self.delete(0, tk.END)
            self.insert(0, value)
            self.hide_listbox()
            self.icursor(tk.END)  # 光标移到末尾
    
    def on_focus_out(self, event):
        """失去焦点时隐藏列表"""
        # 延迟隐藏，以便点击listbox时能正常选择
        self.after(200, self.hide_listbox)


class ComplaintManagementGUI:
    """投诉批办单管理系统GUI"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("投诉批办单管理系统 v2.0.0")
        
        # 设置窗口图标：Windows 用 .ico，统信/麒麟用同内容的 PNG（Tk 跨平台一致）
        try:
            if IS_WINDOWS:
                _ico = get_resource_path('1.ico')
                if os.path.isfile(_ico):
                    self.root.iconbitmap(_ico)
            else:
                _png = get_resource_path('1.png')
                if os.path.isfile(_png):
                    _img = tk.PhotoImage(file=_png)
                    self.root.iconphoto(True, _img)
                    self._icon_image = _img  # 保持引用，避免被 GC
        except Exception:
            pass
        
        # 初始化变量
        self.processor = DocumentProcessor()
        self.current_file = None
        self.original_receipt_number = None  # 追踪原始收文编号，用于更新数据库
        self.file_list = []  # 待打印文件列表
        self.file_checkboxes = {}  # 待打印文件复选框状态 {item_id: BooleanVar}
        self.source_files = []  # 待处理文件列表
        self.main_paned = None  # 主分割窗口
        self.is_restarting_for_theme = False  # 标记是否是为了切换主题而重启
        self.input_history = InputHistoryManager()  # 输入历史管理器
        self.is_processing = False
        self.is_printing = False
        
        # 初始化数据库（自动选择上次使用的数据库）
        self.db = self.init_database()
        
        # 合并撤销功能相关变量
        self.merge_backup = None  # 保存合并前的状态
        self.undo_merge_button = None  # 撤销合并按钮的引用
        
        # 加载路径配置
        self.path_config = self.load_path_config()
        
        # 创建必要的目录
        self.ensure_directories()
        
        # 加载快捷按钮配置
        self.quick_buttons_config = self.load_quick_buttons_config()
        
        # 加载窗口布局配置
        self.window_config = self.load_window_config()
        
        # 加载二维码配置
        self.qrcode_config = self.load_qrcode_config()
        
        # 设置窗口大小和位置
        self.apply_window_config()
        
        # 创建界面
        self.create_widgets()
        self.refresh_file_list()
        self.refresh_source_list()
        
        # 显示下一个收文编号
        self.update_next_receipt_number()
        
        # 恢复分割窗口位置（需要在界面创建后）
        # 增加延迟时间，并多次尝试确保设置成功
        self.root.after(100, self.restore_paned_positions)
        self.root.after(300, self.restore_paned_positions)
        self.root.after(600, self.restore_paned_positions)
        self.root.after(1000, self.restore_paned_positions)
        self.root.after(1500, self.restore_paned_positions)
        
        # 绑定窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # 清理上次运行产生的临时PDF文件
        self.cleanup_temp_pdf_files()
        
    def cleanup_temp_pdf_files(self):
        """清理上次运行产生的临时PDF文件"""
        try:
            import glob
            temp_dir = self.path_config['待打印']
            
            # 查找两种临时PDF文件
            patterns = [
                os.path.join(temp_dir, '批量打印合并_*.pdf'),  # 合并的PDF
                os.path.join(temp_dir, 'temp_*.pdf')  # 转换过程的临时PDF
            ]
            
            all_temp_files = []
            for pattern in patterns:
                all_temp_files.extend(glob.glob(pattern))
            
            if all_temp_files:
                self.log(f"发现 {len(all_temp_files)} 个临时PDF文件，正在清理...")
                for pdf_file in all_temp_files:
                    try:
                        os.remove(pdf_file)
                        self.log(f"已清理临时文件: {os.path.basename(pdf_file)}")
                    except Exception as e:
                        self.log(f"清理临时文件失败 {os.path.basename(pdf_file)}: {e}")
                self.log("临时PDF文件清理完成")
            else:
                self.log("未发现需要清理的临时PDF文件")
                
        except Exception as e:
            self.log(f"清理临时PDF文件失败: {e}")
    
    def bind_text_shortcuts(self, widget):
        """为文本框绑定快捷键
        
        Args:
            widget: tk.Text 或 tk.Entry 控件
        """
        # 判断控件类型
        is_text = isinstance(widget, tk.Text)
        is_entry = isinstance(widget, (ttk.Entry, tk.Entry))
        
        if is_text:
            # Text 控件的快捷键绑定
            # Ctrl+Z: 撤销
            def undo_text(event):
                try:
                    widget.edit_undo()
                except tk.TclError:
                    pass  # 没有可撤销的内容
                return 'break'
            widget.bind('<Control-z>', undo_text)
            widget.bind('<Control-Z>', undo_text)
            
            # Ctrl+Y: 重做
            def redo_text(event):
                try:
                    widget.edit_redo()
                except tk.TclError:
                    pass  # 没有可重做的内容
                return 'break'
            widget.bind('<Control-y>', redo_text)
            widget.bind('<Control-Y>', redo_text)
            
            # Ctrl+A: 全选
            def select_all_text(event):
                widget.tag_add(tk.SEL, "1.0", tk.END)
                widget.mark_set(tk.INSERT, "1.0")
                widget.see(tk.INSERT)
                return 'break'  # 阻止默认行为
            widget.bind('<Control-a>', select_all_text)
            widget.bind('<Control-A>', select_all_text)
            
        elif is_entry:
            # Entry 控件的快捷键绑定
            # Ctrl+Z: 撤销
            def undo_entry(event):
                try:
                    event.widget.edit_undo()
                except tk.TclError:
                    pass  # 没有可撤销的内容
                return 'break'
            widget.bind('<Control-z>', undo_entry)
            widget.bind('<Control-Z>', undo_entry)
            
            # Ctrl+Y: 重做
            def redo_entry(event):
                try:
                    event.widget.edit_redo()
                except tk.TclError:
                    pass  # 没有可重做的内容
                return 'break'
            widget.bind('<Control-y>', redo_entry)
            widget.bind('<Control-Y>', redo_entry)
            
            # Ctrl+A: 全选
            def select_all_entry(event):
                event.widget.select_range(0, tk.END)
                event.widget.icursor(tk.END)
                return 'break'  # 阻止默认行为
            widget.bind('<Control-a>', select_all_entry)
            widget.bind('<Control-A>', select_all_entry)
        
        # 通用快捷键（Text 和 Entry 都支持，但由系统默认处理）
        # Ctrl+C: 复制 (系统默认)
        # Ctrl+X: 剪切 (系统默认)
        # Ctrl+V: 粘贴 (系统默认)
    
    def init_database(self):
        """
        初始化数据库：
        1. 优先使用上次记录的数据库路径
        2. 如果没有记录，检查AppData中是否有数据库文件
        3. 如果AppData中没有任何数据库，才创建默认的complaints.db
        """
        from database import get_appdata_path as db_appdata_path
        import glob
        
        last_db_config = get_appdata_path('last_database.json')
        
        # 1. 尝试加载上次使用的数据库
        if os.path.exists(last_db_config):
            try:
                with open(last_db_config, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    last_db = data.get('last_database', '')
                # 兼容 Windows 版遗留配置里的反斜杠路径（UOS/麒麟不做路径分隔符解析）
                if last_db and not os.path.exists(last_db) and '\\' in last_db:
                    last_db = last_db.replace('\\', os.sep)
                if last_db and os.path.exists(last_db):
                    print(f"使用上次的数据库: {last_db}")
                    return ComplaintDatabase(last_db)
            except Exception as e:
                print(f"加载上次数据库配置失败: {e}")
        
        # 2. 扫描AppData目录中的数据库文件（优先使用年份格式）
        appdata_dir = db_appdata_path()
        # 查找年份格式的数据库（如 complaints-2026.db）
        year_dbs = sorted(glob.glob(os.path.join(appdata_dir, 'complaints-????.db')), reverse=True)
        if year_dbs:
            # 使用最新年份的数据库
            latest_db = year_dbs[0]
            print(f"发现年份数据库，使用: {latest_db}")
            self.save_last_database(latest_db)
            return ComplaintDatabase(latest_db)
        
        # 3. 查找默认数据库（complaints.db）
        default_db = db_appdata_path('complaints.db')
        if os.path.exists(default_db):
            print(f"使用默认数据库: {default_db}")
            self.save_last_database(default_db)
            return ComplaintDatabase(default_db)
        
        # 4. AppData中完全没有数据库，创建默认的complaints.db
        print("未找到任何数据库，创建默认数据库: complaints.db")
        return ComplaintDatabase()  # 创建默认 complaints.db
    
    def save_last_database(self, db_path):
        """保存当前使用的数据库路径到配置文件"""
        try:
            config_file = get_appdata_path('last_database.json')
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump({'last_database': db_path}, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存数据库路径失败: {e}")
    
    def load_path_config(self):
        """加载路径配置"""
        config_file = get_appdata_path('path_config.json')
        default_config = {
            '已处理': '已处理',
            '待打印': '待打印',
            '已打印': '已打印',
            '默认待处理文件夹': ''
        }
        
        try:
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception as e:
            print(f"加载路径配置失败: {e}")
        
        # 保存默认配置
        try:
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(default_config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存默认路径配置失败: {e}")
        
        return default_config
    
    def ensure_directories(self):
        """确保必要的目录存在"""
        for key, dir_name in self.path_config.items():
            if key == '默认待处理文件夹':
                continue  # 跳过默认待处理文件夹，这个可以为空
            
            if not dir_name:
                continue  # 跳过空路径
            
            try:
                # 尝试创建目录
                os.makedirs(dir_name, exist_ok=True)
            except Exception as e:
                print(f"创建目录 {dir_name} 失败: {e}")
                # 如果是绝对路径且失败，尝试使用相对路径
                try:
                    # 提取目录名（最后一部分）
                    base_name = os.path.basename(dir_name.rstrip('/\\'))
                    if base_name:
                        os.makedirs(base_name, exist_ok=True)
                        # 更新配置为相对路径
                        self.path_config[key] = base_name
                        print(f"改用相对路径: {base_name}")
                except Exception as e2:
                    print(f"创建相对路径目录也失败: {e2}")
    
    def create_widgets(self):
        """创建界面组件"""
        # 创建菜单栏
        self.create_menu()
        
        # 创建主框架（左中右三栏）
        self.main_paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.main_paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 左侧面板：待处理文件列表
        left_frame = ttk.Frame(self.main_paned, width=240)
        self.main_paned.add(left_frame, weight=1)
        self.create_source_panel(left_frame)
        
        # 中间面板：待打印文件列表
        middle_frame = ttk.Frame(self.main_paned, width=240)
        self.main_paned.add(middle_frame, weight=1)
        self.create_output_panel(middle_frame)
        
        # 右侧面板：编辑区域
        right_frame = ttk.Frame(self.main_paned)
        self.main_paned.add(right_frame, weight=3)
        self.create_right_panel(right_frame)
        
        # 底部状态栏
        self.create_status_bar()
        
    def create_menu(self):
        """创建菜单栏"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # 文件菜单
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="文件", menu=file_menu)
        file_menu.add_command(label="添加待处理文件", command=self.add_source_files)
        file_menu.add_command(label="添加待处理文件夹", command=self.add_source_folder)
        file_menu.add_separator()
        file_menu.add_command(label="批量处理源文件", command=self.batch_process)
        file_menu.add_separator()
        file_menu.add_command(label="打开待打印目录", command=self.open_output_dir)
        file_menu.add_command(label="打开已打印目录", command=self.open_printed_dir)
        file_menu.add_command(label="打开已处理目录", command=self.open_processed_dir)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.root.quit)
        
        # 编辑菜单
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="编辑", menu=edit_menu)
        edit_menu.add_command(label="保存当前文件", command=self.save_current_file)
        edit_menu.add_command(label="刷新文件列表", command=self.refresh_all_lists)
        
        # 打印菜单
        print_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="打印", menu=print_menu)
        print_menu.add_command(label="打印预览", command=self.print_preview)
        print_menu.add_separator()
        print_menu.add_command(label="打印当前文件", command=self.print_current)
        print_menu.add_separator()
        print_menu.add_command(label="批量转PDF选中", command=self.batch_convert_to_pdf)
        print_menu.add_command(label="批量打印选中（静默）", command=self.silent_batch_print)
        print_menu.add_separator()
        print_menu.add_command(label="批量打印全部", command=self.batch_print_all)
        
        # 工具菜单
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="工具", menu=tools_menu)
        tools_menu.add_command(label="路径设置", command=self.manage_paths)
        tools_menu.add_command(label="设置默认待处理文件夹", command=self.set_default_source_folder)
        tools_menu.add_separator()
        tools_menu.add_command(label="主题切换", command=self.change_theme)
        tools_menu.add_command(label="重置窗口布局", command=self.reset_window_layout)
        tools_menu.add_separator()
        tools_menu.add_command(label="编号管理", command=self.manage_counter)
        tools_menu.add_command(label="二维码设置", command=self.manage_qrcode)
        tools_menu.add_separator()
        
        # 数据库管理子菜单
        db_menu = tk.Menu(tools_menu, tearoff=0)
        tools_menu.add_cascade(label="数据库管理", menu=db_menu)
        db_menu.add_command(label="📊 数据库统计信息", command=self.show_database_stats)
        db_menu.add_separator()
        db_menu.add_command(label="💾 数据库备份", command=self.backup_database)
        db_menu.add_command(label="📥 数据库恢复", command=self.restore_database)
        db_menu.add_separator()
        db_menu.add_command(label="📋 导出为Excel", command=self.export_database_to_excel)
        db_menu.add_separator()
        db_menu.add_command(label="🔧 数据库优化", command=self.optimize_database)
        db_menu.add_command(label="🔍 数据库完整性检查", command=self.check_database_integrity)
        db_menu.add_separator()
        db_menu.add_command(label="⚙️ 自动备份设置", command=self.auto_backup_settings)
        db_menu.add_separator()
        db_menu.add_command(label="🗑️ 数据库初始化（清空）", command=self.initialize_database)
        
        # 帮助菜单
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="帮助", menu=help_menu)
        help_menu.add_command(label="使用说明", command=self.show_help)
        help_menu.add_command(label="关于", command=self.show_about)
        
    def create_source_panel(self, parent):
        """创建待处理文件列表面板"""
        # 标题框架 - 添加背景色和圆角效果
        title_frame = ttk.Frame(parent, bootstyle="primary")
        title_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 标题行 - 使用Frame容器包含标题和搜索框，保持蓝色背景
        title_row = ttk.Frame(title_frame, bootstyle="primary")
        title_row.pack(fill=tk.X, padx=10, pady=8)
        
        # 标题（左侧）
        title = ttk.Label(title_row, text="📋 待处理文件", font=(ui_font('微软雅黑'), 12, 'bold'), 
                         bootstyle="inverse-primary")
        title.pack(side=tk.LEFT)
        
        # 搜索框容器（右侧） - 使用白色背景的Frame模拟输入框外观
        search_container = ttk.Frame(title_row, bootstyle="light")
        search_container.pack(side=tk.RIGHT, padx=(10, 0))
        
        # 搜索图标（在白色容器内左侧）- 使用深色让图标可见
        search_icon = ttk.Label(search_container, text="🔍", foreground="#666666", background="white")
        search_icon.pack(side=tk.LEFT, padx=(5, 2))
        
        # 搜索输入框（在白色容器内，无边框样式）
        self.source_search_var = tk.StringVar()
        self.source_search_var.trace('w', self.highlight_source_search)
        search_entry = ttk.Entry(search_container, textvariable=self.source_search_var, 
                                width=25, bootstyle="light")
        search_entry.pack(side=tk.LEFT, padx=(0, 5))
        
        # 文件列表
        list_frame = ttk.Frame(parent)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 创建Treeview - 添加序号列（删除类型列）
        self.source_tree = ttk.Treeview(list_frame, columns=('序号', '文件名'), show='headings', selectmode='extended', height=15)
        self.source_tree.heading('序号', text='序号')
        self.source_tree.heading('文件名', text='文件名')
        self.source_tree.column('序号', width=40, anchor='center')
        self.source_tree.column('文件名', width=210)
        
        # 配置tag颜色用于高亮搜索结果
        self.source_tree.tag_configure('highlight', foreground='red', font=(ui_font('微软雅黑'), 9, 'bold'))
        
        # 滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.source_tree.yview)
        self.source_tree.configure(yscrollcommand=scrollbar.set)
        
        self.source_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 操作按钮
        # 按钮区域 - 三列排列，增大按钮宽度，使用现代化按钮样式
        btn_container = ttk.Frame(parent)
        btn_container.pack(fill=tk.X, padx=10, pady=10)
        
        # 第一行 - 统一按钮宽度
        btn_row1 = ttk.Frame(btn_container)
        btn_row1.pack(fill=tk.X, pady=2)
        ttk.Button(btn_row1, text="📁添加文件", command=self.add_source_files, width=12, bootstyle="info").pack(side=tk.LEFT, padx=2, ipady=2)
        ttk.Button(btn_row1, text="📂添加文件夹", command=self.add_source_folder, width=12, bootstyle="info").pack(side=tk.LEFT, padx=2, ipady=2)
        ttk.Button(btn_row1, text="⚙处理文件", command=self.batch_process, width=12, bootstyle="success").pack(side=tk.LEFT, padx=2, ipady=2)
        
        # 第二行 - 统一按钮宽度
        btn_row2 = ttk.Frame(btn_container)
        btn_row2.pack(fill=tk.X, pady=2)
        ttk.Button(btn_row2, text="➖移除选中", command=self.remove_source_selected, width=12, bootstyle="warning").pack(side=tk.LEFT, padx=2, ipady=2)
        ttk.Button(btn_row2, text="🗑清空列表", command=self.clear_source_list, width=12, bootstyle="danger").pack(side=tk.LEFT, padx=2, ipady=2)
        ttk.Button(btn_row2, text="🔄刷新列表", command=self.refresh_default_source_folder, width=12, bootstyle="primary").pack(side=tk.LEFT, padx=2, ipady=2)
        
        # 第三行 - 重命名按钮
        btn_row3 = ttk.Frame(btn_container)
        btn_row3.pack(fill=tk.X, pady=2)
        ttk.Button(btn_row3, text="📝重命名docx", command=self.rename_docx_to_doc, width=12, bootstyle="secondary").pack(side=tk.LEFT, padx=2, ipady=2)
    
    def create_output_panel(self, parent):
        """创建待打印文件列表面板"""
        # 标题框架 - 添加背景色
        title_frame = ttk.Frame(parent, bootstyle="success")
        title_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 标题
        title = ttk.Label(title_frame, text="🖨️ 待打印批办单", font=(ui_font('微软雅黑'), 12, 'bold'),
                         bootstyle="inverse-success")
        title.pack(fill=tk.X, padx=10, pady=8)
        
        # 文件列表
        list_frame = ttk.Frame(parent)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 创建Treeview - 添加复选框列、序号列和文件名列
        self.file_tree = ttk.Treeview(list_frame, columns=('选择', '序号', '文件名'), show='headings', selectmode='extended', height=15)
        self.file_tree.heading('选择', text='☐', command=self.toggle_all_checkboxes)  # 点击表头全选/全不选
        self.file_tree.heading('序号', text='序号')
        self.file_tree.heading('文件名', text='文件名')
        self.file_tree.column('选择', width=35, anchor='center')
        self.file_tree.column('序号', width=40, anchor='center')
        self.file_tree.column('文件名', width=175)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=scrollbar.set)
        
        self.file_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 单击或键盘选择时加载文件
        self.file_tree.bind('<<TreeviewSelect>>', lambda e: self.load_file_auto())
        # 双击编辑文件
        self.file_tree.bind('<Double-1>', lambda e: self.load_file())
        # 点击复选框列切换选择状态
        self.file_tree.bind('<Button-1>', self.on_tree_click)
        
        # 按钮区域 - 3列布局，固定宽度
        btn_container = ttk.Frame(parent)
        btn_container.pack(fill=tk.X, padx=10, pady=10)
        
        # 配置列权重和最小宽度，使3列宽度一致
        # 使用uniform参数确保所有列宽度相同
        btn_container.columnconfigure(0, weight=1, uniform='button_col')
        btn_container.columnconfigure(1, weight=1, uniform='button_col')
        btn_container.columnconfigure(2, weight=1, uniform='button_col')
        
        # 第一行 - 3个按钮，设置固定宽度
        ttk.Button(btn_container, text="📄批量转pdf", command=self.batch_convert_to_pdf, width=10, bootstyle="info").grid(row=0, column=0, padx=2, pady=2, sticky='ew', ipady=2)
        ttk.Button(btn_container, text="🖨批量打印", command=self.silent_batch_print, width=10, bootstyle="success").grid(row=0, column=1, padx=2, pady=2, sticky='ew', ipady=2)
        ttk.Button(btn_container, text="📦移至已打印", command=self.move_selected_to_printed, width=10, bootstyle="warning").grid(row=0, column=2, padx=2, pady=2, sticky='ew', ipady=2)
        
        # 第二行 - 3个按钮，设置固定宽度
        ttk.Button(btn_container, text="🗑删除选中", command=self.delete_selected, width=10, bootstyle="danger").grid(row=1, column=0, padx=2, pady=2, sticky='ew', ipady=2)
        ttk.Button(btn_container, text="🔄刷新列表", command=self.refresh_file_list, width=10, bootstyle="primary").grid(row=1, column=1, padx=2, pady=2, sticky='ew', ipady=2)
        ttk.Button(btn_container, text="🔗合并批办单", command=self.merge_selected_documents, width=10, bootstyle="secondary").grid(row=1, column=2, padx=2, pady=2, sticky='ew', ipady=2)
        
        # 第三行 - 撤销合并按钮（与合并批办单按钮同列，初始禁用）
        self.undo_merge_button = ttk.Button(btn_container, text="↩️撤销合并", command=self.undo_merge, width=10, bootstyle="warning", state='disabled')
        self.undo_merge_button.grid(row=2, column=2, padx=2, pady=2, sticky='ew', ipady=2)
        
        
    def create_right_panel(self, parent):
        """创建右侧编辑面板"""
        # 标题栏 - 使用带颜色的框架
        title_frame = ttk.Frame(parent, bootstyle="secondary")
        title_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 左侧：文件标题
        self.file_title_label = ttk.Label(title_frame, text="📝 请选择文件进行编辑", 
                                          font=(ui_font('微软雅黑'), 12, 'bold'),
                                          bootstyle="inverse-secondary")
        self.file_title_label.pack(side=tk.LEFT, padx=10, pady=8)
        
        # 右侧：收文编号显示
        self.receipt_number_label = ttk.Label(title_frame, text="", 
                                             font=(ui_font('微软雅黑'), 12, 'bold'),
                                             bootstyle="inverse-info")
        self.receipt_number_label.pack(side=tk.RIGHT, padx=10, pady=8)
        
        # 创建Notebook（标签页）
        self.notebook = ttk.Notebook(parent)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 表格编辑页
        self.edit_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.edit_tab, text="批办单编辑")
        self.create_edit_tab(self.edit_tab)
        
        # 处理日志页
        log_tab = ttk.Frame(self.notebook)
        self.notebook.add(log_tab, text="处理日志")
        self.create_log_tab(log_tab)
        
        # 数据库检索页
        db_tab = ttk.Frame(self.notebook)
        self.notebook.add(db_tab, text="数据库检索")
        self.create_database_tab(db_tab)

        # 切换到数据库检索页时自动刷新
        self.notebook.bind('<<NotebookTabChanged>>', self._on_tab_changed)
        
    def create_edit_tab(self, parent):
        """创建表格编辑标签页"""
        # 创建主容器，分左右两部分
        main_container = ttk.Frame(parent)
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # 左侧：表单编辑区域
        left_frame = ttk.Frame(main_container)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 创建滚动框架
        canvas = tk.Canvas(left_frame)
        scrollbar = ttk.Scrollbar(left_frame, orient="vertical", command=canvas.yview)
        self.edit_frame = ttk.Frame(canvas)
        
        self.edit_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.edit_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 右侧：快捷按钮区域
        self.right_frame = ttk.Frame(main_container, bootstyle="light")
        self.right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(10, 5))
        
        # 快捷按钮标题框架
        title_container = ttk.Frame(self.right_frame, bootstyle="info")
        title_container.pack(fill=tk.X, pady=(0, 10))
        
        # 快捷按钮标题
        title_label = ttk.Label(title_container, text="⚡ 快捷填充", 
                               font=(ui_font('微软雅黑'), 12, 'bold'),
                               bootstyle="inverse-info")
        title_label.pack(pady=10)
        
        # 动态快捷按钮容器
        self.quick_buttons_frame = ttk.Frame(self.right_frame)
        self.quick_buttons_frame.pack(fill=tk.X)
        
        # 动态创建快捷按钮
        self.rebuild_quick_buttons()
        
        # 添加分隔线
        ttk.Separator(self.right_frame, orient='horizontal', bootstyle="secondary").pack(fill=tk.X, padx=10, pady=8)
        
        # 设置按钮
        ttk.Button(self.right_frame, text="⚙设置预设内容",
                  command=self.open_quick_fill_settings,
                  width=12,
                  bootstyle="secondary-outline").pack(pady=4, padx=10, ipady=4)

        # 拟办规则设置按钮
        if USE_NEW_RULE_ENGINE:
            ttk.Button(self.right_frame, text="📋拟办规则设置",
                      command=self.open_rule_management,
                      width=12,
                      bootstyle="info-outline").pack(pady=4, padx=10, ipady=4)
        
        # 创建表格编辑字段
        self.field_entries = {}
        fields = [
            ('来文单位', '来文单位'),
            ('来文字号', '来文字号'),
            ('收文日期', '收文日期'),
            ('来文类型', '来文类型'),
            ('收文途径', '收文途径'),
            ('收文编号', '收文编号'),
            ('文件标题', '文件标题'),
            ('领导批示', '领导批示'),
            ('承办意见', '拟办意见'),
            ('办理情况', '办理情况'),
        ]
        
        for idx, (field_key, field_name) in enumerate(fields):
            frame = ttk.Frame(self.edit_frame)
            frame.pack(fill=tk.X, padx=10, pady=5)
            
            # 增大标签字体
            label = ttk.Label(frame, text=field_name + ":", width=12, anchor='e', 
                            font=(ui_font('微软雅黑'), 11, 'bold'))
            label.pack(side=tk.LEFT, padx=5)
            
            if field_key in ['领导批示', '承办意见']:
                # 多行文本框 - 适中高度（降低以适配1600*900分辨率）
                if field_key == '领导批示':
                    text_widget = tk.Text(frame, height=4, width=60, font=(ui_font('仿宋_GB2312'), 11), 
                                         undo=True, maxundo=-1)  # 启用撤销功能，无限撤销次数
                else:  # 承办意见（拟办意见）
                    text_widget = tk.Text(frame, height=8, width=60, font=(ui_font('仿宋_GB2312'), 11),
                                         undo=True, maxundo=-1)  # 启用撤销功能，无限撤销次数
                text_widget.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
                self.bind_text_shortcuts(text_widget)  # 绑定快捷键
                self.field_entries[field_key] = text_widget
            elif field_key == '办理情况':
                # 办理情况 - 适中高度
                text_widget = tk.Text(frame, height=4, width=60, font=(ui_font('仿宋_GB2312'), 11),
                                     undo=True, maxundo=-1)  # 启用撤销功能，无限撤销次数
                text_widget.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
                self.bind_text_shortcuts(text_widget)  # 绑定快捷键
                self.field_entries[field_key] = text_widget
            else:
                # 单行输入框（使用Entry控件）
                # 为"来文单位"、"收文途径"、"文件标题"添加自动完成功能
                if field_key in ['来文单位', '收文途径', '文件标题']:
                    entry = AutocompleteEntry(frame, field_key, self.input_history, 
                                            width=60, font=(ui_font('仿宋_GB2312'), 11))
                else:
                    entry = ttk.Entry(frame, width=60, font=(ui_font('仿宋_GB2312'), 11))
                
                # 统一单行输入框的高度（使用ipady增加垂直内边距）
                entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, ipady=3)
                
                # 为文件标题、来文单位、收文途径添加快捷键支持
                if field_key in ['文件标题', '来文单位', '收文途径']:
                    self.bind_text_shortcuts(entry)
                
                self.field_entries[field_key] = entry
                
                # 收文编号设置为只读状态（系统自动生成，不可编辑）
                if field_key == '收文编号':
                    entry.configure(state='readonly')
                
                # 特殊字段添加快捷按钮
                if field_key == '收文日期':
                    ttk.Button(frame, text="📅今天", width=7,
                              command=lambda: self.fill_today(),
                              bootstyle="info-outline").pack(side=tk.LEFT, padx=2)
                
                # 来文字号添加查询按钮
                if field_key == '来文字号':
                    ttk.Button(frame, text="🔍查询", width=7,
                              command=lambda: self.search_by_reference_number(),
                              bootstyle="success-outline").pack(side=tk.LEFT, padx=2)
        
        # 底部按钮区域
        bottom_frame = ttk.Frame(self.edit_frame)
        bottom_frame.pack(fill=tk.X, padx=10, pady=15)
        
        # 左侧：预览、打印、保存、清空按钮 - 使用现代化样式
        left_buttons = ttk.Frame(bottom_frame)
        left_buttons.pack(side=tk.LEFT)
        
        ttk.Button(left_buttons, text="👁预览", command=self.print_preview, width=7, bootstyle="info").pack(side=tk.LEFT, padx=2, ipady=5)
        ttk.Button(left_buttons, text="🖨打印", command=self.print_current, width=7, bootstyle="success").pack(side=tk.LEFT, padx=2, ipady=5)
        ttk.Button(left_buttons, text="💾保存", command=self.save_current_file, width=7, bootstyle="primary").pack(side=tk.LEFT, padx=2, ipady=5)
        ttk.Button(left_buttons, text="🗑清空", command=self.clear_form, width=7, bootstyle="danger").pack(side=tk.LEFT, padx=2, ipady=5)
        
        # 中间：导航按钮
        nav_frame = ttk.Frame(bottom_frame)
        nav_frame.pack(side=tk.LEFT, padx=20)
        
        ttk.Button(nav_frame, text="⬅上一个", command=self.load_previous_file, width=9, bootstyle="secondary").pack(side=tk.LEFT, padx=3, ipady=6)
        ttk.Button(nav_frame, text="➡下一个", command=self.load_next_file, width=9, bootstyle="secondary").pack(side=tk.LEFT, padx=3, ipady=6)
        
        # 右侧：手动添加按钮
        right_buttons = ttk.Frame(bottom_frame)
        right_buttons.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(right_buttons, text="📝手动添加批办单", command=self.manual_add_document, 
                  width=15, bootstyle="warning").pack(side=tk.LEFT, ipady=6)
        
    def create_log_tab(self, parent):
        """创建日志标签页"""
        self.log_text = scrolledtext.ScrolledText(parent, wrap=tk.WORD, 
                                                   font=(ui_font('Consolas'), 9))
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 清除日志按钮 - 使用现代化样式
        ttk.Button(parent, text="🗑清除日志", 
                  command=lambda: self.log_text.delete(1.0, tk.END),
                  bootstyle="danger-outline", width=10).pack(pady=5)
    
    def create_database_tab(self, parent):
        """创建数据库检索标签页"""
        # 搜索区域
        search_frame = ttk.Frame(parent)
        search_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 标题
        ttk.Label(search_frame, text="🔍 数据库检索", 
                 font=(ui_font('微软雅黑'), 12, 'bold')).pack(anchor=tk.W, pady=(0, 10))
        
        # 搜索条件行1
        row1 = ttk.Frame(search_frame)
        row1.pack(fill=tk.X, pady=5)
        
        ttk.Label(row1, text="关键词:", width=10).pack(side=tk.LEFT)
        self.db_keyword_entry = ttk.Entry(row1, width=30)
        self.db_keyword_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(row1, text="搜索字段:", width=10).pack(side=tk.LEFT, padx=(20, 0))
        self.db_field_var = tk.StringVar(value='all')
        field_combo = ttk.Combobox(row1, textvariable=self.db_field_var, width=15,
                                   values=['all', '来文单位', '文件标题', '收文编号', '来文字号', '收文途径'])
        field_combo.pack(side=tk.LEFT, padx=5)
        field_combo.config(state='readonly')
        
        # 搜索条件行2
        row2 = ttk.Frame(search_frame)
        row2.pack(fill=tk.X, pady=5)
        
        ttk.Label(row2, text="开始日期:", width=10).pack(side=tk.LEFT)
        self.db_start_date_entry = ttk.Entry(row2, width=15)
        self.db_start_date_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(row2, text="结束日期:", width=10).pack(side=tk.LEFT, padx=(20, 0))
        self.db_end_date_entry = ttk.Entry(row2, width=15)
        self.db_end_date_entry.pack(side=tk.LEFT, padx=5)
        
        # 按钮区
        btn_frame = ttk.Frame(search_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(btn_frame, text="🔍 搜索", command=self.search_database,
                  bootstyle="primary", width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="🔄 刷新", command=self.refresh_database,
                  bootstyle="info", width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="📊 统计", command=self.show_statistics,
                  bootstyle="success", width=10).pack(side=tk.LEFT, padx=5)
        
        # 数据库管理按钮（在统计按钮右侧，保持适当距离）
        ttk.Button(btn_frame, text="📁 新建数据库", command=self.create_new_database,
                  bootstyle="secondary", width=12).pack(side=tk.LEFT, padx=(20, 5))
        ttk.Button(btn_frame, text="📂 选择数据库", command=self.select_database,
                  bootstyle="secondary", width=12).pack(side=tk.LEFT, padx=5)
        
        # 当前数据库标签（在选择数据库按钮右侧）
        self.current_db_label = ttk.Label(btn_frame, text="", 
                                          font=(ui_font('微软雅黑'), 9), foreground='#666')
        self.current_db_label.pack(side=tk.LEFT, padx=(20, 5))
        self.update_current_db_label()
        
        # 导出Excel按钮（在当前数据库信息右侧）
        ttk.Button(btn_frame, text="📤 导出为Excel", command=self.export_to_excel,
                  bootstyle="warning", width=14).pack(side=tk.LEFT, padx=5)
        
        # 结果列表
        list_frame = ttk.Frame(parent)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        ttk.Label(list_frame, text="检索结果", font=(ui_font('微软雅黑'), 10, 'bold')).pack(anchor=tk.W, pady=(0, 5))
        
        # 创建Treeview（包含序号列）
        columns = ('序号', '收文编号', '来文单位', '文件标题', '收文日期')
        self.db_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=15)
        
        # 初始化排序状态字典（记录每列的排序方向）
        self.db_sort_state = {col: None for col in columns}  # None=未排序, True=升序, False=降序
        
        # 设置列标题和宽度，并绑定点击事件
        for col in columns:
            self.db_tree.heading(col, text=col, anchor='center' if col in ('序号', '收文编号', '收文日期') else 'w',
                                command=lambda c=col: self.sort_db_tree_column(c))
        
        self.db_tree.column('序号', width=60, anchor='center')
        self.db_tree.column('收文编号', width=120, anchor='center')
        self.db_tree.column('来文单位', width=120, anchor='w')
        self.db_tree.column('文件标题', width=450, anchor='w')
        self.db_tree.column('收文日期', width=120, anchor='center')
        
        # 滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.db_tree.yview)
        self.db_tree.configure(yscrollcommand=scrollbar.set)
        
        self.db_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 绑定双击事件（只加载到编辑器，不打开Word）
        self.db_tree.bind('<Double-Button-1>', lambda e: self.load_db_record_to_editor())
        
        # 操作按钮
        action_frame = ttk.Frame(parent)
        action_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        ttk.Button(action_frame, text="📂 打开Word批办单", command=self.open_db_record_in_word,
                  bootstyle="warning", width=18).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="👁 查看详情", command=self.view_db_detail,
                  bootstyle="info", width=12).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="🗑 删除记录", command=self.delete_db_record,
                  bootstyle="danger", width=12).pack(side=tk.LEFT, padx=5)

        # 初始加载数据
        self.root.after(200, self.refresh_database)
        
    def _on_tab_changed(self, event):
        """切换标签页时的处理"""
        selected = self.notebook.select()
        tab_text = self.notebook.tab(selected, 'text')
        if tab_text == '数据库检索':
            self.refresh_database()

    def create_status_bar(self):
        """创建状态栏"""
        # 创建状态栏框架
        status_frame = ttk.Frame(self.root, bootstyle="secondary")
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 状态文本
        self.status_bar = ttk.Label(status_frame, text="✅ 就绪", anchor=tk.W,
                                    bootstyle="inverse-secondary", 
                                    font=(ui_font('微软雅黑'), 9))
        self.status_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=3)
        
    def log(self, message):
        """添加日志"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"[{timestamp}] {message}\n"
        self.log_text.insert(tk.END, log_message)
        self.log_text.see(tk.END)
        self.status_bar.config(text=message)
        
    def refresh_file_list(self):
        """刷新待打印文件列表"""
        # 清空列表
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)
        
        # 获取待打印目录文件
        output_dir = self.path_config['待打印']
        if not os.path.exists(output_dir):
            return
        
        files = [f for f in os.listdir(output_dir) 
                if f.endswith('.docx') and not f.startswith('~$')]
        
        # 按文件编号数字排序（提取#后面的数字）
        files.sort(key=self.extract_number_from_filename)  # 按编号数字升序排列
        
        self.file_list = files
        # 清除旧的复选框状态
        self.file_checkboxes.clear()
        
        for idx, filename in enumerate(files, 1):
            item_id = self.file_tree.insert('', 'end', values=('☐', idx, filename))
            self.file_checkboxes[item_id] = False  # 默认不选中
        
        self.log(f"已刷新待打印文件列表，共 {len(files)} 个文件")
    
    def refresh_source_list(self):
        """刷新待处理文件列表"""
        # 清空列表
        for item in self.source_tree.get_children():
            self.source_tree.delete(item)
        
        # 按文件名升序排列
        self.source_files.sort(key=lambda x: os.path.basename(x))
        
        for idx, filepath in enumerate(self.source_files, 1):
            filename = os.path.basename(filepath)
            self.source_tree.insert('', 'end', values=(idx, filename))
        
        self.log(f"已刷新待处理文件列表，共 {len(self.source_files)} 个文件")
        
        # 刷新后重新应用高亮
        self.highlight_source_search()
    
    def highlight_source_search(self, *args):
        """高亮显示待处理文件列表中的搜索关键词"""
        search_text = self.source_search_var.get().strip()
        
        # 清除所有项的tag
        for item in self.source_tree.get_children():
            self.source_tree.item(item, tags=())
        
        # 如果搜索框为空，不做高亮
        if not search_text:
            return
        
        # 遍历所有项，检查是否包含搜索文本
        for item in self.source_tree.get_children():
            values = self.source_tree.item(item)['values']
            if len(values) >= 2:
                filename = values[1]  # 文件名在第2列
                # 不区分大小写搜索
                if search_text.lower() in filename.lower():
                    # 添加高亮tag
                    self.source_tree.item(item, tags=('highlight',))
    
    def refresh_default_source_folder(self):
        """从默认待处理文件夹刷新待处理文件列表"""
        default_folder = self.path_config.get('默认待处理文件夹', '')
        
        if not default_folder or not os.path.exists(default_folder):
            messagebox.showwarning("警告", "请先设置默认待处理文件夹！\n\n可以在 工具 → 设置默认待处理文件夹 中进行设置。")
            return
        
        # 第一步：移除已经不存在的文件
        removed_files = []
        valid_files = []
        for filepath in self.source_files:
            if os.path.exists(filepath):
                valid_files.append(filepath)
            else:
                removed_files.append(os.path.basename(filepath))
        
        self.source_files = valid_files
        
        # 第二步：扫描文件夹中的PDF、DOCX、DOC文件
        new_files = []
        for root, dirs, files in os.walk(default_folder):
            for file in files:
                if file.lower().endswith(('.pdf', '.docx', '.doc')) and not file.startswith('~$'):
                    filepath = os.path.join(root, file)
                    if filepath not in self.source_files:
                        new_files.append(filepath)
        
        # 第三步：添加新文件到列表
        self.source_files.extend(new_files)
        self.refresh_source_list()
        
        # 第四步：显示刷新结果
        message_parts = []
        if new_files:
            message_parts.append(f"新增 {len(new_files)} 个文件")
        if removed_files:
            message_parts.append(f"移除 {len(removed_files)} 个不存在的文件")
        
        if message_parts:
            messagebox.showinfo("刷新完成", "刷新结果：\n• " + "\n• ".join(message_parts))
            if removed_files:
                self.log(f"已移除不存在的文件: {', '.join(removed_files[:5])}{'...' if len(removed_files) > 5 else ''}")
            if new_files:
                self.log(f"已添加新文件: {len(new_files)} 个")
        else:
            messagebox.showinfo("提示", "列表已是最新，无需更新。")
            self.log("待处理文件列表已是最新")
    
    def set_default_source_folder(self):
        """设置默认待处理文件夹"""
        folder = filedialog.askdirectory(title="选择默认待处理文件夹")
        if folder:
            self.path_config['默认待处理文件夹'] = folder
            
            # 保存配置
            try:
                with open(get_appdata_path('path_config.json'), 'w', encoding='utf-8') as f:
                    json.dump(self.path_config, f, ensure_ascii=False, indent=2)
                messagebox.showinfo("成功", f"默认待处理文件夹已设置为：\n{folder}\n\n将文件放入此文件夹后，点击'刷新列表'按钮即可自动加载。")
                self.log(f"默认待处理文件夹已设置为: {folder}")
            except Exception as e:
                messagebox.showerror("错误", f"保存配置失败：{str(e)}")
                self.log(f"保存配置失败: {e}")
    
    def refresh_all_lists(self):
        """刷新所有列表"""
        self.refresh_file_list()
        self.refresh_source_list()
        
        
    def load_file_auto(self):
        """自动加载选中的文件（单击或键盘选择时触发）"""
        selection = self.file_tree.selection()
        if not selection:
            return
        
        values = self.file_tree.item(selection[0])['values']
        if len(values) >= 3:  # 现在有3列：选择、序号、文件名
            filename = values[2]  # 文件名在第3列（索引2）
            filepath = os.path.join(self.path_config['待打印'], filename)
            
            # 切换到批办单编辑标签页（索引0）
            self.notebook.select(0)
            
            # 加载文件内容
            self._load_file_content(filepath, filename)
    
    def load_file(self):
        """加载选中的文件进行编辑（手动调用）"""
        selection = self.file_tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选择要编辑的文件")
            return
        
        values = self.file_tree.item(selection[0])['values']
        if len(values) >= 3:  # 现在有3列：选择、序号、文件名
            filename = values[2]  # 文件名在第3列（索引2）
            filepath = os.path.join(self.path_config['待打印'], filename)
            self._load_file_content(filepath, filename)
    
    def _load_file_content(self, filepath, filename):
        """加载文件内容的内部方法"""
        
        self.current_file = filepath
        self.file_title_label.config(text=f"📝 正在编辑: {filename}")
        
        # 读取Word文档
        try:
            doc = Document(filepath)
            if len(doc.tables) == 0:
                messagebox.showerror("错误", "文档中没有表格")
                return
            
            table = doc.tables[0]
            
            # 清空所有输入框
            for field_name, entry in self.field_entries.items():
                if isinstance(entry, tk.Text):
                    entry.delete(1.0, tk.END)
                elif isinstance(entry, (tk.Entry, ttk.Entry)):
                    # 如果是只读字段（如收文编号），需要临时解除只读状态
                    if field_name == '收文编号':
                        entry.configure(state='normal')
                        entry.delete(0, tk.END)
                        entry.configure(state='readonly')
                    else:
                        entry.delete(0, tk.END)
            
            # 根据实际表格位置读取数据
            # 表格位置映射（行,列）
            field_mapping = {
                '来文单位': (0, 2),
                '来文字号': (1, 2),
                '收文日期': (1, 9),
                '来文类型': (2, 2),
                '收文途径': (2, 9),
                '收文编号': (3, 9),
                '文件标题': (4, 2),
                '领导批示': (5, 2),
                '承办意见': (6, 2),
                '办理情况': (9, 2),
            }
            
            # 读取并填充字段
            for field_name, (row, col) in field_mapping.items():
                if field_name in self.field_entries:
                    try:
                        if row < len(table.rows) and col < len(table.rows[row].cells):
                            # 获取单元格文本
                            cell_text = table.rows[row].cells[col].text
                            
                            # 对于"承办意见"（拟办意见）和"领导批示"，只去除尾部空白符，保留前置空格
                            if field_name in ['承办意见', '领导批示']:
                                cell_text = cell_text.rstrip('\n\r ')
                            else:
                                # 其他字段正常去除前后空白符
                                cell_text = cell_text.strip()
                            
                            # 填充到对应的输入框
                            widget = self.field_entries[field_name]
                            if isinstance(widget, tk.Text):
                                # 多行文本框
                                widget.insert(1.0, cell_text)
                            elif isinstance(widget, (tk.Entry, ttk.Entry)):
                                # 单行输入框
                                # 如果是只读字段（如收文编号），需要临时解除只读状态
                                if field_name == '收文编号':
                                    widget.configure(state='normal')
                                    widget.insert(0, cell_text)
                                    widget.configure(state='readonly')
                                else:
                                    widget.insert(0, cell_text)
                    except Exception as e:
                        self.log(f"读取字段 {field_name} 失败: {e}")
            
            # 保存原始收文编号，用于后续更新数据库
            # 只有当数据库中存在该记录时，才设置original_receipt_number
            if '收文编号' in self.field_entries:
                receipt_widget = self.field_entries['收文编号']
                if isinstance(receipt_widget, (tk.Entry, ttk.Entry)):
                    receipt_number = receipt_widget.get()
                    if receipt_number:
                        # 检查数据库中是否存在该记录
                        existing_record = self.db.get_complaint_by_receipt_number(receipt_number)
                        if existing_record:
                            self.original_receipt_number = receipt_number
                            self.log(f"已记录原始收文编号（数据库中存在）: {self.original_receipt_number}")
                        else:
                            self.original_receipt_number = None
                            self.log(f"收文编号 {receipt_number} 不在数据库中，将作为新记录添加")
            
            self.log(f"已加载文件: {filename}")
            
        except Exception as e:
            messagebox.showerror("错误", f"加载文件失败: {e}")
            self.log(f"加载文件失败: {e}")
        
    def save_current_file(self):
        """保存当前编辑的文件"""
        if not self.current_file:
            messagebox.showwarning("提示", "没有打开的文件")
            return
        
        try:
            # 读取Word文档
            doc = Document(self.current_file)
            if len(doc.tables) == 0:
                messagebox.showerror("错误", "文档中没有表格")
                return
            
            table = doc.tables[0]
            generator = WordGenerator()
            
            # 表格位置映射（行,列）
            field_mapping = {
                '来文单位': (0, 2),
                '来文字号': (1, 2),
                '收文日期': (1, 9),
                '来文类型': (2, 2),
                '收文途径': (2, 9),
                '收文编号': (3, 9),
                '文件标题': (4, 2),
                '领导批示': (5, 2),
                '承办意见': (6, 2),
                '办理情况': (9, 2),
            }
            
            # 从输入框读取值并写入文档
            for field_name, (row, col) in field_mapping.items():
                if field_name in self.field_entries:
                    try:
                        widget = self.field_entries[field_name]
                        
                        # 获取输入框的值
                        if isinstance(widget, tk.Text):
                            value = widget.get(1.0, tk.END)
                            # 只去掉末尾的换行符，保留前置空格
                            value = value.rstrip('\n\r')
                        elif isinstance(widget, (tk.Entry, ttk.Entry)):
                            value = widget.get()
                        else:
                            value = ''
                        
                        # 写入到对应单元格（使用保持格式的方法）
                        if row < len(table.rows) and col < len(table.rows[row].cells):
                            cell = table.rows[row].cells[col]
                            
                            # 根据字段应用不同的格式
                            if field_name in ['来文单位', '文件标题', '办理情况']:
                                # 添加前置空格
                                generator.set_cell_text_keep_format(cell, value, add_prefix_space=True)
                            elif field_name == '来文字号':
                                # 按最长行字符数定字号，按行数定行距
                                lns = [l for l in value.split('\n') if l.strip()]
                                max_char = max((len(l) for l in lns), default=0)
                                line_count = len(lns)
                                if max_char >= 33:
                                    font_size_pt = 9
                                elif max_char >= 29:
                                    font_size_pt = 10.5
                                elif max_char >= 25:
                                    font_size_pt = 12
                                elif line_count >= 4:
                                    font_size_pt = 10.5
                                elif line_count == 3:
                                    font_size_pt = 12
                                else:
                                    font_size_pt = 14
                                if line_count >= 4:
                                    line_spacing_pt = 10
                                elif line_count == 3:
                                    line_spacing_pt = 11
                                else:
                                    line_spacing_pt = 17
                                generator.set_cell_text_with_custom_format(
                                    cell, value,
                                    font_size_pt=font_size_pt,
                                    line_spacing_pt=line_spacing_pt,
                                    alignment='center'
                                )
                            elif field_name in ['收文日期', '来文类型', '收文途径', '收文编号']:
                                # 水平居中
                                generator.set_cell_text_keep_format(cell, value, alignment='center')
                            elif field_name == '承办意见':
                                # 拟办意见：每个段落添加前置空格，固定28磅行距（兼容Word 2007）
                                generator.set_cell_text_keep_format(cell, value, add_prefix_space=True, line_spacing_pt=28)
                            else:
                                # 其他字段（领导批示）保持默认
                                generator.set_cell_text_keep_format(cell, value)
                    except Exception as e:
                        self.log(f"保存字段 {field_name} 失败: {e}")
            
            # 保存文档
            doc.save(self.current_file)
            
            # 记录输入历史（来文单位、收文途径、文件标题）
            for field_name in ['来文单位', '收文途径', '文件标题']:
                if field_name in self.field_entries:
                    widget = self.field_entries[field_name]
                    if hasattr(widget, 'get'):
                        value = widget.get()
                        self.input_history.add_value(field_name, value)
            
            # 录入数据库
            self.save_to_database(os.path.basename(self.current_file))
            
            messagebox.showinfo("成功", "文件保存成功")
            self.log(f"已保存文件: {os.path.basename(self.current_file)}")
            
        except Exception as e:
            messagebox.showerror("错误", f"保存文件失败: {e}")
            self.log(f"保存文件失败: {e}")
    
    def save_to_database(self, filename):
        """将当前表单数据保存到数据库"""
        try:
            # 从输入框提取数据
            data = {'文件名': filename}
            for field_name, widget in self.field_entries.items():
                if isinstance(widget, tk.Text):
                    value = widget.get(1.0, tk.END).rstrip('\n\r')
                elif isinstance(widget, (tk.Entry, ttk.Entry)):
                    value = widget.get()
                else:
                    value = ''
                data[field_name] = value
            
            # 调试信息
            title = data.get('文件标题', '')
            title_preview = title[:20] + '...' if len(title) > 20 else title
            receipt_number = data.get('收文编号', '')
            self.log(f"保存到数据库: 收文编号={receipt_number}, 来文单位={data.get('来文单位', '')}, 标题={title_preview}")
            
            # 判断是更新还是添加记录
            if self.original_receipt_number:
                # 有原始收文编号，说明是从数据库或文件加载的已有记录，使用原始收文编号更新
                if self.db.update_complaint(self.original_receipt_number, data):
                    self.log(f"✓ 已更新数据库记录: {self.original_receipt_number}")
                else:
                    self.log(f"✗ 更新数据库记录失败: {self.original_receipt_number}")
            else:
                # 没有原始收文编号，说明是新建记录或从待处理文件夹加载的文件
                existing_record = self.db.get_complaint_by_receipt_number(receipt_number) if receipt_number else None
                if existing_record:
                    # 记录已存在，更新
                    if self.db.update_complaint(receipt_number, data):
                        self.log(f"✓ 已更新数据库记录: {receipt_number}")
                    else:
                        self.log(f"✗ 更新数据库记录失败: {receipt_number}")
                else:
                    # 添加新记录
                    if self.db.add_complaint(data):
                        self.log(f"✓ 已添加到数据库: {receipt_number or filename}")
                    else:
                        self.log(f"✗ 添加到数据库失败: {filename}")
                
        except Exception as e:
            self.log(f"录入数据库异常: {e}")
            import traceback
            traceback.print_exc()
        
    def fill_today(self):
        """填充今天的日期"""
        now = datetime.now()
        today = f"{now.year}-{now.month}-{now.day}"
        widget = self.field_entries['收文日期']
        if isinstance(widget, tk.Text):
            widget.delete(1.0, tk.END)
            widget.insert(1.0, today)
        elif isinstance(widget, (tk.Entry, ttk.Entry)):
            widget.delete(0, tk.END)
            widget.insert(0, today)
    
    def clear_form(self):
        """清空当前表单内容（不改变Word文档）"""
        if not messagebox.askyesno("确认", "确定要清空当前表单内容吗？\n\n注意：这不会修改已生成的Word文档。"):
            return
        
        self._clear_form_silently()
        messagebox.showinfo("完成", "表单内容已清空")
    
    def _clear_form_silently(self):
        """静默清空表单（无确认对话框）"""
        # 清空所有输入框
        for field_name, entry in self.field_entries.items():
            if isinstance(entry, tk.Text):
                entry.delete(1.0, tk.END)
            elif isinstance(entry, (tk.Entry, ttk.Entry)):
                # 如果是只读字段（如收文编号），需要临时解除只读状态
                if field_name == '收文编号':
                    entry.configure(state='normal')
                    entry.delete(0, tk.END)
                    entry.configure(state='readonly')
                else:
                    entry.delete(0, tk.END)
        
        # 清空当前文件标识
        self.current_file = None
        self.original_receipt_number = None  # 清空原始收文编号
        self.file_title_label.config(text="📝 请选择文件进行编辑")
        # 注意：不清空收文编号显示，因为它显示的是下一个将生成的编号
        
        self.log("已清空表单内容")
    
    def update_next_receipt_number(self):
        """更新下一个收文编号的显示（使用当前数据库年份）"""
        try:
            from datetime import datetime
            # 获取当前数据库年份并使用对应的计数器
            db_year = self.get_current_db_year()
            counter = CounterManager(year=db_year)
            year_suffix = str(datetime.now().year)[-2:]
            next_receipt = f"C-{year_suffix}-{counter.receipt_number}"
            self.receipt_number_label.config(text=f"📋 下一个收文编号: {next_receipt}")
            self.log(f"已更新收文编号显示: {next_receipt}")
        except Exception as e:
            self.log(f"更新收文编号显示失败: {e}")
    
    def rebuild_quick_buttons(self):
        """根据配置动态重建快捷按钮"""
        for widget in self.quick_buttons_frame.winfo_children():
            widget.destroy()
        styles = ['primary', 'success', 'info', 'warning', 'danger', 'dark']
        for i, button_name in enumerate(self.quick_buttons_config.keys()):
            style = styles[i % len(styles)]
            btn = ttk.Button(
                self.quick_buttons_frame,
                text=button_name,
                command=lambda name=button_name: self.fill_quick_button(name),
                width=12,
                bootstyle=f"{style}-outline"
            )
            btn.pack(pady=4, padx=10, ipady=4)

    def load_quick_buttons_config(self):
        """加载快捷按钮配置"""
        config_file = get_appdata_path('quick_buttons_config.json')
        
        # 默认配置
        default_config = {
            '医政投诉': {
                '拟办意见': '    请莉莉同志阅，请吕宁同志阅示。\r\n    建议：请医政科妥处，请法监科亦阅。',
                '办理情况': '信息中心已转医政科。'
            },
            '疾控投诉': {
                '拟办意见': '    请莉莉同志阅，请高鹏同志阅示。\r\n    建议：请疾控中心（监督所、健教所）妥处，请法监科亦阅。',
                '办理情况': '信息中心已转疾控中心（监督所、健教所）。'
            },
            '公卫投诉': {
                '拟办意见': '    请莉莉同志阅，请吕宁同志阅示。\r\n    建议：请公卫科妥处，请法监科亦阅。',
                '办理情况': '信息中心已转公卫科。'
            },
            '中医投诉': {
                '拟办意见': '    请莉莉同志阅，请吕宁同志阅示。\r\n    建议：请中医科妥处，请法监科亦阅。',
                '办理情况': '信息中心已转中医科。'
            },
            '妇计投诉': {
                '拟办意见': '    请莉莉同志阅，请高鹏同志阅示。\r\n    建议：请妇计中心妥处，请法监科亦阅。',
                '办理情况': '信息中心已转妇计中心。'
            },
            '医托投诉': {
                '拟办意见': '    请莉莉同志阅，请吕宁同志阅示，请高鹏同志亦阅。\r\n    建议：请医政科牵头办理，及时移交线索并做好纠纷化解，请法监科联合公安、市场监管局组织集中打击，请疾控中心（监督所、健教所）严查严处。',
                '办理情况': '信息中心已转医政科。'
            }
        }
        
        # 尝试从文件加载配置
        try:
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                return config
        except Exception as e:
            print(f"加载配置文件失败: {e}")
        
        # 保存默认配置
        try:
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(default_config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存默认配置失败: {e}")
        
        return default_config
    
    def save_quick_buttons_config(self, config):
        """保存快捷按钮配置"""
        config_file = get_appdata_path('quick_buttons_config.json')
        try:
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            self.log("配置已保存")
            return True
        except Exception as e:
            messagebox.showerror("错误", f"保存配置失败: {e}")
            return False
    
    def load_qrcode_config(self):
        """加载二维码配置"""
        config_file = get_appdata_path('qrcode_config.json')
        
        # 默认配置
        default_config = {
            'enabled': False,
            'size': 'small',  # 'large' (2.7cm×2.7cm) or 'small' (1.7cm×1.7cm)
            'fields': {
                '收文编号': True,
                '收文日期': False,
                '来文字号': False,
                '来文单位': False,
                '收文途径': False,
                '文件标题': False
            }
        }
        
        # 尝试从文件加载配置
        try:
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                return config
        except Exception as e:
            print(f"加载二维码配置文件失败: {e}")
        
        # 保存默认配置
        try:
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(default_config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存默认二维码配置失败: {e}")
        
        return default_config
    
    def save_qrcode_config(self, config):
        """保存二维码配置"""
        config_file = get_appdata_path('qrcode_config.json')
        try:
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            self.log("二维码配置已保存")
            return True
        except Exception as e:
            messagebox.showerror("错误", f"保存二维码配置失败: {e}")
            return False
    
    def load_window_config(self):
        """加载窗口布局配置"""
        config_file = get_appdata_path('window_config.json')
        default_config = {
            'maximized': True,  # 默认最大化
            'geometry': '1400x900+100+100',  # 默认大小和位置
            'paned_positions': [240, 480],  # 分割窗口位置
            'theme': 'litera'  # 默认主题
        }
        
        try:
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    # 合并默认配置
                    for key in default_config:
                        if key not in config:
                            config[key] = default_config[key]
                    return config
        except Exception as e:
            self.log(f"加载窗口配置失败: {e}")
        
        return default_config
    
    def save_window_config(self):
        """保存窗口布局配置"""
        config_file = get_appdata_path('window_config.json')
        try:
            # 读取现有配置（保留主题等设置）
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
            else:
                config = {}
            
            # 保存窗口状态（Windows 用 zoomed，Linux 用 Tk 的 zoomed 属性兼容写法）
            _state = ''
            try:
                _state = self.root.state()
            except Exception:
                pass
            config['maximized'] = (_state == 'zoomed') or bool(self.root.attributes('-zoomed'))
            
            # 保存窗口大小和位置
            config['geometry'] = self.root.geometry()
            
            # 保存分割窗口位置
            if self.main_paned:
                sash_positions = []
                # PanedWindow有两个分割条（三栏布局）
                try:
                    for i in range(2):
                        pos = self.main_paned.sashpos(i)
                        # 边界检查：防止保存太极端的值
                        if pos < 100:
                            pos = 240
                        if i == 1 and pos < 300:
                            pos = 480
                        sash_positions.append(pos)
                    config['paned_positions'] = sash_positions
                    print(f"保存分割条位置: {sash_positions}")
                except Exception as e:
                    print(f"获取分割条位置失败: {e}")
            
            # 保存配置
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
                f.flush()
                os.fsync(f.fileno())
            
            print(f"[OK] 窗口配置已保存")
            self.log("窗口布局已保存")
        except Exception as e:
            print(f"保存窗口配置失败: {e}")
            self.log(f"保存窗口配置失败: {e}")
    
    def apply_window_config(self):
        """应用窗口配置"""
        try:
            # 先设置窗口大小和位置
            if not self.window_config.get('maximized', True):
                # 如果不是最大化，设置保存的几何尺寸
                self.root.geometry(self.window_config.get('geometry', '1400x900+100+100'))
            else:
                # 设置默认大小（在最大化之前）
                self.root.geometry('1400x900')
            
            # 更新窗口以确保geometry生效
            self.root.update_idletasks()
            
            # 设置最大化状态（Linux 用 -zoomed 属性，Windows 用 state('zoomed')）
            if self.window_config.get('maximized', True):
                self._maximize_window()
        except Exception as e:
            self.log(f"应用窗口配置失败: {e}")
            # 如果失败，默认最大化
            try:
                self._maximize_window()
            except Exception:
                self.root.geometry('1400x900')

    def _maximize_window(self):
        """跨平台最大化窗口，保持与 Windows 版一致的初始布局。"""
        if IS_WINDOWS:
            self.root.state('zoomed')
        else:
            try:
                self.root.attributes('-zoomed', True)
            except Exception:
                w = self.root.winfo_screenwidth()
                h = self.root.winfo_screenheight()
                self.root.geometry(f'{w}x{h}+0+0')
    
    def restore_paned_positions(self):
        """恢复分割窗口位置"""
        try:
            if self.main_paned and 'paned_positions' in self.window_config:
                positions = self.window_config['paned_positions']
                
                # 获取窗口宽度，确保分割条位置合理
                try:
                    window_width = self.root.winfo_width()
                    if window_width < 100:  # 窗口还未渲染完成
                        return  # 跳过本次，等待下次调用
                except:
                    return
                
                success = True
                for i, pos in enumerate(positions):
                    if i < 2:  # 只有两个分割条
                        try:
                            # 边界检查：防止分割条位置太极端
                            if pos < 100:  # 左侧至少100像素
                                pos = 240
                            if i == 1 and pos < 300:  # 右侧至少300像素（包含左侧+中间）
                                pos = 480
                            
                            # 确保位置不超过窗口宽度
                            if pos > window_width - 100:
                                pos = window_width - 300
                            
                            # 设置分割条位置
                            self.main_paned.sashpos(i, pos)
                            
                            # 强制更新
                            self.main_paned.update_idletasks()
                            
                            # 验证是否设置成功
                            actual_pos = self.main_paned.sashpos(i)
                            if abs(actual_pos - pos) > 10:  # 允许10像素的误差
                                success = False
                            
                            print(f"分割条 {i}: 期望={pos}, 实际={actual_pos}, 窗口宽度={window_width}")
                        except Exception as e:
                            print(f"设置分割条 {i} 失败: {e}")
                            success = False
                
                if success:
                    self.log(f"分割窗口位置已恢复: {positions}")
                    print("[OK] 布局恢复成功")
        except Exception as e:
            print(f"恢复分割窗口位置失败: {e}")
    
    def reset_window_layout(self):
        """重置窗口布局"""
        if messagebox.askyesno("确认", "确定要重置窗口布局为默认设置吗？\n\n这将恢复默认的三栏布局。"):
            try:
                # 重置分割条位置
                if self.main_paned:
                    self.main_paned.sashpos(0, 240)
                    self.main_paned.sashpos(1, 480)
                    self.log("窗口布局已重置为默认设置")
                    messagebox.showinfo("成功", "窗口布局已重置！\n\n分割条位置已恢复为默认值。")
            except Exception as e:
                self.log(f"重置布局失败: {e}")
                messagebox.showerror("错误", f"重置布局失败: {e}")
    
    def change_theme(self):
        """切换主题"""
        # 创建主题选择对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("主题切换")
        dialog.geometry("600x600")  # 增加宽度，确保按钮完整显示
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 标题
        title_frame = ttk.Frame(dialog, bootstyle="info")
        title_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(title_frame, text="🎨 选择主题", 
                 font=(ui_font('微软雅黑'), 14, 'bold'),
                 bootstyle="inverse-info").pack(pady=12)
        
        # 说明
        ttk.Label(dialog, text="选择您喜欢的主题，切换后将重启程序",
                 font=(ui_font('微软雅黑'), 10)).pack(pady=5)
        
        # 获取当前主题
        current_theme = self.window_config.get('theme', 'litera')
        
        # 主题列表
        themes = [
            ("Litera", "litera", "简洁明亮（推荐）"),
            ("Cosmo", "cosmo", "经典扁平化"),
            ("Flatly", "flatly", "现代扁平化"),
            ("Minty", "minty", "清新绿色"),
            ("Lumen", "lumen", "简洁商务"),
            ("Sandstone", "sandstone", "温暖橙色"),
            ("Yeti", "yeti", "经典蓝色"),
            ("Pulse", "pulse", "深色系"),
            ("United", "united", "统一商务"),
            ("Morph", "morph", "渐变色"),
            ("Journal", "journal", "杂志风格"),
            ("Darkly", "darkly", "深色暗黑"),
            ("Superhero", "superhero", "超级英雄"),
            ("Solar", "solar", "太阳能"),
            ("Cyborg", "cyborg", "黑客风格"),
            ("Vapor", "vapor", "蒸汽波"),
        ]
        
        # 创建滚动框架
        canvas = tk.Canvas(dialog, highlightthickness=0)
        scrollbar = ttk.Scrollbar(dialog, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # 选择变量
        selected_theme = tk.StringVar(value=current_theme)
        
        # 添加主题单选按钮
        for display_name, theme_name, description in themes:
            frame = ttk.Frame(scrollable_frame)
            frame.pack(fill=tk.X, padx=20, pady=3)
            
            # 单选按钮
            radio = ttk.Radiobutton(
                frame, 
                text=f"{display_name}",
                value=theme_name,
                variable=selected_theme,
                bootstyle="info"
            )
            radio.pack(side=tk.LEFT, padx=5)
            
            # 描述
            desc_label = ttk.Label(frame, text=description, font=(ui_font('微软雅黑'), 9))
            desc_label.pack(side=tk.LEFT, padx=10)
            
            # 当前主题标记
            if theme_name == current_theme:
                ttk.Label(frame, text="✓ 当前", 
                         font=(ui_font('微软雅黑'), 9, 'bold'),
                         bootstyle="success").pack(side=tk.RIGHT, padx=5)
        
        canvas.pack(side="left", fill="both", expand=True, padx=(20, 0), pady=10)
        scrollbar.pack(side="right", fill="y", padx=(0, 20), pady=10)
        
        # 底部按钮区域 - 使用醒目的样式
        button_frame = ttk.Frame(dialog, bootstyle="light")
        button_frame.pack(fill=tk.X, padx=0, pady=0, side=tk.BOTTOM)
        
        # 添加分隔线
        ttk.Separator(button_frame, orient='horizontal').pack(fill=tk.X, pady=0)
        
        # 按钮容器
        btn_container = ttk.Frame(button_frame)
        btn_container.pack(pady=15)
        
        def apply_theme():
            """应用主题"""
            new_theme = selected_theme.get()
            if new_theme == current_theme:
                messagebox.showinfo("提示", "您选择的主题与当前主题相同")
                return
            
            # 保存主题到配置 - 确保保存成功
            try:
                config_file = get_appdata_path('window_config.json')
                
                # 读取现有配置
                if os.path.exists(config_file):
                    with open(config_file, 'r', encoding='utf-8') as f:
                        config = json.load(f)
                    print(f"读取到的配置: {config}")
                else:
                    config = {
                        'maximized': True,
                        'geometry': '1400x900+100+100',
                        'paned_positions': [240, 480]
                    }
                    print("配置文件不存在，创建新配置")
                
                # 更新主题
                config['theme'] = new_theme
                print(f"准备保存主题: {new_theme}")
                
                # 保存配置 - 强制写入
                with open(config_file, 'w', encoding='utf-8') as f:
                    json.dump(config, f, ensure_ascii=False, indent=2)
                    f.flush()  # 强制刷新缓冲区
                    os.fsync(f.fileno())  # 确保写入磁盘
                
                # 验证保存
                with open(config_file, 'r', encoding='utf-8') as f:
                    verify_config = json.load(f)
                    saved_theme = verify_config.get('theme', 'not found')
                    print(f"验证保存结果，主题: {saved_theme}")
                    if saved_theme != new_theme:
                        raise Exception(f"主题保存验证失败: 期望 {new_theme}, 实际 {saved_theme}")
                
                self.log(f"主题已保存并验证: {new_theme}")
                
                # 提示重启
                if messagebox.askyesno("确认", 
                    f"主题将切换为 {new_theme}\n\n程序需要重启才能应用新主题\n\n是否立即重启？"):
                    # 设置标志，表示是为了切换主题而重启
                    self.is_restarting_for_theme = True
                    # 关闭对话框
                    dialog.destroy()
                    # 稍作延迟确保对话框关闭
                    self.root.after(100, self.restart_application)
            except Exception as e:
                messagebox.showerror("错误", f"保存主题配置失败: {e}")
                self.log(f"保存主题失败: {e}")
                import traceback
                traceback.print_exc()
        
        # 大号按钮，更醒目 - 居中显示
        ttk.Button(btn_container, text="✅ 应用主题", 
                  command=apply_theme,
                  bootstyle="success", width=18).pack(side=tk.LEFT, padx=15, ipady=10)
        
        ttk.Button(btn_container, text="❌ 取消", 
                  command=dialog.destroy,
                  bootstyle="secondary", width=18).pack(side=tk.LEFT, padx=15, ipady=10)
    
    def restart_application(self):
        """重启应用程序"""
        try:
            import sys
            import subprocess
            import time
            
            # 确保配置已经保存（等待一下）
            time.sleep(0.2)
            
            # 验证配置文件
            config_file = get_appdata_path('window_config.json')
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    print(f"重启前验证配置，主题: {config.get('theme', 'not found')}")
            
            # 关闭当前窗口
            self.root.destroy()

            # 重新启动程序：打包后的 Linux 可执行文件自身就是入口，
            # 源码模式则用 Python 解释器重新执行入口脚本。
            if getattr(sys, 'frozen', False):
                subprocess.Popen([sys.executable])
            else:
                script = os.path.join(pc.get_app_dir(), 'gui_app.py')
                subprocess.Popen([sys.executable, script])
        except Exception as e:
            messagebox.showerror("错误", f"重启失败: {e}\n\n请手动重启程序")
    
    def on_closing(self):
        """窗口关闭事件"""
        try:
            # 如果是为了切换主题而重启，只保存窗口布局，不覆盖主题设置
            if self.is_restarting_for_theme:
                print("主题切换重启中，保存窗口布局但不覆盖主题")
                config_file = get_appdata_path('window_config.json')
                
                # 读取现有配置（包含新保存的主题）
                if os.path.exists(config_file):
                    with open(config_file, 'r', encoding='utf-8') as f:
                        config = json.load(f)
                else:
                    config = {}
                
                # 保存窗口状态（是否最大化）
                config['maximized'] = (self.root.state() == 'zoomed')
                
                # 保存窗口大小和位置
                config['geometry'] = self.root.geometry()
                
                # 保存分割窗口位置
                if self.main_paned:
                    sash_positions = []
                    try:
                        for i in range(2):
                            pos = self.main_paned.sashpos(i)
                            if pos < 100:
                                pos = 240
                            if i == 1 and pos < 300:
                                pos = 480
                            sash_positions.append(pos)
                        config['paned_positions'] = sash_positions
                    except:
                        pass
                
                # 保存配置（保留主题设置）
                with open(config_file, 'w', encoding='utf-8') as f:
                    json.dump(config, f, ensure_ascii=False, indent=2)
                    f.flush()
                    os.fsync(f.fileno())
                
                print(f"已保存窗口布局，主题保持为: {config.get('theme', 'unknown')}")
            else:
                # 正常关闭，保存所有配置
                print("正常关闭，保存完整配置")
                self.save_window_config()
        except Exception as e:
            self.log(f"关闭时保存配置失败: {e}")
            print(f"保存配置失败: {e}")
        
        # 关闭窗口
        self.root.destroy()
    
    def open_quick_fill_settings(self):
        """打开快捷按钮设置对话框"""
        settings_win = tk.Toplevel(self.root)
        settings_win.title("快捷按钮设置")
        settings_win.geometry("700x650")
        settings_win.transient(self.root)
        settings_win.grab_set()
        
        ttk.Label(settings_win, text="设置每个快捷按钮的预设内容",
                  font=(ui_font('微软雅黑'), 10)).pack(pady=10)
        
        notebook = ttk.Notebook(settings_win)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        editors = {}
        
        def add_tab(button_name, niban='', banli=''):
            tab = ttk.Frame(notebook)
            notebook.add(tab, text=button_name)
            frame = ttk.Frame(tab)
            frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            ttk.Label(frame, text="拟办意见:", font=(ui_font('微软雅黑'), 10, 'bold')).pack(anchor='w', pady=5)
            niban_text = tk.Text(frame, height=7, width=70, font=(ui_font('仿宋_GB2312'), 10), wrap=tk.WORD,
                                undo=True, maxundo=-1)
            niban_text.pack(fill=tk.BOTH, expand=True, pady=5)
            niban_text.insert(1.0, niban)
            self.bind_text_shortcuts(niban_text)
            
            ttk.Label(frame, text="办理情况:", font=(ui_font('微软雅黑'), 10, 'bold')).pack(anchor='w', pady=5)
            banli_text = tk.Text(frame, height=3, width=70, font=(ui_font('仿宋_GB2312'), 10), wrap=tk.WORD,
                                undo=True, maxundo=-1)
            banli_text.pack(fill=tk.BOTH, expand=True, pady=5)
            banli_text.insert(1.0, banli)
            self.bind_text_shortcuts(banli_text)
            
            def delete_tab(name=button_name):
                if messagebox.askyesno("确认删除", f"确定要删除按钮「{name}」吗？\n\n此操作不可撤销。", parent=settings_win):
                    idx = list(editors.keys()).index(name)
                    notebook.forget(idx)
                    del editors[name]
            
            ttk.Button(frame, text="🗑 删除此按钮", command=delete_tab,
                      bootstyle="danger-outline", width=14).pack(anchor='e', pady=(5, 0))
            
            editors[button_name] = {'拟办意见': niban_text, '办理情况': banli_text}
            notebook.select(tab)
        
        for button_name, cfg in self.quick_buttons_config.items():
            add_tab(button_name, cfg.get('拟办意见', ''), cfg.get('办理情况', ''))
        
        bottom_frame = ttk.Frame(settings_win)
        bottom_frame.pack(fill=tk.X, padx=10, pady=8)
        
        def add_new_button():
            name = simpledialog.askstring("新增按钮", "请输入按钮名称：", parent=settings_win)
            if not name or not name.strip():
                return
            name = name.strip()
            if name in editors:
                messagebox.showwarning("提示", f"按钮「{name}」已存在", parent=settings_win)
                return
            add_tab(name)
        
        def save_settings():
            new_config = {}
            for button_name, fields in editors.items():
                new_config[button_name] = {
                    '拟办意见': fields['拟办意见'].get(1.0, tk.END).strip(),
                    '办理情况': fields['办理情况'].get(1.0, tk.END).strip()
                }
            if self.save_quick_buttons_config(new_config):
                self.quick_buttons_config = new_config
                self.rebuild_quick_buttons()
                messagebox.showinfo("成功", "设置已保存")
                settings_win.destroy()
        
        def reset_defaults():
            if messagebox.askyesno("确认", "确定要恢复所有按钮的默认设置吗？", parent=settings_win):
                config_file = get_appdata_path('quick_buttons_config.json')
                if os.path.exists(config_file):
                    os.remove(config_file)
                self.quick_buttons_config = self.load_quick_buttons_config()
                self.rebuild_quick_buttons()
                messagebox.showinfo("成功", "已恢复默认设置，请重新打开设置窗口")
                settings_win.destroy()
        
        ttk.Button(bottom_frame, text="➕ 新增按钮", command=add_new_button,
                  width=14, bootstyle="success-outline").pack(side=tk.LEFT, padx=5)
        ttk.Button(bottom_frame, text="取消", command=settings_win.destroy,
                  width=10).pack(side=tk.RIGHT, padx=5)
        ttk.Button(bottom_frame, text="恢复默认", command=reset_defaults,
                  width=12).pack(side=tk.RIGHT, padx=5)
        ttk.Button(bottom_frame, text="保存", command=save_settings,
                  width=10, bootstyle="primary").pack(side=tk.RIGHT, padx=5)
    
    def fill_quick_button(self, button_name):
        """通用填充方法"""
        if button_name not in self.quick_buttons_config:
            self.log(f"未找到配置: {button_name}")
            return
        
        config = self.quick_buttons_config[button_name]
        
        # 填充拟办意见
        widget = self.field_entries.get('承办意见')
        if widget and isinstance(widget, tk.Text):
            widget.delete(1.0, tk.END)
            widget.insert(1.0, config['拟办意见'])
        
        # 填充办理情况
        widget = self.field_entries.get('办理情况')
        if widget and isinstance(widget, tk.Text):
            widget.delete(1.0, tk.END)
            widget.insert(1.0, config['办理情况'])
        
        self.log(f"已填充：{button_name}")

        # 自动保存
        if self.current_file:
            self.save_current_file()
            self.log(f"已自动保存：{button_name}")

    def open_rule_management(self):
        """打开规则管理窗口"""
        try:
            from rule_management_ui import RuleManagementWindow
            RuleManagementWindow(self.root)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开规则管理窗口:\n{e}")
            self.log(f"打开规则管理窗口失败: {e}")

    def fill_yizheng(self):
        """填充医政投诉内容"""
        self.fill_quick_button('医政投诉')
    
    def fill_jikong(self):
        """填充疾控投诉内容"""
        self.fill_quick_button('疾控投诉')
    
    def fill_gongwei(self):
        """填充公卫投诉内容"""
        self.fill_quick_button('公卫投诉')
    
    def fill_zhongyi(self):
        """填充中医投诉内容"""
        self.fill_quick_button('中医投诉')
    
    def fill_fuji(self):
        """填充妇计投诉内容"""
        self.fill_quick_button('妇计投诉')
    
    def fill_yituo(self):
        """填充医托投诉内容"""
        self.fill_quick_button('医托投诉')
    
    def load_previous_file(self):
        """加载上一个文件"""
        if not self.file_list:
            return
        
        selection = self.file_tree.selection()
        if not selection:
            # 如果没有选中，选择第一个
            if self.file_list:
                first_item = self.file_tree.get_children()[0]
                self.file_tree.selection_set(first_item)
                self.file_tree.see(first_item)
            return
        
        # 获取当前选中项的索引
        current_item = selection[0]
        all_items = self.file_tree.get_children()
        current_index = all_items.index(current_item)
        
        # 如果不是第一个，选择上一个
        if current_index > 0:
            prev_item = all_items[current_index - 1]
            self.file_tree.selection_set(prev_item)
            self.file_tree.see(prev_item)
    
    def load_next_file(self):
        """加载下一个文件"""
        if not self.file_list:
            return
        
        selection = self.file_tree.selection()
        if not selection:
            # 如果没有选中，选择第一个
            if self.file_list:
                first_item = self.file_tree.get_children()[0]
                self.file_tree.selection_set(first_item)
                self.file_tree.see(first_item)
            return
        
        # 获取当前选中项的索引
        current_item = selection[0]
        all_items = self.file_tree.get_children()
        current_index = all_items.index(current_item)
        
        # 如果不是最后一个，选择下一个
        if current_index < len(all_items) - 1:
            next_item = all_items[current_index + 1]
            self.file_tree.selection_set(next_item)
            self.file_tree.see(next_item)
    
    def manual_add_document(self):
        """手动添加批办单"""
        # 检查是否有填写内容
        has_content = False
        for field_key, widget in self.field_entries.items():
            if isinstance(widget, tk.Text):
                content = widget.get(1.0, tk.END).strip()
            else:
                content = widget.get().strip() if hasattr(widget, 'get') else ''
            
            if content:
                has_content = True
                break
        
        if not has_content:
            messagebox.showwarning("提示", "请先填写表格内容")
            return
        
        if not messagebox.askyesno("确认", "确定要创建新的批办单吗？\n\n将使用当前表格中填写的内容生成新文件。"):
            return
        
        try:
            # 获取编号（使用当前数据库年份）
            db_year = self.get_current_db_year()
            counter = CounterManager(year=db_year)
            file_num = counter.get_next_file_number()
            receipt_num = counter.get_next_receipt_number()
            
            # 创建新文档
            template_path = '投诉批办单示例#1.docx'
            if not os.path.exists(template_path):
                messagebox.showerror("错误", f"找不到模板文件: {template_path}")
                return
            
            doc = Document(template_path)
            if len(doc.tables) == 0:
                messagebox.showerror("错误", "模板文件没有表格")
                return
            
            table = doc.tables[0]
            generator = WordGenerator(output_dir=self.path_config['待打印'])
            
            # 表格位置映射
            field_mapping = {
                '来文单位': (0, 2),
                '来文字号': (1, 2),
                '收文日期': (1, 9),
                '来文类型': (2, 2),
                '收文途径': (2, 9),
                '收文编号': (3, 9),
                '文件标题': (4, 2),
                '领导批示': (5, 2),
                '承办意见': (6, 2),
                '办理情况': (9, 2),
            }
            
            # 从输入框读取并填充
            for field_name, (row, col) in field_mapping.items():
                if field_name in self.field_entries:
                    try:
                        widget = self.field_entries[field_name]
                        if isinstance(widget, tk.Text):
                            value = widget.get(1.0, tk.END)
                            # 只去掉末尾的换行符，保留前置空格
                            value = value.rstrip('\n\r')
                        elif isinstance(widget, (tk.Entry, ttk.Entry)):
                            value = widget.get()
                        else:
                            value = ''
                        
                        # 特殊处理：自动填充某些字段
                        # 收文编号：始终使用新的编号（防止重复），格式：C-年份后两位-编号
                        if field_name == '收文编号':
                            from datetime import datetime
                            year_suffix = str(datetime.now().year)[-2:]  # 获取年份后两位
                            value = f"C-{year_suffix}-{receipt_num}"
                        elif field_name == '收文日期':
                            # 收文日期：始终使用当前日期（不管输入框中有没有旧日期）
                            from datetime import datetime
                            now = datetime.now()
                            value = f"{now.year}-{now.month}-{now.day}"
                        elif field_name == '来文类型' and not value:
                            value = "C"
                        
                        if row < len(table.rows) and col < len(table.rows[row].cells):
                            cell = table.rows[row].cells[col]
                            
                            # 根据字段应用不同的格式
                            if field_name in ['来文单位', '文件标题', '办理情况']:
                                # 添加前置空格
                                generator.set_cell_text_keep_format(cell, value, add_prefix_space=True)
                            elif field_name in ['来文字号', '收文日期', '来文类型', '收文途径', '收文编号']:
                                # 水平居中
                                generator.set_cell_text_keep_format(cell, value, alignment='center')
                            elif field_name == '承办意见':
                                # 拟办意见：每个段落添加前置空格，固定28磅行距（兼容Word 2007）
                                generator.set_cell_text_keep_format(cell, value, add_prefix_space=True, line_spacing_pt=28)
                            else:
                                # 其他字段（领导批示）保持默认
                                generator.set_cell_text_keep_format(cell, value)
                    except Exception as e:
                        self.log(f"填充字段 {field_name} 失败: {e}")
            
            # 保存文档，文件名格式：投诉批办单#C-年份后两位-编号.docx
            from datetime import datetime
            year_suffix = str(datetime.now().year)[-2:]  # 获取年份后两位
            output_filename = f"投诉批办单#C-{year_suffix}-{file_num}.docx"
            output_path = os.path.join(self.path_config['待打印'], output_filename)
            doc.save(output_path)
            
            # 记录输入历史（来文单位、收文途径、文件标题）
            for field_name in ['来文单位', '收文途径', '文件标题']:
                if field_name in self.field_entries:
                    widget = self.field_entries[field_name]
                    if hasattr(widget, 'get'):
                        value = widget.get()
                        self.input_history.add_value(field_name, value)
            
            # 从生成的批办单Word文档重新提取完整数据，然后录入数据库
            try:
                word_data = WordExtractor.extract_complaint_data(output_path)
                if word_data:
                    word_data['文件名'] = output_filename
                    
                    # 添加到数据库
                    receipt_number = word_data.get('收文编号', '')
                    if self.db.add_complaint(word_data):
                        self.log(f"✓ 已添加到数据库（从批办单提取完整数据）: {receipt_number}")
                    else:
                        self.log(f"✗ 添加到数据库失败: {output_filename}")
                else:
                    self.log(f"✗ 从批办单提取数据失败: {output_filename}")
            except Exception as e:
                self.log(f"从批办单提取数据并录入数据库失败: {e}")
            
            # 更新下一个收文编号显示
            self.update_next_receipt_number()
            
            self.log(f"成功创建: {output_filename}")
            messagebox.showinfo("成功", f"批办单已创建！\n\n文件名: {output_filename}")
            
            # 刷新列表
            self.refresh_file_list()

            # 选中新创建的文件并加载到编辑器
            for item in self.file_tree.get_children():
                values = self.file_tree.item(item)['values']
                if len(values) >= 3 and values[2] == output_filename:
                    self.file_tree.selection_set(item)
                    self.file_tree.see(item)
                    self._load_file_content(output_path, output_filename)
                    break
            
        except Exception as e:
            messagebox.showerror("错误", f"创建批办单失败: {e}")
            self.log(f"创建批办单失败: {e}")
        
    def add_source_files(self):
        """添加待处理文件"""
        # Linux 的 Tk 文件对话框不接受 ";" 分隔的多模式，改为元组形式（Windows 同样支持）
        filetypes = [("PDF/Word文件", ("*.pdf", "*.docx", "*.doc")),
                    ("PDF文件", "*.pdf"),
                    ("Word文件", ("*.docx", "*.doc")),
                    ("所有文件", "*.*")]
        files = filedialog.askopenfilenames(title="选择待处理文件", filetypes=filetypes)
        
        if files:
            for filepath in files:
                if filepath not in self.source_files:
                    self.source_files.append(filepath)
            
            self.refresh_source_list()
            self.log(f"已添加 {len(files)} 个文件")
    
    def add_source_folder(self):
        """添加待处理文件夹"""
        folder = filedialog.askdirectory(title="选择包含待处理文件的文件夹")
        
        if folder:
            count = 0
            for root, dirs, files in os.walk(folder):
                for file in files:
                    if file.lower().endswith(('.pdf', '.docx', '.doc')) and not file.startswith('~$'):
                        filepath = os.path.join(root, file)
                        if filepath not in self.source_files:
                            self.source_files.append(filepath)
                            count += 1
            
            self.refresh_source_list()
            self.log(f"从文件夹添加了 {count} 个文件")
    
    def remove_source_selected(self):
        """移除选中的待处理文件"""
        selection = self.source_tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选择要移除的文件")
            return
        
        # 获取要移除的文件名
        filenames_to_remove = []
        for item in selection:
            values = self.source_tree.item(item)['values']
            if len(values) >= 2:
                filenames_to_remove.append(values[1])
        
        # 从列表中移除
        self.source_files = [f for f in self.source_files 
                            if os.path.basename(f) not in filenames_to_remove]
        
        self.refresh_source_list()
        self.log(f"已移除 {len(filenames_to_remove)} 个文件")
    
    def clear_source_list(self):
        """清空待处理文件列表"""
        if messagebox.askyesno("确认", "确定要清空待处理文件列表吗？"):
            self.source_files = []
            self.refresh_source_list()
            self.log("已清空待处理文件列表")
    
    def rename_docx_to_doc(self):
        """将待处理文件列表中所有.docx文件重命名为.doc"""
        # 筛选出所有.docx文件
        docx_files = [f for f in self.source_files if f.lower().endswith('.docx')]
        
        if not docx_files:
            messagebox.showinfo("提示", "待处理文件列表中没有.docx文件")
            return
        
        # 确认对话框
        if not messagebox.askyesno("确认", f"确定要将 {len(docx_files)} 个.docx文件重命名为.doc吗？\n\n注意：这只会修改文件扩展名，不会转换文件格式。"):
            return
        
        renamed_count = 0
        failed_files = []
        
        for old_path in docx_files:
            try:
                # 构造新文件名（将.docx替换为.doc）
                new_path = old_path[:-5] + '.doc'  # 去掉最后5个字符(.docx)，加上.doc
                
                # 检查新文件名是否已存在
                if os.path.exists(new_path):
                    failed_files.append(f"{os.path.basename(old_path)} (目标文件已存在)")
                    continue
                
                # 重命名文件
                os.rename(old_path, new_path)
                
                # 更新source_files列表
                self.source_files[self.source_files.index(old_path)] = new_path
                
                renamed_count += 1
                
            except Exception as e:
                failed_files.append(f"{os.path.basename(old_path)} ({str(e)})")
        
        # 刷新列表显示
        self.refresh_source_list()
        
        # 显示结果
        if renamed_count > 0:
            self.log(f"成功重命名 {renamed_count} 个文件")
        
        if failed_files:
            error_msg = f"成功重命名 {renamed_count} 个文件\n\n以下 {len(failed_files)} 个文件重命名失败：\n" + "\n".join(failed_files[:10])
            if len(failed_files) > 10:
                error_msg += f"\n... 还有 {len(failed_files) - 10} 个文件"
            messagebox.showwarning("部分失败", error_msg)
        else:
            messagebox.showinfo("完成", f"成功重命名 {renamed_count} 个.docx文件为.doc")
    
    def batch_process(self):
        """批量处理源文件"""
        if not self.source_files:
            messagebox.showwarning("提示", "请先添加待处理文件")
            return
        if self.is_processing:
            messagebox.showwarning("提示", "正在批量处理，请稍后再试")
            return
        
        # 在新线程中运行，避免界面卡顿
        def process_thread():
            self.log(f"开始批量处理 {len(self.source_files)} 个文件...")
            processed_dir = self.path_config['已处理']
            success_count = 0
            
            try:
                # 按文件名升序处理
                sorted_files = sorted(self.source_files, key=lambda x: os.path.basename(x))
                remaining_files = []
                
                for filepath in sorted_files:
                    file_success = False
                    try:
                        filename = os.path.basename(filepath)
                        self.log(f"正在处理: {filename}")
                        
                        # 使用处理器处理单个文件
                        generator = WordGenerator(output_dir=self.path_config['待打印'])
                        
                        # 先提取数据，确认可以处理
                        data = None
                        if filepath.lower().endswith('.pdf'):
                            data = PDFExtractor.extract_table_data(filepath)
                        elif filepath.lower().endswith('.docx'):
                            data = WordExtractor.extract_table_data(filepath)
                        elif filepath.lower().endswith('.doc'):
                            data = DOCExtractor.extract_table_data(filepath)
                        
                        # 只有成功提取数据后才获取编号
                        if data:
                            # 获取编号（使用当前数据库年份）
                            db_year = self.get_current_db_year()
                            counter = CounterManager(year=db_year)
                            file_num = counter.get_next_file_number()
                            receipt_num = counter.get_next_receipt_number()
                            
                            # 生成输出文件路径，格式：投诉批办单#C-年份后两位-编号.docx
                            from datetime import datetime
                            year_suffix = str(datetime.now().year)[-2:]  # 获取年份后两位
                            output_filename = f"投诉批办单#C-{year_suffix}-{file_num}.docx"
                            output_path = os.path.join(self.path_config['待打印'], output_filename)
                            
                            # 根据文件类型生成文档
                            if filepath.lower().endswith('.pdf'):
                                if generator.generate_from_pdf(data, file_num, receipt_num, output_path, filename):
                                    success_count += 1
                                    file_success = True
                                    self.log(f"成功生成: {output_filename}")
                                    # 检查是否匹配关键词
                                    if any(kw in filename for kw in ['医政', '公卫', '中医院', '中医', '疾控', '发展', '妇计', '监督']):
                                        self.log(f"  已根据文件名关键词自动填充内容")
                                    
                                    # 从生成的批办单Word文档中读取完整数据后录入数据库
                                    word_data = WordExtractor.extract_complaint_data(output_path)
                                    if word_data:
                                        word_data['文件名'] = output_filename
                                        if self.db.add_complaint(word_data):
                                            self.log(f"✓ 已录入数据库（完整数据）: {word_data.get('收文编号', '')}, 来文单位={word_data.get('来文单位', '')}, 标题={word_data.get('文件标题', '')}")
                                else:
                                    self.log(f"生成文档失败: {filename}")
                                    # 生成失败，需要回退编号
                                    counter.file_number -= 1
                                    counter.receipt_number -= 1
                                    counter.save()
                            elif filepath.lower().endswith('.docx'):
                                if generator.generate_from_word(data, file_num, receipt_num, output_path, filename):
                                    success_count += 1
                                    file_success = True
                                    self.log(f"成功生成: {output_filename}")
                                    
                                    # 从生成的批办单Word文档中读取完整数据后录入数据库
                                    word_data = WordExtractor.extract_complaint_data(output_path)
                                    if word_data:
                                        word_data['文件名'] = output_filename
                                        if self.db.add_complaint(word_data):
                                            self.log(f"✓ 已录入数据库（完整数据）: {word_data.get('收文编号', '')}, 来文单位={word_data.get('来文单位', '')}, 标题={word_data.get('文件标题', '')}")
                                else:
                                    self.log(f"生成文档失败: {filename}")
                                    # 生成失败，需要回退编号
                                    counter.file_number -= 1
                                    counter.receipt_number -= 1
                                    counter.save()
                            elif filepath.lower().endswith('.doc'):
                                if generator.generate_from_doc(data, file_num, receipt_num, output_path, filename):
                                    success_count += 1
                                    file_success = True
                                    self.log(f"成功生成: {output_filename}")
                                    if any(kw in filename for kw in ['医政', '公卫', '中医院', '中医', '疾控', '发展', '妇计', '监督']):
                                        self.log(f"  已根据文件名关键词自动填充内容")
                                    
                                    # 从生成的批办单Word文档中读取完整数据后录入数据库
                                    word_data = WordExtractor.extract_complaint_data(output_path)
                                    if word_data:
                                        word_data['文件名'] = output_filename
                                        if self.db.add_complaint(word_data):
                                            self.log(f"✓ 已录入数据库（完整数据）: {word_data.get('收文编号', '')}, 来文单位={word_data.get('来文单位', '')}, 标题={word_data.get('文件标题', '')}")
                                else:
                                    self.log(f"生成文档失败: {filename}")
                                    # 生成失败，需要回退编号
                                    counter.file_number -= 1
                                    counter.receipt_number -= 1
                                    counter.save()
                        else:
                            # 数据提取失败
                            self.log(f"提取数据失败: {filename}")
                        
                        # 只有当前文件成功生成文档后才移动该文件
                        if file_success:
                            dest_path = os.path.join(processed_dir, filename)
                            # 如果目标文件已存在，添加时间戳
                            if os.path.exists(dest_path):
                                name, ext = os.path.splitext(filename)
                                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                                dest_path = os.path.join(processed_dir, f"{name}_{timestamp}{ext}")
                            
                            import shutil
                            shutil.move(filepath, dest_path)
                            self.log(f"已移动到已处理: {filename}")
                            
                            # 立即刷新待打印列表，显示新生成的文件
                            self.root.after(0, lambda: self.refresh_file_list())
                            # 立即更新待处理列表，移除已处理的文件
                            self.source_files = [f for f in self.source_files if f != filepath]
                            self.root.after(0, lambda: self.refresh_source_list())
                        else:
                            # 保留处理失败的文件在待处理列表中
                            remaining_files.append(filepath)
                        
                    except Exception as e:
                        self.log(f"处理文件失败 {filename}: {e}")
                        import traceback
                        self.log(traceback.format_exc())
                
                # 更新待处理列表，仅保留失败项（已在循环中实时更新，这里只需同步）
                failed_count = len(remaining_files)
                self.source_files = remaining_files
                
                # 最后再刷新一次，确保界面状态正确
                self.root.after(0, lambda: self.refresh_all_lists())
                # 更新下一个收文编号显示
                self.root.after(0, lambda: self.update_next_receipt_number())
                self.log(f"批量处理完成，成功处理 {success_count} 个文件，失败 {failed_count} 个")
                self.root.after(0, lambda: messagebox.showinfo("完成", f"批量处理完成！\n成功: {success_count} 个\n失败: {failed_count} 个\n失败的文件已保留在待处理列表中"))
            except Exception as e:
                self.log(f"批量处理失败: {e}")
                self.root.after(0, lambda: messagebox.showerror("错误", f"批量处理失败: {e}"))
            finally:
                # 释放Word实例（批量处理完成后）
                DOCExtractor.release_word_instance()
                self.is_processing = False
        
        self.is_processing = True
        threading.Thread(target=process_thread, daemon=True).start()
        
    def print_preview(self):
        """打印预览"""
        self.log("点击了预览按钮")
        self.log(f"当前文件: {self.current_file}")
        
        if not self.current_file:
            messagebox.showwarning("提示", "没有打开的文件")
            return
        
        # 确保使用绝对路径
        if not os.path.isabs(self.current_file):
            abs_path = os.path.abspath(self.current_file)
            self.log(f"转换为绝对路径: {abs_path}")
        else:
            abs_path = self.current_file
        
        if not os.path.exists(abs_path):
            messagebox.showerror("错误", f"文件不存在: {abs_path}")
            self.log(f"文件不存在: {abs_path}")
            return
        
        try:
            # 预览前添加二维码（如果启用）
            if self.qrcode_config.get('enabled', False):
                self.log("预览前添加二维码...")
                self.add_qrcode_before_print(abs_path)
            
            # 用系统默认程序打开（Windows 等价 cmd /c start ""，Linux 为 xdg-open/WPS）
            self.log(f"正在打开文件: {abs_path}")
            if not open_path(abs_path):
                raise RuntimeError("未找到可用的系统打开命令（xdg-open / gio / open）")
            self.log(f"✓ 已打开预览: {os.path.basename(abs_path)}")
            
        except Exception as e:
            messagebox.showerror("错误", f"打开预览失败: {e}")
            self.log(f"✗ 打开预览失败: {e}")
            import traceback
            self.log(traceback.format_exc())
    
    def print_current(self):
        """打印当前文件"""
        if not self.current_file:
            return
        
        try:
            import shutil
            
            # 保存当前文件路径（因为清空表单会将 self.current_file 设为 None）
            current_file_path = self.current_file
            filename = os.path.basename(current_file_path)
            
            self.log(f"开始打印流程: {filename}")
            self.log(f"文件路径: {current_file_path}")
            
            # 打印前添加二维码（如果启用）
            if self.qrcode_config.get('enabled', False):
                self.log("打印前添加二维码...")
                self.add_qrcode_before_print(current_file_path)
            
            # 使用系统打印通道（Windows=ShellExecute，UOS/麒麟=LibreOffice + CUPS lp）
            ok, info = print_file(current_file_path)
            if not ok:
                raise RuntimeError(info)
            self.log(f"已发送打印: {filename} ({info})")
            
            # 立即复制到已打印目录（覆盖同名文件）
            printed_dir = self.path_config['已打印']
            dest_path = os.path.join(printed_dir, filename)
            
            # 直接覆盖，不添加时间戳
            shutil.copy2(current_file_path, dest_path)
            self.log(f"文件已复制到已打印目录: {filename}")
            
            # 打印完成后清空表单
            self._clear_form_silently()
            self.log("打印完成，已清空表单")
            
            # 使用重试机制删除文件
            def delete_with_retry(retry_count=0):
                # 从配置文件读取延迟设置，如果没有则使用默认值
                try:
                    with open(pc.get_user_data_path('print_delay_config.json'), 'r', encoding='utf-8') as f:
                        delay_config = json.load(f)
                        max_retries = delay_config.get('max_retries', 5)  # 默认最多重试5次
                        initial_delay = delay_config.get('initial_delay_seconds', 15) * 1000  # 默认初始延迟15秒
                        retry_delay = delay_config.get('retry_delay_seconds', 8) * 1000  # 默认重试间隔8秒
                except:
                    # 配置文件不存在或读取失败，使用默认值
                    max_retries = 5  # 最多重试5次
                    initial_delay = 15000  # 初始延迟15秒
                    retry_delay = 8000  # 每次重试间隔8秒
                
                current_delay = initial_delay if retry_count == 0 else retry_delay
                
                self.log(f"尝试删除文件 (第{retry_count + 1}次): {filename}")
                
                try:
                    if os.path.exists(current_file_path):
                        os.remove(current_file_path)
                        self.log(f"已删除待打印文件: {filename}")
                        # 删除成功，刷新列表
                        self.refresh_file_list()
                        self.log("文件列表刷新完成")
                        return
                    else:
                        self.log(f"文件不存在，无法删除: {filename}")
                        return
                        
                except PermissionError as e:
                    if retry_count < max_retries:
                        next_delay = retry_delay
                        self.log(f"文件被占用，{next_delay/1000}秒后重试 (第{retry_count + 1}/{max_retries + 1}次): {e}")
                        # 递归重试
                        self.root.after(next_delay, lambda: delete_with_retry(retry_count + 1))
                    else:
                        self.log(f"删除文件失败，已达到最大重试次数 ({max_retries + 1}次): {e}")
                        self.log("建议手动删除文件或稍后再试")
                        self.log("如需调整延迟时间，请编辑 AppData/print_delay_config.json 文件")
                        
                except Exception as e:
                    self.log(f"删除文件失败: {e}")
            
            # 读取初始延迟配置
            try:
                with open(pc.get_user_data_path('print_delay_config.json'), 'r', encoding='utf-8') as f:
                    delay_config = json.load(f)
                    initial_delay = delay_config.get('initial_delay_seconds', 15) * 1000
            except:
                initial_delay = 15000  # 默认15秒
            
            self.log(f"设置{initial_delay/1000}秒延迟删除任务（带重试机制）")
            self.root.after(initial_delay, delete_with_retry)
            
        except Exception as e:
            messagebox.showerror("错误", f"打印失败: {e}")
            self.log(f"打印失败: {e}")
        
    def convert_word_to_pdf(self, word_path):
        """将Word文档转换为PDF（Windows 用 docx2pdf，UOS/麒麟用 LibreOffice/WPS）

        Args:
            word_path: Word文档路径
        Returns:
            转换成功的PDF路径，失败返回None
        """
        try:
            import time
            from PyPDF2 import PdfReader

            word_path = os.path.abspath(word_path)
            temp_dir = self.path_config['待打印']
            base_name = os.path.splitext(os.path.basename(word_path))[0]
            timestamp = int(time.time() * 1000)
            temp_pdf = os.path.abspath(os.path.join(temp_dir, f"temp_{timestamp}_{base_name}.pdf"))

            self.log(f"正在转换Word为PDF: {os.path.basename(word_path)}")
            self.log(f"源文件: {word_path}")
            self.log(f"目标文件: {temp_pdf}")

            for _ in range(3):
                produced = word_to_pdf(word_path, temp_pdf)
                if produced and os.path.isfile(produced) and os.path.getsize(produced) > 0:
                    try:
                        reader = PdfReader(produced)
                        if len(reader.pages) > 0:
                            self.log(f"转换完成: {os.path.basename(produced)} ({os.path.getsize(produced)} bytes)")
                            return produced
                    except Exception as e:
                        self.log(str(e))
                time.sleep(0.6)

            self.log("PDF文件未生成或为空（请确认已安装 LibreOffice 或 WPS）")
            return None

        except Exception as e:
            self.log(f"Word转PDF失败 {os.path.basename(word_path)}: {e}")
            self.log(f"错误类型: {type(e).__name__}")
            import traceback
            self.log(f"详细错误: {traceback.format_exc()}")
            return None
    
    def merge_pdf_documents(self, pdf_list, output_path):
        """合并多个PDF文档为一个PDF
        
        Args:
            pdf_list: PDF文件路径列表
            output_path: 输出PDF路径
            
        Returns:
            成功返回True，失败返回False
        """
        try:
            # PyPDF2 >= 3 已移除 PdfMerger，统一使用 PdfWriter（Windows/Linux 行为一致）
            from PyPDF2 import PdfWriter
            
            if not pdf_list:
                self.log("没有要合并的PDF文档")
                return False
            
            self.log(f"开始合并 {len(pdf_list)} 个PDF文档...")
            
            merger = PdfWriter()
            
            for idx, pdf_path in enumerate(pdf_list, 1):
                try:
                    self.log(f"正在添加第 {idx}/{len(pdf_list)} 个PDF: {os.path.basename(pdf_path)}")
                    merger.append(pdf_path)
                except Exception as e:
                    self.log(f"添加PDF失败 {os.path.basename(pdf_path)}: {e}")
                    continue
            
            # 保存合并后的PDF
            with open(output_path, 'wb') as out_f:
                merger.write(out_f)
            merger.close()
            
            self.log(f"PDF合并完成，已保存到: {output_path}")
            return True
            
        except Exception as e:
            self.log(f"合并PDF失败: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def add_qrcode_before_print(self, filepath):
        """
        打印前添加二维码到文档
        只有启用二维码功能时才执行
        
        参数:
            filepath: Word文档路径
        
        返回:
            True: 成功添加或无需添加
            False: 添加失败
        """
        # 检查是否启用二维码功能
        if not self.qrcode_config.get('enabled', False):
            self.log(f"  二维码功能未启用，跳过")
            return True  # 未启用，直接返回成功
        
        try:
            from docx import Document
            
            self.log(f"  正在添加二维码...")
            
            # 打开文档
            doc = Document(filepath)
            
            # 从文档中提取数据用于二维码
            qr_data = {}
            fields = self.qrcode_config.get('fields', {})
            
            self.log(f"  配置的字段: {fields}")
            
            # 从文档表格中读取数据
            if len(doc.tables) > 0:
                table = doc.tables[0]
                
                if fields.get('收文编号', False) and len(table.rows) > 3:
                    receipt_cell = table.rows[3].cells[9] if len(table.rows[3].cells) > 9 else None
                    if receipt_cell:
                        qr_data['收文编号'] = receipt_cell.text.strip()
                
                if fields.get('收文日期', False) and len(table.rows) > 1:
                    date_cell = table.rows[1].cells[9] if len(table.rows[1].cells) > 9 else None
                    if date_cell:
                        qr_data['收文日期'] = date_cell.text.strip()
                
                if fields.get('来文字号', False) and len(table.rows) > 1:
                    ref_cell = table.rows[1].cells[2] if len(table.rows[1].cells) > 2 else None
                    if ref_cell:
                        qr_data['来文字号'] = ref_cell.text.strip()
                
                if fields.get('来文单位', False) and len(table.rows) > 0:
                    unit_cell = table.rows[0].cells[2] if len(table.rows[0].cells) > 2 else None
                    if unit_cell:
                        qr_data['来文单位'] = unit_cell.text.strip()
                
                if fields.get('收文途径', False) and len(table.rows) > 2:
                    way_cell = table.rows[2].cells[9] if len(table.rows[2].cells) > 9 else None
                    if way_cell:
                        qr_data['收文途径'] = way_cell.text.strip()
                
                if fields.get('文件标题', False) and len(table.rows) > 4:
                    title_cell = table.rows[4].cells[2] if len(table.rows[4].cells) > 2 else None
                    if title_cell:
                        qr_data['文件标题'] = title_cell.text.strip()
            
            self.log(f"  提取的数据: {qr_data}")
            
            # 如果有数据，添加二维码
            if qr_data:
                size = self.qrcode_config.get('size', 'small')
                self.log(f"  二维码尺寸: {size}")
                
                # ===== 先删除已存在的所有二维码段落 =====
                body = doc._element.body
                qr_removed_count = 0
                
                # 遍历所有元素，找出并删除二维码段落
                # 注意：要从后往前删除，避免索引变化
                elements_to_remove = []
                
                for i, element in enumerate(body):
                    if not element.tag.endswith('}p'):
                        continue  # 只检查段落
                    
                    is_qr_paragraph = False
                    
                    # 检查spacing属性: line="1" lineRule="exact"
                    for child in element:
                        if child.tag.endswith('}pPr'):
                            for ppr_child in child:
                                if ppr_child.tag.endswith('}spacing'):
                                    ns_w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
                                    line_val = ppr_child.get(f'{ns_w}line')
                                    line_rule = ppr_child.get(f'{ns_w}lineRule')
                                    if line_val == '1' and line_rule == 'exact':
                                        # 进一步确认：检查是否有anchor
                                        has_anchor = False
                                        for child2 in element:
                                            if child2.tag.endswith('}r'):
                                                for run_child in child2:
                                                    if run_child.tag.endswith('}drawing'):
                                                        for draw_child in run_child:
                                                            if draw_child.tag.endswith('}anchor'):
                                                                has_anchor = True
                                                                break
                                        if has_anchor:
                                            is_qr_paragraph = True
                                            elements_to_remove.append(element)
                                        break
                            break
                
                # 删除所有二维码段落
                for element in elements_to_remove:
                    body.remove(element)
                    qr_removed_count += 1
                
                if qr_removed_count > 0:
                    self.log(f"  ✓ 已删除 {qr_removed_count} 个旧二维码段落")
                else:
                    self.log(f"  未检测到旧二维码")
                # ===== 删除逻辑结束 =====
                
                # 使用 WordGenerator 的静态方法添加新二维码
                try:
                    from main import WordGenerator
                    WordGenerator.add_qrcode_to_document(doc, qr_data, size)
                except Exception as add_err:
                    self.log(f"  ✗ 添加二维码失败: {add_err}")
                    import traceback
                    traceback.print_exc()
                    return False
                
                # 保存文档
                doc.save(filepath)
                self.log(f"  ✓ 已添加二维码并保存: {os.path.basename(filepath)}")
                return True
            else:
                self.log(f"  未选择二维码字段或无数据，跳过: {os.path.basename(filepath)}")
                return True
                
        except Exception as e:
            self.log(f"  ✗ 添加二维码失败 {os.path.basename(filepath)}: {e}")
            import traceback
            self.log(traceback.format_exc())
            return False  # 添加失败，但不影响打印
    
    def silent_batch_print(self):
        """批量打印：静默打印勾选的Word文档，每个文件间隔5秒"""
        if self.is_printing:
            messagebox.showwarning("提示", "正在批量打印，请稍后再试")
            return
        
        # 使用复选框勾选的项
        checked_items = [item for item, checked in self.file_checkboxes.items() if checked]
        if not checked_items:
            messagebox.showwarning("提示", "请先在文件名前的方框内勾选要打印的批办单")
            return
        
        try:
            import time
            import shutil
            self.is_printing = True
            
            # 获取勾选的文件路径列表（按列表顺序）
            files_to_print = []
            for item in checked_items:
                values = self.file_tree.item(item)['values']
                if len(values) >= 3:  # 有3列：选择、序号、文件名
                    filename = values[2]  # 文件名在第3列（索引2）
                    filepath = os.path.join(self.path_config['待打印'], filename)
                    if os.path.exists(filepath) and filename.endswith('.docx'):
                        files_to_print.append((filepath, filename))
            
            if not files_to_print:
                messagebox.showwarning("提示", "没有可打印的Word文档")
                self.is_printing = False
                return
            
            # 确认打印
            estimated_time = len(files_to_print) * 5 + 20  # 间隔时间 + 打印机处理时间
            result = messagebox.askyesno(
                "确认打印",
                f"将要打印 {len(files_to_print)} 个文档\n"
                f"每个文档间隔 5 秒\n"
                f"打印完成后等待 20 秒确保打印机处理\n"
                f"预计总时间: {estimated_time} 秒\n\n"
                f"是否继续？"
            )
            
            if not result:
                self.is_printing = False
                return
            
            self.log(f"开始批量打印 {len(files_to_print)} 个文档...")
            
            # 逐个打印文档
            success_count = 0
            fail_count = 0
            printed_files = []  # 记录成功打印的文件
            
            for i, (filepath, filename) in enumerate(files_to_print, 1):
                try:
                    self.log(f"[{i}/{len(files_to_print)}] 正在处理: {filename}")
                    
                    # 打印前添加二维码（如果启用）
                    if self.qrcode_config.get('enabled', False):
                        self.add_qrcode_before_print(filepath)
                        # 添加短暂延迟确保文件写入完成
                        time.sleep(0.5)
                    
                    # 系统打印通道：Windows=ShellExecute，UOS/麒麟=LO + CUPS lp
                    ok, info = print_file(filepath)
                    if not ok:
                        raise RuntimeError(info)

                    success_count += 1
                    printed_files.append((filepath, filename))
                    self.log(f"✓ 已发送到打印机: {filename} ({info})")
                    
                    # 等待5秒后打印下一个（最后一个不等待）
                    if i < len(files_to_print):
                        self.log(f"等待5秒后打印下一个...")
                        time.sleep(5)
                    
                except Exception as e:
                    self.log(f"✗ 打印失败 {filename}: {e}")
                    fail_count += 1
            
            # 等待打印机处理所有文件（避免文件被过早删除）
            if printed_files:
                wait_time = 20  # 等待20秒让打印机处理文件
                self.log(f"所有文件已发送到打印队列，等待 {wait_time} 秒确保打印机读取文件...")
                time.sleep(wait_time)
                self.log("等待完成，开始移动文件...")
            
            # 将成功打印的文件移至已打印目录（覆盖同名文件）
            printed_dir = self.path_config['已打印']
            for filepath, filename in printed_files:
                try:
                    dest_path = os.path.join(printed_dir, filename)
                    
                    # 如果目标文件已存在，先删除（实现覆盖）
                    if os.path.exists(dest_path):
                        os.remove(dest_path)
                    
                    # 复制文件到已打印目录
                    shutil.copy2(filepath, dest_path)
                    
                    # 删除待打印目录中的文件
                    os.remove(filepath)
                    self.log(f"已移至已打印目录: {filename}")
                except Exception as e:
                    self.log(f"移动文件失败 {filename}: {e}")
            
            # 刷新文件列表
            self.refresh_file_list()
            
            # 清空撤销合并备份并禁用按钮（打印后不再需要撤销）
            if self.merge_backup:
                self.merge_backup = None
                if self.undo_merge_button:
                    self.undo_merge_button.config(state='disabled')
                self.log("✓ 已清空撤销合并备份（打印后不再需要撤销）")
            
            # 显示结果
            self.log(f"批量打印完成！成功: {success_count} 个，失败: {fail_count} 个")
            messagebox.showinfo(
                "打印完成",
                f"批量打印完成！\n\n"
                f"✓ 成功发送到打印机: {success_count} 个\n"
                f"✗ 失败: {fail_count} 个\n\n"
                f"已将成功打印的文件移至'已打印'文件夹"
            )
            
        except Exception as e:
            messagebox.showerror("错误", f"批量打印失败: {e}")
            self.log(f"批量打印失败: {e}")
            import traceback
            traceback.print_exc()
        finally:
            self.is_printing = False
    
    def batch_convert_to_pdf(self):
        """批量转pdf：将勾选的文件转换为PDF并合并"""
        if self.is_printing:
            messagebox.showwarning("提示", "正在批量打印，请稍后再试")
            return
        # 使用复选框勾选的项，而不是TreeView的选中高亮
        checked_items = [item for item, checked in self.file_checkboxes.items() if checked]
        if not checked_items:
            messagebox.showwarning("提示", "请先在文件名前的方框内勾选要打印的批办单")
            return
        
        try:
            import shutil
            import time
            self.is_printing = True
            
            # 获取勾选的文件路径列表
            files_to_print = []
            for item in checked_items:
                values = self.file_tree.item(item)['values']
                if len(values) >= 3:  # 现在有3列：选择、序号、文件名
                    filename = values[2]  # 文件名在第3列（索引2）
                    filepath = os.path.join(self.path_config['待打印'], filename)
                    if os.path.exists(filepath):
                        files_to_print.append(filepath)
            
            if not files_to_print:
                return
            
            self.log(f"准备批量打印 {len(files_to_print)} 个文件")
            
            # 步骤2: 将所有Word文档转换为PDF
            import tempfile
            pdf_files = []
            temp_files = []  # 保存临时文件列表用于后续清理
            successful_word_files = []  # 记录转换成功的Word文件
            
            failed_count = 0
            for filepath in files_to_print:
                pdf_path = self.convert_word_to_pdf(filepath)
                if pdf_path:
                    pdf_files.append(pdf_path)
                    temp_files.append(pdf_path)
                    successful_word_files.append(filepath)
                else:
                    self.log(f"转换PDF失败，跳过文件: {os.path.basename(filepath)}")
                    failed_count += 1
            
            if not pdf_files:
                messagebox.showerror("错误", "没有成功转换的PDF文件")
                return
            
            # 步骤3: 合并所有PDF
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            merged_pdf_filename = f"批量打印合并_{timestamp}.pdf"
            merged_pdf_path = os.path.join(self.path_config['待打印'], merged_pdf_filename)
            
            if not self.merge_pdf_documents(pdf_files, merged_pdf_path):
                # 清理临时PDF文件
                for temp_file in temp_files:
                    try:
                        if os.path.exists(temp_file):
                            os.remove(temp_file)
                    except:
                        pass
                messagebox.showerror("错误", "合并PDF失败")
                return
            
            # 步骤3.1: 复制成功转换的文件到已打印目录（覆盖同名文件）
            printed_dir = self.path_config['已打印']
            for filepath in successful_word_files:
                try:
                    filename = os.path.basename(filepath)
                    dest_path = os.path.join(printed_dir, filename)
                    # 直接覆盖，不添加时间戳
                    shutil.copy2(filepath, dest_path)
                    self.log(f"已复制到已打印目录: {filename}")
                except Exception as e:
                    self.log(f"复制文件失败 {filename}: {e}")

            # 删除待打印列表的Word文件和临时PDF文件
            for filepath in successful_word_files:
                try:
                    if os.path.exists(filepath):
                        os.remove(filepath)
                        self.log(f"已删除: {os.path.basename(filepath)}")
                except Exception as e:
                    self.log(f"删除文件失败 {os.path.basename(filepath)}: {e}")
            
            for temp_file in temp_files:
                try:
                    if os.path.exists(temp_file):
                        os.remove(temp_file)
                        self.log(f"已删除临时PDF: {os.path.basename(temp_file)}")
                except Exception as e:
                    self.log(f"删除临时PDF失败: {e}")
            
            # 刷新文件列表
            self.refresh_file_list()
            
            # 步骤4: 使用默认PDF软件打开合并的PDF，让用户手动打印
            self.log(f"正在打开PDF文件: {merged_pdf_filename}")
            try:
                open_path(merged_pdf_path)
                self.log("PDF文件已打开，请手动完成打印")
                messagebox.showinfo("提示", 
                    f"已将 {len(successful_word_files)} 个文档转换并合并为PDF\n"
                    f"未能转换的文档: {failed_count} 个（已保留在待打印列表）\n\n"
                    f"✅ 已成功文件已备份到'已打印'文件夹\n"
                    f"✅ 已删除成功转换的待打印文件\n"
                    f"✅ PDF文件已打开\n\n"
                    f"请在PDF阅读器中手动完成打印操作\n"
                    f"打印完成后，临时PDF文件将在下次启动程序时自动清理")
                
            except Exception as e:
                self.log(f"打开PDF失败: {e}")
                messagebox.showerror("错误", f"打开PDF失败: {e}\n\n请手动打开文件: {merged_pdf_path}")
            finally:
                self.is_printing = False
                
        except Exception as e:
            messagebox.showerror("错误", f"批量打印失败: {e}")
            self.log(f"批量打印失败: {e}")
            import traceback
            traceback.print_exc()
        
    def batch_print_all(self):
        """批量打印全部文件（合并后打印）"""
        if not self.file_list:
            return
        if self.is_printing:
            messagebox.showwarning("提示", "正在批量打印，请稍后再试")
            return
        
        try:
            import shutil
            import time
            self.is_printing = True
            
            # 获取所有文件路径列表（已按编号排序）
            files_to_print = []
            for filename in self.file_list:
                filepath = os.path.join(self.path_config['待打印'], filename)
                if os.path.exists(filepath):
                    files_to_print.append(filepath)
            
            if not files_to_print:
                return
            
            self.log(f"准备批量打印全部 {len(files_to_print)} 个文件")
            
            # 步骤2: 将所有Word文档转换为PDF
            import tempfile
            pdf_files = []
            temp_files = []  # 保存临时文件列表用于后续清理
            successful_word_files = []  # 记录转换成功的Word文件
            
            failed_count = 0
            for filepath in files_to_print:
                pdf_path = self.convert_word_to_pdf(filepath)
                if pdf_path:
                    pdf_files.append(pdf_path)
                    temp_files.append(pdf_path)
                    successful_word_files.append(filepath)
                else:
                    self.log(f"转换PDF失败，跳过文件: {os.path.basename(filepath)}")
                    failed_count += 1
            
            if not pdf_files:
                messagebox.showerror("错误", "没有成功转换的PDF文件")
                return
            
            # 步骤3: 合并所有PDF
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            merged_pdf_filename = f"批量打印合并_{timestamp}.pdf"
            merged_pdf_path = os.path.join(self.path_config['待打印'], merged_pdf_filename)
            
            if not self.merge_pdf_documents(pdf_files, merged_pdf_path):
                # 清理临时PDF文件
                for temp_file in temp_files:
                    try:
                        if os.path.exists(temp_file):
                            os.remove(temp_file)
                    except:
                        pass
                messagebox.showerror("错误", "合并PDF失败")
                return
            
            # 步骤3.1: 复制成功转换的文件到已打印目录（覆盖同名文件）
            printed_dir = self.path_config['已打印']
            for filepath in successful_word_files:
                try:
                    filename = os.path.basename(filepath)
                    dest_path = os.path.join(printed_dir, filename)
                    # 直接覆盖，不添加时间戳
                    shutil.copy2(filepath, dest_path)
                    self.log(f"已复制到已打印目录: {filename}")
                except Exception as e:
                    self.log(f"复制文件失败 {filename}: {e}")

            # 删除待打印列表的Word文件和临时PDF文件
            for filepath in successful_word_files:
                try:
                    if os.path.exists(filepath):
                        os.remove(filepath)
                        self.log(f"已删除: {os.path.basename(filepath)}")
                except Exception as e:
                    self.log(f"删除文件失败 {os.path.basename(filepath)}: {e}")
            
            for temp_file in temp_files:
                try:
                    if os.path.exists(temp_file):
                        os.remove(temp_file)
                        self.log(f"已删除临时PDF: {os.path.basename(temp_file)}")
                except Exception as e:
                    self.log(f"删除临时PDF失败: {e}")
            
            # 刷新文件列表
            self.refresh_file_list()
            
            # 步骤4: 使用默认PDF软件打开合并的PDF，让用户手动打印
            self.log(f"正在打开PDF文件: {merged_pdf_filename}")
            try:
                open_path(merged_pdf_path)
                self.log("PDF文件已打开，请手动完成打印")
                
            except Exception as e:
                self.log(f"打开PDF失败: {e}")
                messagebox.showerror("错误", f"打开PDF失败: {e}\n\n请手动打开文件: {merged_pdf_path}")
            else:
                messagebox.showinfo("提示", 
                    f"已将 {len(successful_word_files)} 个文档转换并合并为PDF\n"
                    f"未能转换的文档: {failed_count} 个（已保留在待打印列表）")
                
                # 批量打印完成后自动清空表单
                self._clear_form_silently()
                self.log("批量打印完成，已自动清空表单")
            finally:
                self.is_printing = False
                
        except Exception as e:
            messagebox.showerror("错误", f"批量打印失败: {e}")
            self.log(f"批量打印失败: {e}")
            import traceback
            traceback.print_exc()
    
    def move_to_printed(self, filepath):
        """移动文件到已打印目录（覆盖同名文件）"""
        try:
            import shutil
            filename = os.path.basename(filepath)
            dest_dir = self.path_config['已打印']
            dest_path = os.path.join(dest_dir, filename)
            
            # 如果目标文件已存在，先删除（实现覆盖）
            if os.path.exists(dest_path):
                os.remove(dest_path)
                self.log(f"已删除旧文件: {filename}")
            
            # 移动文件，不添加时间戳
            shutil.move(filepath, dest_path)
            self.log(f"已移动到已打印: {filename}")
        except Exception as e:
            self.log(f"移动文件失败 {filename}: {e}")
    
    def move_selected_to_printed(self):
        """手动将选中文件移至已打印目录"""
        selection = self.file_tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选择要移动的文件")
            return
        
        if not messagebox.askyesno("确认", f"确定要将选中的 {len(selection)} 个文件移至已打印文件夹吗？"):
            return
        
        count = 0
        moved_files = []
        for item in selection:
            values = self.file_tree.item(item)['values']
            if len(values) >= 3:  # 现在有3列：选择、序号、文件名
                filename = values[2]  # 文件名在第3列（索引2）
                filepath = os.path.join(self.path_config['待打印'], filename)
                try:
                    self.move_to_printed(filepath)
                    count += 1
                    moved_files.append(filepath)
                except Exception as e:
                    self.log(f"移动失败 {filename}: {e}")
        
        self.refresh_file_list()
        
        # 如果移动的是当前编辑的文件，清空表单
        if self.current_file and self.current_file in moved_files:
            self._clear_form_silently()
        
        # 如果待打印列表已清空，清空撤销合并备份并禁用按钮
        if not self.file_list:
            if self.merge_backup:
                self.merge_backup = None
                if self.undo_merge_button:
                    self.undo_merge_button.config(state='disabled')
                self.log("✓ 已清空撤销合并备份（待打印列表已清空）")
        
        messagebox.showinfo("完成", f"已移动 {count} 个文件到已打印目录")
        
    def delete_selected(self):
        """删除选中的文件"""
        selection = self.file_tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选择要删除的文件")
            return
        
        if not messagebox.askyesno("确认", f"确定要删除选中的 {len(selection)} 个文件吗？"):
            return
        
        count = 0
        deleted_files = []
        for item in selection:
            values = self.file_tree.item(item)['values']
            if len(values) >= 3:  # 现在有3列：选择、序号、文件名
                filename = values[2]  # 文件名在第3列（索引2）
                filepath = os.path.join(self.path_config['待打印'], filename)
                try:
                    os.remove(filepath)
                    count += 1
                    deleted_files.append(filepath)
                    self.log(f"已删除文件: {filename}")
                except Exception as e:
                    self.log(f"删除失败 {filename}: {e}")
        
        self.refresh_file_list()
        
        # 如果删除的是当前编辑的文件，清空表单
        if self.current_file and self.current_file in deleted_files:
            self._clear_form_silently()
        
        # 如果待打印列表已清空，清空撤销合并备份并禁用按钮
        if not self.file_list:
            if self.merge_backup:
                self.merge_backup = None
                if self.undo_merge_button:
                    self.undo_merge_button.config(state='disabled')
                self.log("✓ 已清空撤销合并备份（待打印列表已清空）")
        
        messagebox.showinfo("完成", f"已删除 {count} 个文件")
    
    def on_tree_click(self, event):
        """处理树形视图的点击事件，切换复选框状态"""
        region = self.file_tree.identify_region(event.x, event.y)
        if region != 'cell':
            return
        
        # 获取点击的列
        column = self.file_tree.identify_column(event.x)
        
        # 如果点击的是第一列（复选框列，#1表示第一列）
        if column == '#1':
            item = self.file_tree.identify_row(event.y)
            if item:
                # 切换复选框状态
                current_state = self.file_checkboxes.get(item, False)
                new_state = not current_state
                self.file_checkboxes[item] = new_state
                
                # 更新显示
                values = list(self.file_tree.item(item)['values'])
                values[0] = '☑' if new_state else '☐'
                self.file_tree.item(item, values=values)
    
    def toggle_all_checkboxes(self):
        """全选/全不选复选框"""
        # 检查是否所有项都已选中
        all_checked = all(self.file_checkboxes.values()) if self.file_checkboxes else False
        
        # 切换到相反状态
        new_state = not all_checked
        
        # 更新所有项
        for item in self.file_tree.get_children():
            self.file_checkboxes[item] = new_state
            values = list(self.file_tree.item(item)['values'])
            values[0] = '☑' if new_state else '☐'
            self.file_tree.item(item, values=values)
        
        # 更新表头
        self.file_tree.heading('选择', text='☑' if new_state else '☐')
        
        self.log(f"已{'全选' if new_state else '取消全选'}批办单")
    
    def merge_selected_documents(self):
        """合并选中的批办单"""
        try:
            # 获取所有选中的项
            checked_items = [item for item, checked in self.file_checkboxes.items() if checked]
            
            if len(checked_items) < 2:
                messagebox.showwarning("提示", "请至少选择2个批办单进行合并")
                return
            
            # 获取选中项的文件名（按序号排序）
            selected_files = []
            for item in checked_items:
                values = self.file_tree.item(item)['values']
                if len(values) >= 3:
                    filename = values[2]  # 文件名在第3列
                    selected_files.append(filename)
            
            if len(selected_files) < 2:
                messagebox.showwarning("提示", "请至少选择2个批办单进行合并")
                return
            
            # 确认合并操作
            file_list = '\n'.join([f"  {i+1}. {f}" for i, f in enumerate(selected_files)])
            confirm_msg = f"即将合并以下 {len(selected_files)} 个批办单：\n\n{file_list}\n\n" \
                         f"第一个批办单将保留，其他批办单的'来文字号'将合并到第一个批办单中，\n" \
                         f"然后删除其他批办单。\n\n是否继续？"
            
            if not messagebox.askyesno("确认合并", confirm_msg):
                return
            
            self.log(f"开始合并 {len(selected_files)} 个批办单...")
            
            # ========== 保存合并前的状态（用于撤销） ==========
            self.log("保存合并前状态...")
            output_dir = self.path_config['待打印']
            
            # 保存所有待合并文件的完整内容和数据库记录
            backup_files = {}
            backup_db_records = {}
            
            for filename in selected_files:
                filepath = os.path.join(output_dir, filename)
                try:
                    # 读取文件内容（二进制）
                    with open(filepath, 'rb') as f:
                        file_content = f.read()
                    backup_files[filename] = file_content
                    
                    # 读取数据库记录（支持新旧格式收文编号）
                    import re as _re
                    _ym = _re.search(r'#C-(\d{2})-(\d+)', filename)
                    if _ym:
                        receipt_number = f"C-{_ym.group(1)}-{_ym.group(2)}"
                    else:
                        receipt_number = f"C-{self.extract_number_from_filename(filename)}"
                    db_record = self.db.get_complaint(receipt_number)
                    if db_record:
                        backup_db_records[receipt_number] = db_record
                    
                    self.log(f"  已备份: {filename}")
                except Exception as e:
                    self.log(f"  备份失败 {filename}: {e}")
            
            # 保存当前的收文编号计数器（使用当前数据库对应的计数器）
            db_year = self.get_current_db_year()
            current_counter = CounterManager(year=db_year)
            backup_counter = {
                'file_number': current_counter.file_number,
                'receipt_number': current_counter.receipt_number
            }
            
            # 保存所有当前文件列表（用于恢复重新编号）
            all_current_files = {}
            for f in os.listdir(output_dir):
                if f.endswith('.docx') and not f.startswith('~$'):
                    fpath = os.path.join(output_dir, f)
                    try:
                        with open(fpath, 'rb') as file:
                            all_current_files[f] = file.read()
                    except:
                        pass
            
            # 保存备份信息
            self.merge_backup = {
                'files': backup_files,
                'db_records': backup_db_records,
                'counter': backup_counter,
                'all_files': all_current_files,
                'selected_files': selected_files.copy()
            }
            self.log(f"✓ 已保存合并前状态（{len(backup_files)}个文件，{len(backup_db_records)}条数据库记录）")
            # ========== 状态保存完成 ==========
            
            # 读取所有文件的来文字号
            output_dir = self.path_config['待打印']
            all_reference_numbers = []
            
            # "来文字号"在表格中的位置：第2行（索引1），第3列（索引2）
            for idx, filename in enumerate(selected_files):
                filepath = os.path.join(output_dir, filename)
                
                try:
                    # 读取Word文档
                    doc = Document(filepath)
                    
                    if len(doc.tables) == 0:
                        self.log(f"  [{idx+1}] {filename}: 文档中没有表格")
                        all_reference_numbers.append("")
                        continue
                    
                    table = doc.tables[0]
                    
                    # 直接根据已知位置读取"来文字号"：第2行（索引1），第3列（索引2）
                    reference_number = ""
                    if len(table.rows) > 1 and len(table.rows[1].cells) > 2:
                        reference_number = table.rows[1].cells[2].text.strip()
                    
                    if reference_number:
                        all_reference_numbers.append(reference_number)
                        self.log(f"  [{idx+1}] {filename}: {reference_number}")
                    else:
                        self.log(f"  [{idx+1}] {filename}: 来文字号为空")
                        all_reference_numbers.append("")
                        
                except Exception as e:
                    self.log(f"读取文件失败 {filename}: {e}")
                    all_reference_numbers.append("")
            
            # 合并来文字号（用换行符连接）
            merged_reference_number = '\n'.join([num for num in all_reference_numbers if num])
            
            if not merged_reference_number:
                messagebox.showerror("错误", "所有批办单都没有找到来文字号，无法合并")
                return
            
            self.log(f"合并后的来文字号：\n{merged_reference_number}")
            
            # 更新第一个文件的来文字号
            first_file = selected_files[0]
            first_filepath = os.path.join(output_dir, first_file)
            
            try:
                doc = Document(first_filepath)
                
                if len(doc.tables) == 0:
                    messagebox.showerror("错误", f"{first_file} 中没有表格")
                    return
                
                table = doc.tables[0]
                
                # 直接根据已知位置更新"来文字号"：第2行（索引1），第3列（索引2）
                if len(table.rows) > 1 and len(table.rows[1].cells) > 2:
                    target_cell = table.rows[1].cells[2]
                    
                    # 根据合并文件数量设置不同格式
                    generator = WordGenerator()
                    file_count = len(selected_files)

                    # 按各行字符数确定字号（取最长行决定档位），行数兜底
                    lines = [l for l in merged_reference_number.split('\n') if l]
                    max_char = max((len(l) for l in lines), default=0)
                    if max_char >= 33:
                        font_size_pt = 9    # 小五号
                    elif max_char >= 29:
                        font_size_pt = 10.5  # 五号
                    elif max_char >= 25:
                        font_size_pt = 12   # 小四号
                    elif file_count >= 4:
                        font_size_pt = 10.5  # 行数兜底：4+行
                    elif file_count == 3:
                        font_size_pt = 12   # 行数兜底：3行
                    else:
                        font_size_pt = 14   # 四号（默认）

                    # 按文件数量确定行间距（原逻辑）
                    if file_count == 2:
                        line_spacing_pt = 17
                    elif file_count == 3:
                        line_spacing_pt = 11
                    else:
                        line_spacing_pt = 10

                    generator.set_cell_text_with_custom_format(
                        target_cell,
                        merged_reference_number,
                        font_size_pt=font_size_pt,
                        line_spacing_pt=line_spacing_pt,
                        alignment='center'
                    )
                    self.log(f"使用格式：{font_size_pt}pt，行距{line_spacing_pt}磅（最长行{max_char}字，共{file_count}个文件）")
                    
                    # 保存文档
                    doc.save(first_filepath)
                    self.log(f"已更新 {first_file} 的来文字号")
                else:
                    messagebox.showerror("错误", f"{first_file} 表格结构不正确")
                    return
                    
            except Exception as e:
                messagebox.showerror("错误", f"更新第一个文件失败: {e}")
                self.log(f"更新文件失败: {e}")
                return
            
            # 删除其他文件及对应的数据库记录
            deleted_count = 0
            for filename in selected_files[1:]:  # 跳过第一个文件
                filepath = os.path.join(output_dir, filename)
                try:
                    # 从文件名提取收文编号（支持新旧格式）
                    import re as _re
                    _ym = _re.search(r'#C-(\d{2})-(\d+)', filename)
                    if _ym:
                        receipt_number = f"C-{_ym.group(1)}-{_ym.group(2)}"
                    else:
                        receipt_number = f"C-{self.extract_number_from_filename(filename)}"
                    
                    # 删除文件
                    os.remove(filepath)
                    deleted_count += 1
                    self.log(f"已删除文件: {filename}")
                    
                    # 删除数据库记录
                    if self.db.delete_complaint(receipt_number):
                        self.log(f"已删除数据库记录: {receipt_number}")
                    else:
                        self.log(f"警告: 数据库记录删除失败: {receipt_number}")
                        
                except Exception as e:
                    self.log(f"删除失败 {filename}: {e}")
            
            # 更新第一个文件（合并后的文件）的数据库记录
            self.log("更新合并后文件的数据库记录...")
            try:
                import re as _re
                _ym = _re.search(r'#C-(\d{2})-(\d+)', first_file)
                if _ym:
                    first_receipt_number = f"C-{_ym.group(1)}-{_ym.group(2)}"
                else:
                    first_receipt_number = f"C-{self.extract_number_from_filename(first_file)}"
                
                # 从合并后的批办单Word文件中重新读取完整数据
                word_data = WordExtractor.extract_complaint_data(first_filepath)
                if word_data:
                    word_data['文件名'] = first_file
                    
                    if self.db.update_complaint(first_receipt_number, word_data):
                        self.log(f"✓ 已更新数据库记录（完整数据）: {first_receipt_number}, 来文字号已合并")
                    else:
                        self.log(f"警告: 数据库记录更新失败: {first_receipt_number}")
                else:
                    self.log(f"✗ 从批办单提取数据失败: {first_file}")
            except Exception as e:
                self.log(f"更新合并后文件的数据库记录失败: {e}")
            
            # 获取第一个被合并文件的编号
            first_merged_number = self.extract_number_from_filename(first_file)
            self.log(f"第一个被合并文件的编号: {first_merged_number}")
            self.log("开始重新编号批办单（仅编号第一个合并文件之后的文件）...")
            try:
                self.renumber_documents_after_merge(first_merged_number)
            except Exception as e:
                self.log(f"重新编号失败: {e}")

            # 刷新列表
            self.refresh_file_list()

            # 定位合并后的文件（使用第一个被合并文件的编号）
            merged_filename = None
            for f in os.listdir(output_dir):
                if f.endswith('.docx') and not f.startswith('~$'):
                    if self.extract_number_from_filename(f) == first_merged_number:
                        merged_filename = f
                        break

            if merged_filename:
                merged_filepath = os.path.join(output_dir, merged_filename)
                if os.path.exists(merged_filepath):
                    self._load_file_content(merged_filepath, merged_filename)
                    self.log(f"✓ 已加载合并后的批办单: {merged_filename}")
                    # 在列表中选中该文件
                    for item in self.file_tree.get_children():
                        values = self.file_tree.item(item)['values']
                        if len(values) >= 3 and values[2] == merged_filename:
                            self.file_tree.selection_set(item)
                            self.file_tree.see(item)
                            break

            # 启用撤销合并按钮
            if self.undo_merge_button:
                self.undo_merge_button.config(state='normal')
                self.log("✓ 撤销合并功能已启用")
            
            # 更新下一个收文编号显示
            self.update_next_receipt_number()

            messagebox.showinfo("完成",
                              f"批办单合并完成！\n\n"
                              f"✅ 已合并 {len(selected_files)} 个批办单的来文字号\n"
                              f"✅ 已保存到第一个批办单（编号保持不变）\n"
                              f"✅ 已删除其他 {deleted_count} 个批办单\n"
                              f"✅ 已重新编号第一个合并文件之后的批办单\n\n"
                              f"💡 如需撤销，请点击'撤销合并'按钮")
            
        except Exception as e:
            messagebox.showerror("错误", f"合并批办单失败: {e}")
            self.log(f"合并批办单失败: {e}")
            import traceback
            traceback.print_exc()
    
    def extract_number_from_filename(self, filename):
        """从文件名中提取编号数字（支持新格式 #C-YY-N 和旧格式 #N）"""
        import re
        # 新格式：投诉批办单#C-26-3.docx → 提取末尾数字 3
        match = re.search(r'#C-\d+-(\d+)', filename)
        if match:
            return int(match.group(1))
        # 旧格式：投诉批办单#3.docx → 提取数字 3
        match = re.search(r'#(\d+)', filename)
        if match:
            return int(match.group(1))
        return 0
    
    def undo_merge(self):
        """撤销上一次的合并操作"""
        try:
            if not self.merge_backup:
                messagebox.showwarning("提示", "没有可撤销的合并操作")
                return
            
            # 确认撤销操作
            selected_files = self.merge_backup.get('selected_files', [])
            file_list = '\n'.join([f"  {i+1}. {f}" for i, f in enumerate(selected_files)])
            confirm_msg = f"即将撤销上一次的合并操作：\n\n{file_list}\n\n" \
                         f"将恢复所有被合并的批办单，并恢复收文编号。\n\n是否继续？"
            
            if not messagebox.askyesno("确认撤销", confirm_msg):
                return
            
            self.log("=" * 60)
            self.log("开始撤销合并操作...")
            
            output_dir = self.path_config['待打印']
            
            # 步骤1：删除当前所有文件（因为重新编号后文件名已改变）
            self.log("步骤1：清理当前文件...")
            current_files = [f for f in os.listdir(output_dir) if f.endswith('.docx') and not f.startswith('~$')]
            for filename in current_files:
                filepath = os.path.join(output_dir, filename)
                try:
                    os.remove(filepath)
                    self.log(f"  已删除: {filename}")
                except Exception as e:
                    self.log(f"  删除失败 {filename}: {e}")
            
            # 步骤2：恢复合并前的所有文件
            self.log("步骤2：恢复合并前的文件...")
            all_files = self.merge_backup.get('all_files', {})
            restored_count = 0
            for filename, file_content in all_files.items():
                filepath = os.path.join(output_dir, filename)
                try:
                    with open(filepath, 'wb') as f:
                        f.write(file_content)
                    restored_count += 1
                    self.log(f"  已恢复: {filename}")
                except Exception as e:
                    self.log(f"  恢复失败 {filename}: {e}")
            
            self.log(f"✓ 已恢复 {restored_count} 个文件")
            
            # 步骤3：恢复数据库记录
            self.log("步骤3：恢复数据库记录...")
            
            # 先清空当前数据库中的所有记录
            try:
                # 获取所有收文编号
                all_receipts = []
                for filename in current_files:
                    file_number = self.extract_number_from_filename(filename)
                    receipt_number = f"C-{file_number}"
                    all_receipts.append(receipt_number)
                
                # 删除这些记录
                for receipt_number in all_receipts:
                    self.db.delete_complaint(receipt_number)
                
                self.log(f"  已清理 {len(all_receipts)} 条数据库记录")
            except Exception as e:
                self.log(f"  清理数据库记录失败: {e}")
            
            # 恢复备份的数据库记录
            db_records = self.merge_backup.get('db_records', {})
            restored_db_count = 0
            for receipt_number, record in db_records.items():
                try:
                    # 删除主键字段（如果存在）
                    record_copy = record.copy()
                    if 'id' in record_copy:
                        del record_copy['id']
                    
                    # 插入记录
                    if self.db.add_complaint(record_copy):
                        restored_db_count += 1
                        self.log(f"  已恢复数据库记录: {receipt_number}")
                    else:
                        self.log(f"  恢复数据库记录失败: {receipt_number}")
                except Exception as e:
                    self.log(f"  恢复数据库记录失败 {receipt_number}: {e}")
            
            self.log(f"✓ 已恢复 {restored_db_count} 条数据库记录")
            
            # 步骤4：恢复收文编号计数器
            self.log("步骤4：恢复收文编号计数器...")
            backup_counter = self.merge_backup.get('counter', {})
            if backup_counter:
                # 使用当前数据库对应的计数器（支持年份数据库）
                db_year = self.get_current_db_year()
                counter = CounterManager(year=db_year)
                counter.file_number = backup_counter.get('file_number', 1)
                counter.receipt_number = backup_counter.get('receipt_number', 1)
                counter.save()
                
                # 同时更新processor的counter（为了兼容性）
                self.processor.counter.file_number = backup_counter.get('file_number', 1)
                self.processor.counter.receipt_number = backup_counter.get('receipt_number', 1)
                self.processor.counter.save()
                
                self.log(f"  文件编号: {counter.file_number}")
                self.log(f"  收文编号: {counter.receipt_number}")
                self.log("✓ 收文编号计数器已恢复")
                
                # 重新加载counter实例，确保从文件读取最新值
                counter.load()
                self.log(f"  重新加载后 - 文件编号: {counter.file_number}, 收文编号: {counter.receipt_number}")
            
            # 步骤5：刷新界面并更新下一个收文编号
            self.log("步骤5：刷新界面并更新收文编号...")
            self.refresh_file_list()
            self.update_next_receipt_number()
            
            # 清空备份并禁用撤销按钮
            self.merge_backup = None
            if self.undo_merge_button:
                self.undo_merge_button.config(state='disabled')
                self.log("✓ 撤销合并按钮已禁用")
            
            self.log("=" * 60)
            self.log("✓ 撤销合并操作完成")
            
            messagebox.showinfo("完成", 
                              f"撤销合并操作完成！\n\n"
                              f"✅ 已恢复 {restored_count} 个批办单文件\n"
                              f"✅ 已恢复 {restored_db_count} 条数据库记录\n"
                              f"✅ 已恢复收文编号计数器\n"
                              f"✅ 所有状态已恢复到合并前")
            
        except Exception as e:
            messagebox.showerror("错误", f"撤销合并失败: {e}")
            self.log(f"撤销合并失败: {e}")
            import traceback
            traceback.print_exc()
    
    def renumber_all_documents(self):
        """重新编号所有待打印批办单，保持连续
        
        规则：保持第一个文件的编号不变，后续文件依次递增
        例如：#140、#142、#145 -> #140、#141、#142
        """
        import re
        
        output_dir = self.path_config['待打印']
        
        # 获取所有批办单文件
        files = [f for f in os.listdir(output_dir) if f.endswith('.docx') and not f.startswith('~$')]
        
        if not files:
            return
        
        # 按当前编号排序
        files.sort(key=self.extract_number_from_filename)
        
        self.log(f"找到 {len(files)} 个批办单需要重新编号")
        
        # 获取第一个文件的编号作为起始编号
        first_number = self.extract_number_from_filename(files[0])
        self.log(f"起始编号: {first_number}")
        
        # 创建临时文件映射，避免重命名冲突
        temp_files = []
        for idx, old_filename in enumerate(files):
            old_path = os.path.join(output_dir, old_filename)
            temp_path = os.path.join(output_dir, f"_temp_{idx}_{old_filename}")
            
            try:
                # 先重命名为临时文件名
                os.rename(old_path, temp_path)
                # 新编号 = 起始编号 + 索引
                new_number = first_number + idx
                temp_files.append((temp_path, old_filename, new_number))
            except Exception as e:
                self.log(f"临时重命名失败 {old_filename}: {e}")
        
        # 重新编号并更新文档和数据库
        for temp_path, old_filename, new_number in temp_files:
            try:
                # 生成新文件名（保留原文件的年份信息，若无则用当前年份）
                import re
                from datetime import datetime
                year_match = re.search(r'#C-(\d{2})-', old_filename)
                if year_match:
                    year_suffix = year_match.group(1)
                else:
                    year_suffix = str(datetime.now().year)[-2:]
                new_filename = f"投诉批办单#C-{year_suffix}-{new_number}.docx"
                new_path = os.path.join(output_dir, new_filename)
                
                # 提取旧的收文编号（支持新旧两种格式）
                old_number = self.extract_number_from_filename(old_filename)
                old_year_match = re.search(r'#C-(\d{2})-', old_filename)
                if old_year_match:
                    old_year = old_year_match.group(1)
                    old_receipt_number = f"C-{old_year}-{old_number}"
                else:
                    old_receipt_number = f"C-{old_number}"
                new_receipt_number = f"C-{year_suffix}-{new_number}"
                
                # 更新文档内的收文编号
                doc = Document(temp_path)
                
                if len(doc.tables) > 0:
                    table = doc.tables[0]
                    # "收文编号"在第4行（索引3），第10列（索引9）
                    if len(table.rows) > 3 and len(table.rows[3].cells) > 9:
                        receipt_cell = table.rows[3].cells[9]
                        generator = WordGenerator()
                        generator.set_cell_text_keep_format(receipt_cell, new_receipt_number, alignment='center')
                    
                    # 保存文档
                    doc.save(temp_path)
                
                # 重命名为最终文件名
                os.rename(temp_path, new_path)
                
                # 同步更新数据库：从批办单Word文件重新读取所有内容并更新
                old_record = self.db.get_complaint_by_receipt_number(old_receipt_number)
                if old_record:
                    # 从更新后的批办单Word文件中重新读取完整的10个字段数据
                    word_data = WordExtractor.extract_complaint_data(new_path)
                    if word_data:
                        from datetime import datetime
                        # 保留部分原有数据（收文日期、领导批示），其他使用Word中的最新内容
                        word_data['收文日期'] = old_record.get('收文日期', datetime.now().strftime('%Y-%m-%d'))
                        word_data['领导批示'] = old_record.get('领导批示', '')
                        word_data['文件名'] = new_filename
                        
                        if old_receipt_number != new_receipt_number:
                            # 编号改变：删除旧记录，添加新记录
                            self.db.delete_complaint(old_receipt_number)
                            word_data['收文编号'] = new_receipt_number
                            if self.db.add_complaint(word_data):
                                self.log(f"  数据库已更新（完整数据）: {old_receipt_number} -> {new_receipt_number}")
                            else:
                                self.log(f"  警告: 数据库更新失败: {old_receipt_number} -> {new_receipt_number}")
                        else:
                            # 编号未变：直接更新记录内容
                            if self.db.update_complaint(old_receipt_number, word_data):
                                self.log(f"  数据库已更新（完整数据）: {old_receipt_number}")
                            else:
                                self.log(f"  警告: 数据库更新失败: {old_receipt_number}")
                    else:
                        self.log(f"  警告: 无法从批办单读取数据: {new_filename}")
                else:
                    self.log(f"  提示: 数据库中未找到记录 {old_receipt_number}")
                
                self.log(f"重新编号: {old_filename} -> {new_filename} (收文编号: {new_receipt_number})")
                
            except Exception as e:
                self.log(f"重新编号失败 {old_filename}: {e}")
                import traceback
                traceback.print_exc()
                # 如果失败，至少恢复原文件名
                try:
                    fallback_path = os.path.join(output_dir, old_filename)
                    if os.path.exists(temp_path):
                        os.rename(temp_path, fallback_path)
                except:
                    pass
        
        self.log(f"批办单重新编号完成，共处理 {len(temp_files)} 个文件，编号范围: C-{first_number} 到 C-{first_number + len(temp_files) - 1}")
        
        # 重新编号后，更新计数器，使下一个编号连续
        if temp_files:
            # 获取最后一个文件的编号
            last_number = first_number + len(temp_files) - 1
            
            # 使用当前数据库对应的计数器（支持年份数据库）
            db_year = self.get_current_db_year()
            counter = CounterManager(year=db_year)
            counter.receipt_number = last_number + 1
            counter.file_number = last_number + 1
            counter.save()
            
            # 同时更新processor的counter（为了兼容性）
            self.processor.counter.receipt_number = last_number + 1
            self.processor.counter.file_number = last_number + 1
            self.processor.counter.save()
            
            self.log(f"已更新计数器，下一个编号: C-{last_number + 1}")
            
            # 更新界面显示
            self.update_next_receipt_number()
    
    def renumber_documents_after_merge(self, first_merged_number):
        """合并后重新编号：只编号第一个被合并文件之后的文件
        
        规则：
        - 第一个被合并文件的编号保持不变
        - 第一个被合并文件之前的文件编号不变
        - 第一个被合并文件之后的文件按顺序重新编号
        
        参数:
            first_merged_number: 第一个被合并文件的编号
        
        例如：
            合并前：C-26-1, C-26-2, C-26-4(合并), C-26-5, C-26-6(合并), C-26-7
            合并后删除C-26-6，剩余：C-26-1, C-26-2, C-26-4, C-26-5, C-26-7
            重新编号后：C-26-1(不变), C-26-2(不变), C-26-4(不变), C-26-5(不变), C-26-6(C-26-7重编号)
        """
        import re
        
        output_dir = self.path_config['待打印']
        
        # 获取所有批办单文件
        files = [f for f in os.listdir(output_dir) if f.endswith('.docx') and not f.startswith('~')]
        
        if not files:
            return
        
        # 按当前编号排序
        files.sort(key=self.extract_number_from_filename)
        
        self.log(f"找到 {len(files)} 个批办单")
        self.log(f"第一个被合并文件的编号: {first_merged_number}")
        
        # 找到需要重新编号的文件（第一个被合并文件之后的文件）
        files_to_renumber = []
        for f in files:
            file_number = self.extract_number_from_filename(f)
            if file_number > first_merged_number:
                files_to_renumber.append(f)
        
        if not files_to_renumber:
            self.log("没有需要重新编号的文件，更新计数器...")
            # 即使没有文件需要重新编号，也要更新计数器
            # 找到所有文件中的最大编号
            max_number = max([self.extract_number_from_filename(f) for f in files]) if files else first_merged_number
            
            # 更新计数器为最大编号 + 1
            db_year = self.get_current_db_year()
            counter = CounterManager(year=db_year)
            counter.receipt_number = max_number + 1
            counter.file_number = max_number + 1
            counter.save()
            
            # 同时更新processor的counter（为了兼容性）
            self.processor.counter.receipt_number = max_number + 1
            self.processor.counter.file_number = max_number + 1
            self.processor.counter.save()
            
            self.log(f"已更新计数器，下一个编号: C-{max_number + 1}")
            
            # 更新界面显示
            self.update_next_receipt_number()
            return
        
        self.log(f"需要重新编号 {len(files_to_renumber)} 个文件")
        
        # 创建临时文件映射，避免重命名冲突
        temp_files = []
        for idx, old_filename in enumerate(files_to_renumber):
            old_path = os.path.join(output_dir, old_filename)
            temp_path = os.path.join(output_dir, f"_temp_merge_{idx}_{old_filename}")
            
            try:
                # 先重命名为临时文件名
                os.rename(old_path, temp_path)
                # 新编号 = 第一个被合并文件编号 + 1 + 索引
                new_number = first_merged_number + 1 + idx
                temp_files.append((temp_path, old_filename, new_number))
            except Exception as e:
                self.log(f"临时重命名失败 {old_filename}: {e}")
        
        # 重新编号并更新文档和数据库
        for temp_path, old_filename, new_number in temp_files:
            try:
                # 生成新文件名（保留原文件的年份信息，若无则用当前年份）
                import re
                from datetime import datetime
                year_match = re.search(r'#C-(\d{2})-', old_filename)
                if year_match:
                    year_suffix = year_match.group(1)
                else:
                    year_suffix = str(datetime.now().year)[-2:]
                new_filename = f"投诉批办单#C-{year_suffix}-{new_number}.docx"
                new_path = os.path.join(output_dir, new_filename)
                
                # 提取旧的收文编号
                old_number = self.extract_number_from_filename(old_filename)
                old_year_match = re.search(r'#C-(\d{2})-', old_filename)
                if old_year_match:
                    old_year = old_year_match.group(1)
                    old_receipt_number = f"C-{old_year}-{old_number}"
                else:
                    old_receipt_number = f"C-{old_number}"
                new_receipt_number = f"C-{year_suffix}-{new_number}"
                
                # 更新文档内的收文编号
                doc = Document(temp_path)
                
                if len(doc.tables) > 0:
                    table = doc.tables[0]
                    # "收文编号"在第4行（索引3），第10列（索引9）
                    if len(table.rows) > 3 and len(table.rows[3].cells) > 9:
                        receipt_cell = table.rows[3].cells[9]
                        generator = WordGenerator()
                        generator.set_cell_text_keep_format(receipt_cell, new_receipt_number, alignment='center')
                    
                    # 保存文档
                    doc.save(temp_path)
                
                # 重命名为最终文件名
                os.rename(temp_path, new_path)
                
                # 同步更新数据库
                old_record = self.db.get_complaint_by_receipt_number(old_receipt_number)
                if old_record:
                    # 从更新后的批办单Word文件中重新读取完整的10个字段数据
                    word_data = WordExtractor.extract_complaint_data(new_path)
                    if word_data:
                        from datetime import datetime
                        # 保留部分原有数据（收文日期、领导批示）
                        word_data['收文日期'] = old_record.get('收文日期', datetime.now().strftime('%Y-%m-%d'))
                        word_data['领导批示'] = old_record.get('领导批示', '')
                        word_data['文件名'] = new_filename
                        
                        if old_receipt_number != new_receipt_number:
                            # 编号改变：删除旧记录，添加新记录
                            self.db.delete_complaint(old_receipt_number)
                            word_data['收文编号'] = new_receipt_number
                            if self.db.add_complaint(word_data):
                                self.log(f"  数据库已更新：{old_receipt_number} -> {new_receipt_number}")
                            else:
                                self.log(f"  警告: 数据库更新失败: {old_receipt_number} -> {new_receipt_number}")
                        else:
                            # 编号未变：直接更新记录内容
                            if self.db.update_complaint(old_receipt_number, word_data):
                                self.log(f"  数据库已更新: {old_receipt_number}")
                            else:
                                self.log(f"  警告: 数据库更新失败: {old_receipt_number}")
                    else:
                        self.log(f"  警告: 无法从批办单读取数据: {new_filename}")
                else:
                    self.log(f"  提示: 数据库中未找到记录 {old_receipt_number}")
                
                self.log(f"  重新编号: {old_filename} -> {new_filename}")
                
            except Exception as e:
                self.log(f"重新编号失败 {old_filename}: {e}")
                import traceback
                traceback.print_exc()
                # 如果失败，尝试恢复原文件名
                try:
                    fallback_path = os.path.join(output_dir, old_filename)
                    if os.path.exists(temp_path):
                        os.rename(temp_path, fallback_path)
                except:
                    pass
        
        if temp_files:
            # 获取最后一个文件的编号
            last_number = first_merged_number + len(temp_files)
            self.log(f"重新编号完成，编号范围: C-{first_merged_number + 1} 到 C-{last_number}")
            
            # 更新计数器
            db_year = self.get_current_db_year()
            counter = CounterManager(year=db_year)
            counter.receipt_number = last_number + 1
            counter.file_number = last_number + 1
            counter.save()
            
            # 同时更新processor的counter（为了兼容性）
            self.processor.counter.receipt_number = last_number + 1
            self.processor.counter.file_number = last_number + 1
            self.processor.counter.save()
            
            self.log(f"已更新计数器，下一个编号: C-{last_number + 1}")
            
            # 更新界面显示
            self.update_next_receipt_number()
        
    def open_output_dir(self):
        """打开待打印目录"""
        output_dir = os.path.abspath(self.path_config['待打印'])
        open_path(output_dir)
    
    def open_printed_dir(self):
        """打开已打印目录"""
        printed_dir = os.path.abspath(self.path_config['已打印'])
        open_path(printed_dir)
    
    def open_processed_dir(self):
        """打开已处理目录"""
        processed_dir = os.path.abspath(self.path_config['已处理'])
        open_path(processed_dir)
    
    def manage_paths(self):
        """路径配置管理"""
        dialog = tk.Toplevel(self.root)
        dialog.title("路径设置")
        dialog.geometry("600x300")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="文件夹路径配置", font=(ui_font('微软雅黑'), 12, 'bold')).pack(pady=10)
        
        # 创建输入框
        entries = {}
        for key, default_path in self.path_config.items():
            frame = ttk.Frame(dialog)
            frame.pack(fill=tk.X, padx=20, pady=10)
            
            ttk.Label(frame, text=key + ":", width=10).pack(side=tk.LEFT)
            entry = ttk.Entry(frame, width=40)
            entry.insert(0, default_path)
            entry.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
            
            def browse(e=entry, k=key):
                folder = filedialog.askdirectory(title=f"选择{k}文件夹", initialdir=e.get())
                if folder:
                    e.delete(0, tk.END)
                    e.insert(0, folder)
            
            ttk.Button(frame, text="浏览...", command=browse, width=8).pack(side=tk.LEFT)
            entries[key] = entry
        
        def save_paths():
            """保存路径配置"""
            new_config = {}
            for key, entry in entries.items():
                path = entry.get().strip()
                # "默认待处理文件夹"可以为空，其他路径不能为空
                if not path and key != '默认待处理文件夹':
                    messagebox.showerror("错误", f"{key}路径不能为空")
                    return
                new_config[key] = path
            
            # 保存到配置文件
            try:
                with open(get_appdata_path('path_config.json'), 'w', encoding='utf-8') as f:
                    json.dump(new_config, f, ensure_ascii=False, indent=2)
                
                self.path_config = new_config
                self.ensure_directories()
                messagebox.showinfo("成功", "路径配置已保存")
                dialog.destroy()
                self.refresh_all_lists()
            except Exception as e:
                messagebox.showerror("错误", f"保存配置失败: {e}")
        
        # 按钮
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=20)
        ttk.Button(btn_frame, text="保存", command=save_paths, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy, width=15).pack(side=tk.LEFT, padx=5)
        
    def manage_counter(self):
        """编号管理（支持按年份管理）"""
        # 获取当前数据库年份
        db_year = self.get_current_db_year()
        
        # 使用年份特定的计数器
        counter = CounterManager(year=db_year)
        
        dialog = tk.Toplevel(self.root)
        dialog.title("编号管理")
        dialog.geometry("450x180")
        
        # 显示当前数据库年份
        year_text = f"{db_year}年" if db_year else "默认"
        ttk.Label(dialog, text=f"当前编号设置 ({year_text}数据库)", 
                 font=(ui_font('微软雅黑'), 12, 'bold')).pack(pady=10)
        
        # 统一编号（文件编号和收文编号相同）
        frame1 = ttk.Frame(dialog)
        frame1.pack(fill=tk.X, padx=20, pady=5)
        ttk.Label(frame1, text="当前编号:").pack(side=tk.LEFT)
        counter_var = tk.StringVar(value=str(counter.file_number))
        ttk.Entry(frame1, textvariable=counter_var, width=10).pack(side=tk.LEFT, padx=10)
        ttk.Label(frame1, text="（文件编号和收文编号统一）", font=(ui_font('微软雅黑'), 9), foreground='gray').pack(side=tk.LEFT)
        
        # 说明文字
        if db_year:
            info_text = f"此编号仅用于{db_year}年数据库，切换数据库将使用对应年份的编号"
        else:
            info_text = "当前使用默认编号，建议切换到年份数据库进行管理"
        ttk.Label(dialog, text=info_text, font=(ui_font('微软雅黑'), 9), foreground='#666').pack(pady=5)
        
        def save_counter():
            try:
                # 更新counter实例的值（文件编号和收文编号保持一致）
                new_number = int(counter_var.get())
                counter.set_numbers(new_number, new_number)
                
                # 更新收文编号显示
                self.update_next_receipt_number()
                
                year_info = f"{db_year}年" if db_year else "默认"
                self.log(f"{year_info}数据库编号已更新: {new_number}")
                messagebox.showinfo("成功", f"{year_info}数据库编号已更新为: {new_number}")
                dialog.destroy()
            except ValueError:
                messagebox.showerror("错误", "请输入有效的数字")
        
        ttk.Button(dialog, text="保存", command=save_counter, bootstyle="primary").pack(pady=15)
    
    def manage_qrcode(self):
        """二维码设置"""
        dialog = tk.Toplevel(self.root)
        dialog.title("二维码设置")
        dialog.geometry("500x580")
        dialog.resizable(False, False)
        
        # 主框架
        main_frame = ttk.Frame(dialog)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 标题
        ttk.Label(main_frame, text="📱 二维码设置", 
                 font=(ui_font('微软雅黑'), 14, 'bold')).pack(pady=(0, 15))
        
        # 启用开关
        enable_frame = ttk.Frame(main_frame)
        enable_frame.pack(fill=tk.X, pady=8)
        
        enable_var = tk.BooleanVar(value=self.qrcode_config['enabled'])
        ttk.Checkbutton(enable_frame, text="启用二维码功能", 
                       variable=enable_var,
                       bootstyle="success-round-toggle").pack(side=tk.LEFT)
        
        # 二维码尺寸选择
        size_frame = ttk.LabelFrame(main_frame, text="二维码尺寸", padding=15)
        size_frame.pack(fill=tk.X, pady=8)
        
        size_var = tk.StringVar(value=self.qrcode_config['size'])
        ttk.Radiobutton(size_frame, text="大 (2.7cm × 2.7cm)", 
                       variable=size_var, value='large').pack(anchor=tk.W, pady=3)
        ttk.Radiobutton(size_frame, text="小 (1.7cm × 1.7cm)", 
                       variable=size_var, value='small').pack(anchor=tk.W, pady=3)
        
        # 编码字段选择
        fields_frame = ttk.LabelFrame(main_frame, text="编码内容（选择需要编入二维码的字段）", padding=15)
        fields_frame.pack(fill=tk.X, pady=8)
        
        field_vars = {}
        fields = ['收文编号', '收文日期', '来文字号', '来文单位', '收文途径', '文件标题']
        
        for field in fields:
            var = tk.BooleanVar(value=self.qrcode_config['fields'].get(field, False))
            field_vars[field] = var
            ttk.Checkbutton(fields_frame, text=field, variable=var,
                           bootstyle="primary-round-toggle").pack(anchor=tk.W, pady=3)
        
        # 说明文字
        info_text = "二维码将放置在批办单右下角（悬浮于文字之上）"
        ttk.Label(main_frame, text=info_text, font=(ui_font('微软雅黑'), 9), 
                 foreground='#666').pack(pady=8)
        
        # 按钮框架 - 确保可见
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=15, side=tk.BOTTOM)
        
        def save_config():
            # 更新配置
            self.qrcode_config['enabled'] = enable_var.get()
            self.qrcode_config['size'] = size_var.get()
            for field, var in field_vars.items():
                self.qrcode_config['fields'][field] = var.get()
            
            # 保存到文件
            if self.save_qrcode_config(self.qrcode_config):
                messagebox.showinfo("成功", "二维码设置已保存")
                dialog.destroy()
        
        ttk.Button(btn_frame, text="保存", command=save_config, 
                  bootstyle="success", width=12).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy, 
                  width=12).pack(side=tk.LEFT, padx=5)
        
    def generate_report(self):
        """生成统计报表"""
        if not self.file_list:
            messagebox.showwarning("提示", "没有文件可统计")
            return
        
        report = f"投诉批办单统计报表\n"
        report += f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        report += f"=" * 50 + "\n\n"
        report += f"文件总数: {len(self.file_list)}\n"
        report += f"存储位置: {os.path.abspath('待打印')}\n\n"
        report += f"文件列表:\n"
        
        for idx, filename in enumerate(self.file_list, 1):
            filepath = os.path.join('待打印', filename)
            size = os.path.getsize(filepath) / 1024  # KB
            mtime = datetime.fromtimestamp(os.path.getmtime(filepath)).strftime('%Y-%m-%d %H:%M')
            report += f"{idx}. {filename} ({size:.1f}KB, {mtime})\n"
        
        # 显示报表
        dialog = tk.Toplevel(self.root)
        dialog.title("统计报表")
        dialog.geometry("700x500")
        
        text = scrolledtext.ScrolledText(dialog, wrap=tk.WORD, font=(ui_font('Consolas'), 9))
        text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        text.insert(1.0, report)
        text.config(state=tk.DISABLED)
        
        # 按钮框架
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="导出文本报表", 
                  command=lambda: self.export_report(report)).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="导出Excel报表", 
                  command=lambda: self.export_to_excel()).pack(side=tk.LEFT, padx=5)
        
    def export_report(self, report):
        """导出报表到文本文件"""
        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
            initialfile=f"统计报表_{datetime.now().strftime('%Y%m%d')}.txt"
        )
        
        if filename:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(report)
            messagebox.showinfo("成功", "文本报表已导出")
            self.log(f"已导出文本报表: {filename}")
    
    def export_to_excel(self):
        """导出到Excel文件"""
        if not self.file_list:
            messagebox.showwarning("提示", "没有文件可导出")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")],
            initialfile=f"批办单统计_{datetime.now().strftime('%Y%m%d')}.xlsx"
        )
        
        if not filename:
            return
        
        try:
            # 创建工作簿
            wb = Workbook()
            ws = wb.active
            ws.title = "批办单列表"
            
            # 设置列宽
            ws.column_dimensions['A'].width = 8
            ws.column_dimensions['B'].width = 30
            ws.column_dimensions['C'].width = 15
            ws.column_dimensions['D'].width = 12
            ws.column_dimensions['E'].width = 20
            ws.column_dimensions['F'].width = 20
            
            # 标题行样式
            title_font = Font(name=ui_font('微软雅黑'), size=14, bold=True, color='FFFFFF')
            title_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
            title_alignment = Alignment(horizontal='center', vertical='center')
            
            # 写入标题
            ws.merge_cells('A1:F1')
            cell = ws['A1']
            cell.value = f"投诉批办单统计表 - {datetime.now().strftime('%Y年%m月%d日')}"
            cell.font = title_font
            cell.fill = title_fill
            cell.alignment = title_alignment
            ws.row_dimensions[1].height = 30
            
            # 表头样式
            header_font = Font(name=ui_font('微软雅黑'), size=11, bold=True)
            header_fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
            header_alignment = Alignment(horizontal='center', vertical='center')
            
            # 写入表头
            headers = ['序号', '文件名', '文件编号', '文件大小', '创建日期', '修改日期']
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=2, column=col)
                cell.value = header
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_alignment
            ws.row_dimensions[2].height = 25
            
            # 内容样式
            content_font = Font(name=ui_font('微软雅黑'), size=10)
            content_alignment = Alignment(horizontal='left', vertical='center')
            
            # 写入数据
            for idx, filename in enumerate(self.file_list, 1):
                filepath = os.path.join('待打印', filename)
                
                try:
                    # 提取编号
                    number = filename.replace('投诉批办单#', '').replace('.docx', '')
                except:
                    number = '-'
                
                # 文件大小
                try:
                    size_kb = os.path.getsize(filepath) / 1024
                    size_str = f"{size_kb:.1f}KB"
                except:
                    size_str = '-'
                
                # 创建时间
                try:
                    ctime = datetime.fromtimestamp(os.path.getctime(filepath)).strftime('%Y-%m-%d %H:%M')
                except:
                    ctime = '-'
                
                # 修改时间
                try:
                    mtime = datetime.fromtimestamp(os.path.getmtime(filepath)).strftime('%Y-%m-%d %H:%M')
                except:
                    mtime = '-'
                
                row = idx + 2
                data = [idx, filename, number, size_str, ctime, mtime]
                
                for col, value in enumerate(data, 1):
                    cell = ws.cell(row=row, column=col)
                    cell.value = value
                    cell.font = content_font
                    if col == 1:  # 序号居中
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    else:
                        cell.alignment = content_alignment
            
            # 添加汇总信息
            summary_row = len(self.file_list) + 4
            ws.merge_cells(f'A{summary_row}:F{summary_row}')
            summary_cell = ws[f'A{summary_row}']
            summary_cell.value = f"总计: {len(self.file_list)} 个文件"
            summary_cell.font = Font(name=ui_font('微软雅黑'), size=11, bold=True)
            summary_cell.alignment = Alignment(horizontal='left', vertical='center')
            
            # 保存文件
            wb.save(filename)
            messagebox.showinfo("成功", f"Excel报表已导出\n保存位置: {filename}")
            self.log(f"已导出Excel报表: {filename}")
            
            # 询问是否打开
            if messagebox.askyesno("打开文件", "是否立即打开Excel文件？"):
                open_path(filename)
                
        except Exception as e:
            messagebox.showerror("错误", f"导出Excel失败: {e}")
            self.log(f"导出Excel失败: {e}")
    
    def sort_db_tree_column(self, col):
        """点击列标题时对数据库检索结果进行排序
        
        Args:
            col: 列名
        """
        try:
            # 获取当前列的排序状态
            current_state = self.db_sort_state.get(col)
            
            # 切换排序状态：None -> 升序(True) -> 降序(False) -> 升序(True) ...
            if current_state is None or current_state is False:
                # 升序
                reverse = False
                self.db_sort_state[col] = True
                sort_indicator = ' ▲'
            else:
                # 降序
                reverse = True
                self.db_sort_state[col] = False
                sort_indicator = ' ▼'
            
            # 重置其他列的排序状态
            for other_col in self.db_sort_state:
                if other_col != col:
                    self.db_sort_state[other_col] = None
            
            # 获取所有行数据
            data = []
            for item in self.db_tree.get_children():
                values = self.db_tree.item(item)['values']
                tags = self.db_tree.item(item)['tags']
                data.append((values, tags))
            
            # 根据列进行排序
            col_index = list(self.db_tree['columns']).index(col)
            
            # 定义排序键函数
            def sort_key(item):
                value = item[0][col_index]
                
                # 对于收文编号（如 C-26-3 或旧格式 C-123），提取末尾数字部分进行排序
                if col == '收文编号' and isinstance(value, str):
                    import re
                    # 新格式 C-YY-N，取末尾数字
                    match = re.search(r'C-\d+-(\d+)$', value)
                    if match:
                        return int(match.group(1))
                    # 旧格式 C-N
                    match = re.search(r'(\d+)$', value)
                    if match:
                        return int(match.group(1))
                    return 0
                
                # 对于日期，转换为可比较的格式
                if col == '收文日期' and isinstance(value, str):
                    try:
                        # 尝试解析日期格式 YYYY-M-D 或 YYYY-MM-DD
                        parts = value.split('-')
                        if len(parts) == 3:
                            return (int(parts[0]), int(parts[1]), int(parts[2]))
                    except:
                        pass
                    return (0, 0, 0)
                
                # 对于序号，转换为整数
                if col == '序号' and isinstance(value, (int, str)):
                    try:
                        return int(value)
                    except:
                        return 0
                
                # 其他列按字符串排序
                return str(value) if value else ''
            
            data.sort(key=sort_key, reverse=reverse)
            
            # 清空树并重新插入排序后的数据
            for item in self.db_tree.get_children():
                self.db_tree.delete(item)
            
            for idx, (values, tags) in enumerate(data, start=1):
                # 更新序号
                new_values = list(values)
                new_values[0] = idx
                self.db_tree.insert('', 'end', values=tuple(new_values), tags=tags)
            
            # 更新列标题，显示排序指示器
            for column in self.db_tree['columns']:
                if column == col:
                    self.db_tree.heading(column, text=column + sort_indicator)
                else:
                    self.db_tree.heading(column, text=column)
            
            self.log(f"已按 {col} {'升序' if not reverse else '降序'} 排序")
            
        except Exception as e:
            self.log(f"排序失败: {e}")
            import traceback
            traceback.print_exc()
    
    def search_database(self):
        """搜索数据库"""
        try:
            keyword = self.db_keyword_entry.get().strip()
            field = self.db_field_var.get()
            start_date = self.db_start_date_entry.get().strip()
            end_date = self.db_end_date_entry.get().strip()

            results = self.db.search_complaints(keyword=keyword, field=field,
                                                start_date=start_date, end_date=end_date)

            # 清空列表
            for item in self.db_tree.get_children():
                self.db_tree.delete(item)

            # 填充结果（添加序号）
            for idx, record in enumerate(results, start=1):
                self.db_tree.insert('', 'end', values=(
                    idx,  # 序号
                    record.get('收文编号', ''),
                    record.get('来文单位', ''),
                    record.get('文件标题', ''),
                    record.get('收文日期', '')
                ), tags=(record.get('收文编号', ''),))  # 使用收文编号作为标识

            self.log(f"搜索完成，找到 {len(results)} 条记录")

            # 默认按收文编号降序排序（自动执行排序）
            if len(results) > 0:
                self.db_sort_state['收文编号'] = True  # 设置为升序状态
                self.sort_db_tree_column('收文编号')  # 切换为降序，显示最新记录在前

        except Exception as e:
            messagebox.showerror("错误", f"搜索失败: {e}")
            self.log(f"数据库搜索失败: {e}")
    
    def refresh_database(self):
        """刷新数据库列表（显示全部）"""
        try:
            # 清空搜索条件
            self.db_keyword_entry.delete(0, tk.END)
            self.db_field_var.set('all')
            self.db_start_date_entry.delete(0, tk.END)
            self.db_end_date_entry.delete(0, tk.END)
            
            # 搜索全部
            self.search_database()
            
        except Exception as e:
            self.log(f"刷新数据库失败: {e}")
    
    def create_new_database(self):
        """创建新的年度数据库"""
        try:
            # 弹出对话框输入年份
            from tkinter import simpledialog
            year = simpledialog.askstring("新建数据库", "请输入年份 (例如: 2027):", 
                                         initialvalue=str(datetime.now().year + 1))
            
            if not year:
                return
            
            # 验证年份格式
            try:
                year_int = int(year)
                if year_int < 2000 or year_int > 2100:
                    messagebox.showerror("错误", "年份必须在 2000-2100 之间")
                    return
            except ValueError:
                messagebox.showerror("错误", "请输入有效的年份数字")
                return
            
            # 生成数据库文件名
            from database import get_appdata_path
            db_filename = f"complaints-{year}.db"
            db_path = get_appdata_path(db_filename)
            
            # 检查文件是否已存在
            if os.path.exists(db_path):
                result = messagebox.askyesno("确认", 
                    f"数据库 {db_filename} 已存在，是否切换到该数据库？")
                if result:
                    self.switch_database(db_path)
                return
            
            # 创建新数据库
            self.db = ComplaintDatabase(db_path)
            # 保存当前数据库路径，下次启动自动使用
            self.save_last_database(db_path)
            self.update_current_db_label()
            self.refresh_database()
            
            # 更新收文编号显示（使用新数据库对应的计数器）
            self.update_next_receipt_number()
            
            messagebox.showinfo("成功", f"已创建数据库: {db_filename}")
            self.log(f"创建新数据库: {db_filename}")
            
        except Exception as e:
            messagebox.showerror("错误", f"创建数据库失败: {e}")
            self.log(f"创建数据库失败: {e}")
    
    def select_database(self):
        """选择已存在的数据库文件"""
        try:
            from tkinter import filedialog
            from database import get_appdata_path
            
            # 打开文件选择对话框，初始目录为AppData
            appdata_dir = get_appdata_path()
            db_path = filedialog.askopenfilename(
                title="选择数据库文件",
                initialdir=appdata_dir,
                filetypes=[("数据库文件", "*.db"), ("所有文件", "*.*")]
            )
            
            if not db_path:
                return
            
            # 切换数据库
            self.switch_database(db_path)
            
        except Exception as e:
            messagebox.showerror("错误", f"选择数据库失败: {e}")
            self.log(f"选择数据库失败: {e}")
    
    def switch_database(self, db_path):
        """切换到指定的数据库"""
        try:
            # 初始化新的数据库连接
            self.db = ComplaintDatabase(db_path)
            # 保存当前数据库路径，下次启动自动使用
            self.save_last_database(db_path)
            self.update_current_db_label()
            self.refresh_database()
            
            # 更新收文编号显示（使用新数据库对应的计数器）
            self.update_next_receipt_number()
            
            db_filename = os.path.basename(db_path)
            messagebox.showinfo("成功", f"已切换到数据库: {db_filename}")
            self.log(f"切换到数据库: {db_filename}")
            
        except Exception as e:
            messagebox.showerror("错误", f"切换数据库失败: {e}")
            self.log(f"切换数据库失败: {e}")
    
    def update_current_db_label(self):
        """更新当前数据库显示标签"""
        try:
            db_filename = os.path.basename(self.db.db_path)
            self.current_db_label.config(text=f"当前数据库: {db_filename}")
        except:
            self.current_db_label.config(text="当前数据库: complaints.db")
    
    def get_current_db_year(self):
        """获取当前数据库的年份"""
        import re
        try:
            db_filename = os.path.basename(self.db.db_path)
            # 从数据库文件名提取年份，如 complaints-2026.db → 2026
            year_match = re.search(r'complaints-(\d{4})\.db', db_filename)
            if year_match:
                return year_match.group(1)
        except:
            pass
        # 默认返回None，表示使用旧的counter.json
        return None
    
    def show_statistics(self):
        """显示统计信息"""
        try:
            stats = self.db.get_statistics()
            
            stats_text = "数据库统计信息\n\n"
            for key, value in stats.items():
                stats_text += f"{key}: {value}\n"
            
            messagebox.showinfo("统计信息", stats_text)
            
        except Exception as e:
            messagebox.showerror("错误", f"获取统计信息失败: {e}")
    
    def export_to_excel(self):
        """导出数据库记录到Excel文件"""
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, Alignment, Border, Side
            from tkinter import filedialog
            import re
            
            # 获取所有记录
            records = self.db.search_complaints()
            
            if not records:
                messagebox.showwarning("提示", "没有可导出的记录")
                return
            
            # 反转记录顺序，从最早到最新（正序）
            records = list(reversed(records))
            
            # 提取年份（从当前数据库文件名或第一条记录）
            year = datetime.now().year
            try:
                db_filename = os.path.basename(self.db.db_path)
                # 从数据库文件名提取年份，如 complaints-2026.db
                year_match = re.search(r'complaints-(\d{4})\.db', db_filename)
                if year_match:
                    year = year_match.group(1)
                else:
                    # 尝试从第一条记录的收文日期提取年份
                    if records and records[0].get('收文日期'):
                        date_str = records[0]['收文日期']
                        year_from_date = re.search(r'(\d{4})', date_str)
                        if year_from_date:
                            year = year_from_date.group(1)
            except:
                pass
            
            # 让用户选择保存位置
            default_filename = f"{year}年投诉收文登记.xlsx"
            file_path = filedialog.asksaveasfilename(
                title="导出Excel文件",
                defaultextension=".xlsx",
                initialfile=default_filename,
                filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
            )
            
            if not file_path:
                return
            
            # 创建工作簿
            wb = Workbook()
            ws = wb.active
            ws.title = "收文登记"
            
            # 第一行：标题
            title = f"西安市雁塔区卫生健康局{year}年投诉收文登记"
            ws.merge_cells('A1:L1')
            title_cell = ws['A1']
            title_cell.value = title
            title_cell.font = Font(name=ui_font('黑体'), size=18, bold=True)
            title_cell.alignment = Alignment(horizontal='center', vertical='center')
            ws.row_dimensions[1].height = 30
            
            # 第二行：表头（将"文件标题"改为"摘要"）
            headers = ['序号', '编号', '收文编号', '收文日期', '来文单位', '来文字号', 
                      '密级', '摘要', '拟办意见', '领导批示', '办理结果', '收文途径']
            for col_num, header in enumerate(headers, 1):
                cell = ws.cell(row=2, column=col_num)
                cell.value = header
                cell.font = Font(name=ui_font('宋体'), size=11, bold=True)
                cell.alignment = Alignment(horizontal='center', vertical='center')
            
            # 设置列宽（调整列宽）
            ws.column_dimensions['A'].width = 6   # 序号
            ws.column_dimensions['B'].width = 8   # 编号
            ws.column_dimensions['C'].width = 12  # 收文编号
            ws.column_dimensions['D'].width = 12  # 收文日期
            ws.column_dimensions['E'].width = 15  # 来文单位（调小）
            ws.column_dimensions['F'].width = 30  # 来文字号（增加）
            ws.column_dimensions['G'].width = 8   # 密级
            ws.column_dimensions['H'].width = 35  # 摘要（加大）
            ws.column_dimensions['I'].width = 42  # 拟办意见（增加）
            ws.column_dimensions['J'].width = 20  # 领导批示（调小）
            ws.column_dimensions['K'].width = 25  # 办理结果（加大）
            ws.column_dimensions['L'].width = 12  # 收文途径
            
            # 数据行（从第三行开始）
            for row_num, record in enumerate(records, 3):
                # record是字典格式
                
                # 序号
                cell = ws.cell(row=row_num, column=1, value=row_num - 2)
                cell.font = Font(name=ui_font('宋体'), size=11)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                
                # 编号（留空）
                cell = ws.cell(row=row_num, column=2, value='')
                cell.font = Font(name=ui_font('宋体'), size=11)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                
                # 收文编号
                cell = ws.cell(row=row_num, column=3, value=record.get('收文编号', ''))
                cell.font = Font(name=ui_font('宋体'), size=11)
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
                
                # 收文日期
                cell = ws.cell(row=row_num, column=4, value=record.get('收文日期', ''))
                cell.font = Font(name=ui_font('宋体'), size=11)
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
                
                # 来文单位
                cell = ws.cell(row=row_num, column=5, value=record.get('来文单位', ''))
                cell.font = Font(name=ui_font('宋体'), size=11)
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
                
                # 来文字号 - 左对齐，自动换行
                cell = ws.cell(row=row_num, column=6, value=record.get('来文字号', ''))
                cell.font = Font(name=ui_font('宋体'), size=11)
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                
                # 密级（留空）
                cell = ws.cell(row=row_num, column=7, value='')
                cell.font = Font(name=ui_font('宋体'), size=11)
                cell.alignment = Alignment(horizontal='center', vertical='center')
                
                # 摘要（文件标题）- 左对齐，自动换行
                cell = ws.cell(row=row_num, column=8, value=record.get('文件标题', ''))
                cell.font = Font(name=ui_font('宋体'), size=11)
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                
                # 拟办意见（承办意见）- 左对齐，自动换行
                cell = ws.cell(row=row_num, column=9, value=record.get('承办意见', ''))
                cell.font = Font(name=ui_font('宋体'), size=11)
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                
                # 领导批示 - 左对齐，自动换行
                cell = ws.cell(row=row_num, column=10, value=record.get('领导批示', ''))
                cell.font = Font(name=ui_font('宋体'), size=11)
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                
                # 办理结果（办理情况）- 左对齐，自动换行
                cell = ws.cell(row=row_num, column=11, value=record.get('办理情况', ''))
                cell.font = Font(name=ui_font('宋体'), size=11)
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                
                # 收文途径
                cell = ws.cell(row=row_num, column=12, value=record.get('收文途径', ''))
                cell.font = Font(name=ui_font('宋体'), size=11)
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
            
            # 从第二行开始添加边框（包含表头）
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            for row in range(2, len(records) + 3):  # 从第二行开始
                for col in range(1, 13):
                    ws.cell(row=row, column=col).border = thin_border
            
            # 保存文件
            wb.save(file_path)
            
            messagebox.showinfo("成功", f"已导出 {len(records)} 条记录到:\n{file_path}")
            self.log(f"导出数据到Excel: {file_path}, 共 {len(records)} 条记录")
            
        except ImportError:
            messagebox.showerror("错误", "导出Excel功能需要安装openpyxl库\n请运行: pip install openpyxl")
            self.log("导出失败: 未安装openpyxl库")
        except Exception as e:
            messagebox.showerror("错误", f"导出失败: {e}")
            self.log(f"导出Excel失败: {e}")
            import traceback
            traceback.print_exc()
    
    def load_db_record_to_editor(self):
        """加载数据库记录到编辑界面（双击专用，不打开Word）"""
        selection = self.db_tree.selection()
        if not selection:
            return
        
        try:
            # 获取选中记录的收文编号
            values = self.db_tree.item(selection[0])['values']
            receipt_number = values[1]  # 序号列是0，收文编号是1
            
            # 从数据库获取完整记录
            record = self.db.get_complaint_by_receipt_number(receipt_number)
            if not record:
                messagebox.showerror("错误", "未找到该记录")
                return
            
            # 使用模板生成临时Word文档
            from docx import Document
            template_path = '投诉批办单示例#1.docx'
            if not os.path.exists(template_path):
                messagebox.showerror("错误", f"模板文件不存在: {template_path}")
                return
            
            doc = Document(template_path)
            if len(doc.tables) == 0:
                messagebox.showerror("错误", "模板文件中没有表格")
                return
            
            table = doc.tables[0]
            generator = WordGenerator()
            
            # 表格位置映射
            field_mapping = {
                '来文单位': (0, 2),
                '来文字号': (1, 2),
                '收文日期': (1, 9),
                '来文类型': (2, 2),
                '收文途径': (2, 9),
                '收文编号': (3, 9),
                '文件标题': (4, 2),
                '领导批示': (5, 2),
                '承办意见': (6, 2),
                '办理情况': (9, 2),
            }
            
            # 填充数据
            for field_name, (row, col) in field_mapping.items():
                value = record.get(field_name, '')
                if row < len(table.rows) and col < len(table.rows[row].cells):
                    cell = table.rows[row].cells[col]
                    
                    # 应用格式
                    if field_name in ['来文单位', '文件标题', '办理情况']:
                        generator.set_cell_text_keep_format(cell, value, add_prefix_space=True)
                    elif field_name == '来文字号':
                        # 按最长行字符数定字号，按行数定行距
                        lns = [l for l in value.split('\n') if l.strip()]
                        max_char = max((len(l) for l in lns), default=0)
                        line_count = len(lns)
                        if max_char >= 33:
                            font_size_pt = 9
                        elif max_char >= 29:
                            font_size_pt = 10.5
                        elif max_char >= 25:
                            font_size_pt = 12
                        elif line_count >= 4:
                            font_size_pt = 10.5
                        elif line_count == 3:
                            font_size_pt = 12
                        else:
                            font_size_pt = 14
                        if line_count >= 4:
                            line_spacing_pt = 10
                        elif line_count == 3:
                            line_spacing_pt = 11
                        else:
                            line_spacing_pt = 17
                        generator.set_cell_text_with_custom_format(
                            cell, value,
                            font_size_pt=font_size_pt,
                            line_spacing_pt=line_spacing_pt,
                            alignment='center'
                        )
                        self.log(f'来文字号格式：{font_size_pt}pt，行距{line_spacing_pt}磅（最长行{max_char}字）')
                    elif field_name in ['收文日期', '来文类型', '收文途径', '收文编号']:
                        generator.set_cell_text_keep_format(cell, value, alignment='center')
                    elif field_name == '承办意见':
                        # 拟办意见：行距28磅
                        # 检查是否已有前导空格，如果没有则添加（兼容旧数据）
                        if value and not value.startswith(' ') and not value.startswith('　'):
                            generator.set_cell_text_keep_format(cell, value, add_prefix_space=True, line_spacing_pt=28)
                        else:
                            generator.set_cell_text_keep_format(cell, value, line_spacing_pt=28)
                    else:
                        generator.set_cell_text_keep_format(cell, value)
            
            # 使用原始文件名保存到待打印目录
            original_filename = record.get('文件名', f'投诉批办单#{receipt_number}.docx')
            output_path = os.path.join(self.path_config['待打印'], original_filename)
            doc.save(output_path)
            
            self.log(f"已从数据库加载记录到编辑器: {original_filename}")
            
            # 刷新待打印列表
            self.refresh_file_list()
            
            # 切换到批办单编辑标签页（索引0）
            self.notebook.select(0)
            
            # 找到并选中该文件
            for item in self.file_tree.get_children():
                values = self.file_tree.item(item)['values']
                if len(values) >= 3 and values[2] == original_filename:  # 文件名在第3列（索引2）
                    self.file_tree.selection_set(item)
                    self.file_tree.see(item)
                    # 加载文件到编辑器
                    self.load_file_auto()
                    break
            
        except Exception as e:
            messagebox.showerror("错误", f"加载记录失败: {e}")
            self.log(f"从数据库加载记录失败: {e}")
            import traceback
            traceback.print_exc()
    
    def open_db_record_in_word(self):
        """打开数据库记录的Word批办单（按钮专用，只打开临时Word，不加载到编辑器）"""
        selection = self.db_tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选择要打开的记录")
            return
        
        try:
            # 获取选中记录的收文编号
            values = self.db_tree.item(selection[0])['values']
            receipt_number = values[1]  # 序号列是0，收文编号是1
            
            # 从数据库获取完整记录
            record = self.db.get_complaint_by_receipt_number(receipt_number)
            if not record:
                messagebox.showerror("错误", "未找到该记录")
                return
            
            # 使用模板生成临时Word文档
            from docx import Document
            import tempfile
            
            template_path = '投诉批办单示例#1.docx'
            if not os.path.exists(template_path):
                messagebox.showerror("错误", f"模板文件不存在: {template_path}")
                return
            
            doc = Document(template_path)
            if len(doc.tables) == 0:
                messagebox.showerror("错误", "模板文件中没有表格")
                return
            
            table = doc.tables[0]
            generator = WordGenerator()
            
            # 表格位置映射
            field_mapping = {
                '来文单位': (0, 2),
                '来文字号': (1, 2),
                '收文日期': (1, 9),
                '来文类型': (2, 2),
                '收文途径': (2, 9),
                '收文编号': (3, 9),
                '文件标题': (4, 2),
                '领导批示': (5, 2),
                '承办意见': (6, 2),
                '办理情况': (9, 2),
            }
            
            # 填充数据
            for field_name, (row, col) in field_mapping.items():
                value = record.get(field_name, '')
                if row < len(table.rows) and col < len(table.rows[row].cells):
                    cell = table.rows[row].cells[col]
                    
                    # 应用格式
                    if field_name in ['来文单位', '文件标题', '办理情况']:
                        generator.set_cell_text_keep_format(cell, value, add_prefix_space=True)
                    elif field_name in ['来文字号', '收文日期', '来文类型', '收文途径', '收文编号']:
                        generator.set_cell_text_keep_format(cell, value, alignment='center')
                    elif field_name == '承办意见':
                        generator.set_cell_text_keep_format(cell, value, line_spacing_pt=28)
                    else:
                        generator.set_cell_text_keep_format(cell, value)
            
            # 生成临时文件名（使用时间戳避免冲突）
            import time
            timestamp = int(time.time() * 1000)  # 毫秒级时间戳
            # 清理收文编号中的特殊字符，避免文件名问题
            safe_receipt = receipt_number.replace('/', '_').replace('\\', '_').replace(':', '_')
            temp_filename = f'投诉批办单_{safe_receipt}_{timestamp}.docx'
            temp_dir = tempfile.gettempdir()
            temp_path = os.path.join(temp_dir, temp_filename)
            
            # 如果文件已存在且被占用，尝试删除旧文件
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                    self.log(f"已删除旧的临时文件: {temp_filename}")
                except Exception as e:
                    # 文件被占用，使用新的时间戳生成新文件名
                    timestamp = int(time.time() * 1000) + 1
                    temp_filename = f'投诉批办单_{safe_receipt}_{timestamp}.docx'
                    temp_path = os.path.join(temp_dir, temp_filename)
                    self.log(f"旧文件被占用，使用新文件名: {temp_filename}")
            
            # 保存到临时目录
            doc.save(temp_path)
            
            self.log(f"已生成临时Word文档: {temp_filename}")
            
            # 打开Word文档
            try:
                if not open_path(temp_path):
                    raise RuntimeError("未找到可用的系统打开命令")
                self.log(f"已打开临时Word文档: {temp_filename}")
            except Exception as e:
                self.log(f"打开Word文档失败: {e}")
                messagebox.showerror("错误", f"打开Word文档失败: {e}")
            
        except Exception as e:
            messagebox.showerror("错误", f"打开记录失败: {e}")
            self.log(f"从数据库打开记录失败: {e}")
            import traceback
            traceback.print_exc()
    
    def generate_word_from_db(self):
        """从数据库记录生成Word文档"""
        selection = self.db_tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选择要生成Word文档的记录")
            return
        
        try:
            # 获取选中记录的收文编号
            values = self.db_tree.item(selection[0])['values']
            receipt_number = values[1]  # 序号列是0，收文编号是1
            
            # 从数据库获取完整记录
            record = self.db.get_complaint_by_receipt_number(receipt_number)
            if not record:
                messagebox.showerror("错误", "未找到该记录")
                return
            
            # 使用模板生成Word文档
            from docx import Document
            template_path = '投诉批办单示例#1.docx'
            if not os.path.exists(template_path):
                messagebox.showerror("错误", f"模板文件不存在: {template_path}")
                return
            
            doc = Document(template_path)
            if len(doc.tables) == 0:
                messagebox.showerror("错误", "模板文件中没有表格")
                return
            
            table = doc.tables[0]
            generator = WordGenerator()
            
            # 表格位置映射
            field_mapping = {
                '来文单位': (0, 2),
                '来文字号': (1, 2),
                '收文日期': (1, 9),
                '来文类型': (2, 2),
                '收文途径': (2, 9),
                '收文编号': (3, 9),
                '文件标题': (4, 2),
                '领导批示': (5, 2),
                '承办意见': (6, 2),
                '办理情况': (9, 2),
            }
            
            # 填充数据
            for field_name, (row, col) in field_mapping.items():
                value = record.get(field_name, '')
                if row < len(table.rows) and col < len(table.rows[row].cells):
                    cell = table.rows[row].cells[col]
                    
                    # 应用格式
                    if field_name in ['来文单位', '文件标题', '办理情况']:
                        generator.set_cell_text_keep_format(cell, value, add_prefix_space=True)
                    elif field_name in ['来文字号', '收文日期', '来文类型', '收文途径', '收文编号']:
                        generator.set_cell_text_keep_format(cell, value, alignment='center')
                    elif field_name == '承办意见':
                        generator.set_cell_text_keep_format(cell, value, line_spacing_pt=28)
                    else:
                        generator.set_cell_text_keep_format(cell, value)
            
            # 保存文档
            output_filename = f"{record.get('文件名', receipt_number)}.docx"
            output_path = os.path.join(self.path_config['待打印'], output_filename)
            doc.save(output_path)
            
            self.log(f"已从数据库生成Word文档: {output_filename}")
            messagebox.showinfo("成功", f"Word文档已生成！\n\n文件名: {output_filename}\n保存位置: {self.path_config['待打印']}")
            
            # 刷新待打印列表
            self.refresh_file_list()
            
        except Exception as e:
            messagebox.showerror("错误", f"生成Word文档失败: {e}")
            self.log(f"从数据库生成Word失败: {e}")
            import traceback
            traceback.print_exc()
    
    def view_db_detail(self):
        """查看数据库记录详情"""
        selection = self.db_tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选择要查看的记录")
            return
        
        try:
            values = self.db_tree.item(selection[0])['values']
            receipt_number = values[1]  # 序号列是0，收文编号是1
            
            record = self.db.get_complaint_by_receipt_number(receipt_number)
            if not record:
                messagebox.showerror("错误", "未找到该记录")
                return
            
            # 创建详情窗口
            detail_win = tk.Toplevel(self.root)
            detail_win.title(f"详情 - {receipt_number}")
            detail_win.geometry("600x700")
            
            # 创建滚动文本框
            text = scrolledtext.ScrolledText(detail_win, wrap=tk.WORD, font=(ui_font('微软雅黑'), 10))
            text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            # 显示详情
            detail_text = f"""收文编号: {record.get('收文编号', '')}
来文单位: {record.get('来文单位', '')}
来文字号: {record.get('来文字号', '')}
收文日期: {record.get('收文日期', '')}
来文类型: {record.get('来文类型', '')}
收文途径: {record.get('收文途径', '')}
文件标题: {record.get('文件标题', '')}

领导批示:
{record.get('领导批示', '')}

承办意见:
{record.get('承办意见', '')}

办理情况:
{record.get('办理情况', '')}

---
文件名: {record.get('文件名', '')}
创建时间: {record.get('创建时间', '')}
更新时间: {record.get('更新时间', '')}
"""
            text.insert(1.0, detail_text)
            text.config(state=tk.DISABLED)
            
        except Exception as e:
            messagebox.showerror("错误", f"查看详情失败: {e}")
    
    def delete_db_record(self):
        """删除数据库记录"""
        selection = self.db_tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选择要删除的记录")
            return
        
        try:
            values = self.db_tree.item(selection[0])['values']
            receipt_number = values[1]  # 序号列是0，收文编号是1
            
            if not messagebox.askyesno("确认删除", f"确定要删除记录 {receipt_number} 吗？\n\n注意：这只会删除数据库记录，不会删除实际的Word文件。"):
                return
            
            if self.db.delete_complaint(receipt_number):
                self.log(f"已删除数据库记录: {receipt_number}")
                # 刷新列表
                self.search_database()
                messagebox.showinfo("成功", "记录已删除")
            else:
                messagebox.showerror("错误", "删除失败")
                
        except Exception as e:
            messagebox.showerror("错误", f"删除记录失败: {e}")
            
    def show_help(self):
        """显示帮助"""
        help_text = """
投诉批办单管理系统 使用说明

1. 批量处理
   - 点击"批量处理"按钮处理程序目录中的PDF和Word文件
   - 生成的文件自动保存在"待打印"目录

2. 编辑文件
   - 双击文件列表中的文件进行编辑
   - 在右侧面板修改内容后点击"保存"

3. 快捷填充
   - 切换到"快捷填充"标签页
   - 点击预设按钮快速填充常用内容

4. 打印功能
   - 打印当前文件：打印正在编辑的文件
   - 批量打印选中：打印文件列表中选中的文件
   - 批量打印全部：打印所有文件

5. 其他功能
   - 编号管理：修改文件和收文编号起始值
   - 统计报表：查看和导出文件统计信息
        """
        
        dialog = tk.Toplevel(self.root)
        dialog.title("使用说明")
        dialog.geometry("600x500")
        
        text = scrolledtext.ScrolledText(dialog, wrap=tk.WORD, font=(ui_font('微软雅黑'), 10))
        text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        text.insert(1.0, help_text)
        text.config(state=tk.DISABLED)
        
    def show_about(self):
        """显示关于"""
        about_text = """
投诉批办单管理系统 v1.38.2

功能特点：
• 批量处理PDF和Word文件
• 数据库检索和管理
• 批办单编辑和打印
• 文件合并和重新编号
• 模板化处理流程
• 二维码智能生成
• 待处理文件快速搜索（新）

技术栈：
• Python 3.x
• tkinter / ttkbootstrap
• python-docx
• pdfplumber

版本：v1.38.2
日期：2026-07-12
        """
        
        messagebox.showinfo("关于", about_text)
    
    # ==================== 数据库管理功能 ====================
    
    def show_database_stats(self):
        """显示数据库统计信息"""
        try:
            import sqlite3
            from datetime import datetime
            
            db_path = 'complaints.db'
            if not os.path.exists(db_path):
                messagebox.showwarning("提示", "数据库文件不存在")
                return
            
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # 获取总记录数
            cursor.execute("SELECT COUNT(*) FROM complaints")
            total_count = cursor.fetchone()[0]
            
            # 获取文件大小
            file_size = os.path.getsize(db_path)
            size_mb = file_size / (1024 * 1024)
            
            # 获取最早和最新的记录
            cursor.execute("SELECT MIN(收文日期), MAX(收文日期) FROM complaints WHERE 收文日期 != ''")
            date_range = cursor.fetchone()
            earliest_date = date_range[0] if date_range[0] else "无"
            latest_date = date_range[1] if date_range[1] else "无"
            
            # 获取不同来文单位的数量
            cursor.execute("SELECT COUNT(DISTINCT 来文单位) FROM complaints WHERE 来文单位 != ''")
            unit_count = cursor.fetchone()[0]
            
            # 获取不同来文类型的统计
            cursor.execute("SELECT 来文类型, COUNT(*) FROM complaints WHERE 来文类型 != '' GROUP BY 来文类型")
            type_stats = cursor.fetchall()
            type_info = "\n".join([f"  • {t[0]}: {t[1]}件" for t in type_stats]) if type_stats else "  无"
            
            conn.close()
            
            stats_text = f"""
📊 数据库统计信息
{'=' * 40}

📁 文件信息：
  • 文件路径：{db_path}
  • 文件大小：{size_mb:.2f} MB ({file_size:,} 字节)

📋 记录统计：
  • 总记录数：{total_count} 条
  • 来文单位数：{unit_count} 个
  
📅 日期范围：
  • 最早记录：{earliest_date}
  • 最新记录：{latest_date}

📝 来文类型统计：
{type_info}

⏰ 统计时间：{datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
            """
            
            # 创建对话框
            dialog = tk.Toplevel(self.root)
            dialog.title("数据库统计信息")
            dialog.geometry("500x550")
            dialog.resizable(False, False)
            
            # 文本区域
            text_area = scrolledtext.ScrolledText(dialog, wrap=tk.WORD, font=(ui_font('微软雅黑'), 10))
            text_area.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            text_area.insert(1.0, stats_text)
            text_area.config(state=tk.DISABLED)
            
            # 按钮框架
            btn_frame = ttk.Frame(dialog)
            btn_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
            
            ttk.Button(btn_frame, text="关闭", command=dialog.destroy, 
                      bootstyle="secondary", width=15).pack(side=tk.RIGHT)
            
            self.log("已显示数据库统计信息")
            
        except Exception as e:
            messagebox.showerror("错误", f"获取数据库统计信息失败：{e}")
            self.log(f"获取数据库统计信息失败：{e}")
    
    def backup_database(self):
        """数据库备份"""
        try:
            from datetime import datetime
            import shutil
            
            db_path = 'complaints.db'
            if not os.path.exists(db_path):
                messagebox.showwarning("提示", "数据库文件不存在")
                return
            
            # 生成默认备份文件名（带时间戳）
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            default_filename = f"complaints_backup_{timestamp}.db"
            
            # 选择保存位置
            backup_path = filedialog.asksaveasfilename(
                title="选择备份保存位置",
                defaultextension=".db",
                initialfile=default_filename,
                filetypes=[("数据库文件", "*.db"), ("所有文件", "*.*")]
            )
            
            if not backup_path:
                return
            
            # 复制数据库文件
            shutil.copy2(db_path, backup_path)
            
            file_size = os.path.getsize(backup_path)
            size_mb = file_size / (1024 * 1024)
            
            messagebox.showinfo("成功", 
                              f"数据库备份成功！\n\n"
                              f"备份位置：{backup_path}\n"
                              f"文件大小：{size_mb:.2f} MB")
            self.log(f"数据库已备份至：{backup_path}")
            
        except Exception as e:
            messagebox.showerror("错误", f"数据库备份失败：{e}")
            self.log(f"数据库备份失败：{e}")
    
    def restore_database(self):
        """数据库恢复"""
        try:
            import shutil
            
            # 选择备份文件
            backup_path = filedialog.askopenfilename(
                title="选择要恢复的备份文件",
                filetypes=[("数据库文件", "*.db"), ("所有文件", "*.*")]
            )
            
            if not backup_path:
                return
            
            # 检查备份文件的记录数
            backup_count = 0
            try:
                import sqlite3
                backup_conn = sqlite3.connect(backup_path)
                backup_cursor = backup_conn.cursor()
                backup_cursor.execute("SELECT COUNT(*) FROM complaints")
                backup_count = backup_cursor.fetchone()[0]
                backup_conn.close()
                self.log(f"备份文件包含 {backup_count} 条记录")
                
                if backup_count == 0:
                    if not messagebox.askyesno("警告", 
                                              f"备份文件中没有记录！\n\n"
                                              f"确定要从空备份恢复吗？",
                                              icon='warning'):
                        self.log("用户取消了从空备份恢复")
                        return
            except Exception as e:
                messagebox.showerror("错误", f"无法读取备份文件：{e}\n\n可能不是有效的数据库文件。")
                self.log(f"读取备份文件失败：{e}")
                return
            
            # 确认操作
            if not messagebox.askyesno("确认恢复", 
                                      "⚠️ 警告：恢复数据库将覆盖当前所有数据！\n\n"
                                      "建议先备份当前数据库。\n\n"
                                      "确定要继续恢复吗？",
                                      icon='warning'):
                return
            
            db_path = 'complaints.db'
            
            # 先关闭当前数据库连接，释放文件句柄
            try:
                del self.db
                import gc
                gc.collect()  # 强制垃圾回收，确保连接关闭
                import time
                time.sleep(0.2)  # 等待一小段时间，确保文件解锁
            except:
                pass
            
            # 备份当前数据库
            if os.path.exists(db_path):
                from datetime import datetime
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                auto_backup = f"complaints_before_restore_{timestamp}.db"
                shutil.copy2(db_path, auto_backup)
                self.log(f"当前数据库已自动备份为：{auto_backup}")
            
            # 恢复数据库
            shutil.copy2(backup_path, db_path)
            self.log(f"已从 {backup_path} 复制到 {db_path}")
            
            # 重新初始化数据库连接
            self.db = ComplaintDatabase()
            self.log(f"数据库连接已重新初始化")
            
            # 验证恢复后的记录数
            restored_count = 0
            try:
                results = self.db.search_complaints(limit=100000)
                restored_count = len(results)
                self.log(f"恢复后数据库包含 {restored_count} 条记录")
            except Exception as e:
                self.log(f"警告：无法统计恢复后的记录数：{e}")
            
            messagebox.showinfo("成功", 
                              f"数据库恢复成功！\n\n"
                              f"从备份文件：{os.path.basename(backup_path)}\n"
                              f"备份记录数：{backup_count} 条\n"
                              f"恢复记录数：{restored_count} 条\n\n"
                              f"当前数据库已自动备份")
            self.log(f"数据库已从备份恢复：{backup_path}")
            
            # 刷新界面
            self.refresh_all_lists()
            
        except Exception as e:
            messagebox.showerror("错误", f"数据库恢复失败：{e}")
            self.log(f"数据库恢复失败：{e}")
    
    def export_database_to_excel(self):
        """导出数据库为Excel"""
        try:
            from datetime import datetime
            
            # 获取所有记录
            results = self.db.search_complaints(limit=100000)  # 获取所有记录
            
            if not results:
                messagebox.showinfo("提示", "数据库中没有记录")
                return
            
            # 生成默认文件名
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            default_filename = f"投诉批办单数据_{timestamp}.xlsx"
            
            # 选择保存位置
            excel_path = filedialog.asksaveasfilename(
                title="选择Excel保存位置",
                defaultextension=".xlsx",
                initialfile=default_filename,
                filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
            )
            
            if not excel_path:
                return
            
            # 导入openpyxl（如果没有安装，提示用户）
            try:
                from openpyxl import Workbook
                from openpyxl.styles import Font, Alignment, PatternFill
            except ImportError:
                messagebox.showerror("错误", 
                                   "缺少openpyxl模块！\n\n"
                                   "请在命令行执行：\n"
                                   "pip install openpyxl")
                return
            
            # 创建Excel工作簿
            wb = Workbook()
            ws = wb.active
            ws.title = "投诉批办单数据"
            
            # 定义表头
            headers = ['序号', '收文编号', '来文单位', '来文字号', '收文日期', '来文类型', 
                      '收文途径', '文件标题', '领导批示', '承办意见', '办理情况', '文件名']
            
            # 写入表头（带样式）
            header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF", size=11)
            header_alignment = Alignment(horizontal="center", vertical="center")
            
            for col, header in enumerate(headers, start=1):
                cell = ws.cell(row=1, column=col, value=header)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = header_alignment
            
            # 写入数据
            for idx, record in enumerate(results, start=2):
                ws.cell(row=idx, column=1, value=idx-1)  # 序号
                ws.cell(row=idx, column=2, value=record.get('收文编号', ''))
                ws.cell(row=idx, column=3, value=record.get('来文单位', ''))
                ws.cell(row=idx, column=4, value=record.get('来文字号', ''))
                ws.cell(row=idx, column=5, value=record.get('收文日期', ''))
                ws.cell(row=idx, column=6, value=record.get('来文类型', ''))
                ws.cell(row=idx, column=7, value=record.get('收文途径', ''))
                ws.cell(row=idx, column=8, value=record.get('文件标题', ''))
                ws.cell(row=idx, column=9, value=record.get('领导批示', ''))
                ws.cell(row=idx, column=10, value=record.get('承办意见', ''))
                ws.cell(row=idx, column=11, value=record.get('办理情况', ''))
                ws.cell(row=idx, column=12, value=record.get('文件名', ''))
            
            # 调整列宽
            ws.column_dimensions['A'].width = 8
            ws.column_dimensions['B'].width = 12
            ws.column_dimensions['C'].width = 20
            ws.column_dimensions['D'].width = 15
            ws.column_dimensions['E'].width = 12
            ws.column_dimensions['F'].width = 10
            ws.column_dimensions['G'].width = 12
            ws.column_dimensions['H'].width = 30
            ws.column_dimensions['I'].width = 20
            ws.column_dimensions['J'].width = 30
            ws.column_dimensions['K'].width = 30
            ws.column_dimensions['L'].width = 25
            
            # 保存Excel文件
            wb.save(excel_path)
            
            messagebox.showinfo("成功", 
                              f"数据导出成功！\n\n"
                              f"导出位置：{excel_path}\n"
                              f"记录数量：{len(results)} 条")
            self.log(f"数据已导出至Excel：{excel_path}，共{len(results)}条记录")
            
        except Exception as e:
            messagebox.showerror("错误", f"导出Excel失败：{e}")
            self.log(f"导出Excel失败：{e}")
            import traceback
            traceback.print_exc()
    
    def optimize_database(self):
        """数据库优化（VACUUM）"""
        try:
            import sqlite3
            
            db_path = 'complaints.db'
            if not os.path.exists(db_path):
                messagebox.showwarning("提示", "数据库文件不存在")
                return
            
            # 获取优化前的大小
            size_before = os.path.getsize(db_path)
            size_before_mb = size_before / (1024 * 1024)
            
            # 执行VACUUM优化
            conn = sqlite3.connect(db_path)
            conn.execute("VACUUM")
            conn.close()
            
            # 获取优化后的大小
            size_after = os.path.getsize(db_path)
            size_after_mb = size_after / (1024 * 1024)
            saved_mb = (size_before - size_after) / (1024 * 1024)
            
            messagebox.showinfo("优化完成", 
                              f"数据库优化完成！\n\n"
                              f"优化前：{size_before_mb:.2f} MB\n"
                              f"优化后：{size_after_mb:.2f} MB\n"
                              f"节省空间：{saved_mb:.2f} MB")
            self.log(f"数据库优化完成，节省空间：{saved_mb:.2f} MB")
            
        except Exception as e:
            messagebox.showerror("错误", f"数据库优化失败：{e}")
            self.log(f"数据库优化失败：{e}")
    
    def check_database_integrity(self):
        """数据库完整性检查"""
        try:
            import sqlite3
            
            db_path = 'complaints.db'
            if not os.path.exists(db_path):
                messagebox.showwarning("提示", "数据库文件不存在")
                return
            
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # 执行完整性检查
            cursor.execute("PRAGMA integrity_check")
            result = cursor.fetchone()[0]
            
            conn.close()
            
            if result == "ok":
                messagebox.showinfo("检查完成", 
                                  "✅ 数据库完整性检查通过！\n\n"
                                  "数据库文件完好无损。")
                self.log("数据库完整性检查通过")
            else:
                messagebox.showwarning("警告", 
                                     f"⚠️ 数据库完整性检查发现问题：\n\n{result}\n\n"
                                     "建议尽快备份数据并重新创建数据库。")
                self.log(f"数据库完整性检查发现问题：{result}")
            
        except Exception as e:
            messagebox.showerror("错误", f"数据库完整性检查失败：{e}")
            self.log(f"数据库完整性检查失败：{e}")
    
    def auto_backup_settings(self):
        """自动备份设置"""
        try:
            # 创建设置对话框
            dialog = tk.Toplevel(self.root)
            dialog.title("自动备份设置")
            dialog.geometry("450x300")
            dialog.resizable(False, False)
            
            # 主框架
            main_frame = ttk.Frame(dialog, padding=20)
            main_frame.pack(fill=tk.BOTH, expand=True)
            
            # 标题
            ttk.Label(main_frame, text="⚙️ 自动备份设置", 
                     font=(ui_font('微软雅黑'), 12, 'bold')).pack(pady=(0, 20))
            
            # 启用自动备份
            enable_var = tk.BooleanVar(value=False)
            ttk.Checkbutton(main_frame, text="启用自动备份", 
                          variable=enable_var, bootstyle="success-round-toggle").pack(anchor=tk.W, pady=10)
            
            # 备份频率
            freq_frame = ttk.Frame(main_frame)
            freq_frame.pack(fill=tk.X, pady=10)
            ttk.Label(freq_frame, text="备份频率：").pack(side=tk.LEFT)
            freq_var = tk.StringVar(value="每天")
            freq_combo = ttk.Combobox(freq_frame, textvariable=freq_var, 
                                     values=["每天", "每周", "每月"], state="readonly", width=15)
            freq_combo.pack(side=tk.LEFT, padx=10)
            
            # 备份位置
            path_frame = ttk.Frame(main_frame)
            path_frame.pack(fill=tk.X, pady=10)
            ttk.Label(path_frame, text="备份位置：").pack(side=tk.LEFT)
            path_var = tk.StringVar(value="./backup")
            path_entry = ttk.Entry(path_frame, textvariable=path_var, width=25)
            path_entry.pack(side=tk.LEFT, padx=10)
            
            def choose_backup_dir():
                dir_path = filedialog.askdirectory(title="选择备份目录")
                if dir_path:
                    path_var.set(dir_path)
            
            ttk.Button(path_frame, text="浏览", command=choose_backup_dir, 
                      bootstyle="secondary", width=8).pack(side=tk.LEFT)
            
            # 保留份数
            keep_frame = ttk.Frame(main_frame)
            keep_frame.pack(fill=tk.X, pady=10)
            ttk.Label(keep_frame, text="保留备份数：").pack(side=tk.LEFT)
            keep_var = tk.IntVar(value=7)
            ttk.Spinbox(keep_frame, from_=1, to=30, textvariable=keep_var, width=15).pack(side=tk.LEFT, padx=10)
            ttk.Label(keep_frame, text="份").pack(side=tk.LEFT)
            
            # 提示信息
            info_text = "注意：此功能需要程序持续运行才能生效\n建议使用系统任务计划程序实现定时备份"
            ttk.Label(main_frame, text=info_text, foreground="gray", 
                     font=(ui_font('微软雅黑'), 9)).pack(pady=20)
            
            # 按钮
            btn_frame = ttk.Frame(main_frame)
            btn_frame.pack(fill=tk.X, pady=(20, 0))
            
            def save_settings():
                # TODO: 保存设置到配置文件
                messagebox.showinfo("提示", 
                                  "自动备份设置已保存！\n\n"
                                  "注意：此功能正在开发中，暂未完全实现。\n"
                                  "建议使用Windows任务计划程序定期手动备份。")
                dialog.destroy()
            
            ttk.Button(btn_frame, text="保存", command=save_settings, 
                      bootstyle="success", width=15).pack(side=tk.RIGHT, padx=5)
            ttk.Button(btn_frame, text="取消", command=dialog.destroy, 
                      bootstyle="secondary", width=15).pack(side=tk.RIGHT)
            
        except Exception as e:
            messagebox.showerror("错误", f"打开自动备份设置失败：{e}")
    
    def initialize_database(self):
        """数据库初始化（清空所有记录）"""
        try:
            import sqlite3
            
            db_path = 'complaints.db'
            if not os.path.exists(db_path):
                messagebox.showwarning("提示", "数据库文件不存在")
                return
            
            # 第一次确认
            if not messagebox.askyesno("⚠️ 危险操作", 
                                      "⚠️⚠️⚠️ 警告 ⚠️⚠️⚠️\n\n"
                                      "数据库初始化将删除所有记录！\n\n"
                                      "此操作不可恢复！\n\n"
                                      "确定要继续吗？",
                                      icon='warning'):
                return
            
            # 第二次确认（输入确认）
            confirm_dialog = tk.Toplevel(self.root)
            confirm_dialog.title("确认操作")
            confirm_dialog.geometry("400x200")
            confirm_dialog.resizable(False, False)
            
            frame = ttk.Frame(confirm_dialog, padding=20)
            frame.pack(fill=tk.BOTH, expand=True)
            
            ttk.Label(frame, text="⚠️ 最终确认", 
                     font=(ui_font('微软雅黑'), 12, 'bold'), 
                     foreground="red").pack(pady=(0, 20))
            
            ttk.Label(frame, text="请输入 \"确认清空\" 来继续操作：", 
                     wraplength=350).pack(pady=10)
            
            confirm_entry = ttk.Entry(frame, width=30)
            confirm_entry.pack(pady=10)
            confirm_entry.focus()
            
            result = {'confirmed': False}
            
            def do_initialize():
                if confirm_entry.get() == "确认清空":
                    result['confirmed'] = True
                    confirm_dialog.destroy()
                else:
                    messagebox.showerror("错误", "输入不正确！操作已取消。")
            
            btn_frame = ttk.Frame(frame)
            btn_frame.pack(pady=20)
            ttk.Button(btn_frame, text="确认", command=do_initialize, 
                      bootstyle="danger", width=12).pack(side=tk.LEFT, padx=5)
            ttk.Button(btn_frame, text="取消", command=confirm_dialog.destroy, 
                      bootstyle="secondary", width=12).pack(side=tk.LEFT, padx=5)
            
            # 绑定回车键
            confirm_entry.bind('<Return>', lambda e: do_initialize())
            
            # 等待对话框关闭
            confirm_dialog.wait_window()
            
            if not result['confirmed']:
                self.log("数据库初始化已取消")
                return
            
            # 自动备份当前数据库
            from datetime import datetime
            import shutil
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_path = f"complaints_before_init_{timestamp}.db"
            shutil.copy2(db_path, backup_path)
            self.log(f"初始化前已自动备份：{backup_path}")
            
            # 清空所有记录
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM complaints")
            conn.commit()
            conn.close()
            
            # 重新初始化数据库连接
            self.db = ComplaintDatabase()
            
            messagebox.showinfo("完成", 
                              f"数据库已初始化！\n\n"
                              f"所有记录已清空。\n"
                              f"原数据已备份至：{backup_path}")
            self.log("数据库已初始化，所有记录已清空")
            
            # 刷新界面
            self.refresh_all_lists()
            
        except Exception as e:
            messagebox.showerror("错误", f"数据库初始化失败：{e}")
            self.log(f"数据库初始化失败：{e}")
    
    def search_by_reference_number(self):
        """根据来文字号查询数据库中的记录（支持部分匹配）"""
        try:
            # 获取来文字号输入框的内容
            ref_number = self.field_entries['来文字号'].get().strip()

            if not ref_number:
                messagebox.showwarning("提示", "请先输入来文字号")
                return

            self.log(f"开始查询来文字号: '{ref_number}'")

            # 在数据库中搜索包含该字号的记录
            all_records = self.db.search_complaints(keyword=ref_number, field='来文字号')

            # 获取当前日期，用于排除当日记录
            from datetime import datetime
            today = datetime.now().date()

            # 过滤掉当日的记录
            records = []
            for rec in all_records:
                rec_date_str = rec.get('收文日期', '').strip()

                # 尝试解析记录的日期
                try:
                    # 尝试多种日期格式
                    rec_date = None
                    for date_format in ['%Y-%m-%d', '%Y-%#m-%#d', '%Y-%m-%d', '%Y/%m/%d', '%Y/%#m/%#d']:
                        try:
                            rec_date = datetime.strptime(rec_date_str, date_format).date()
                            break
                        except:
                            continue

                    # 如果解析失败，尝试手动解析
                    if rec_date is None and rec_date_str:
                        parts = rec_date_str.replace('/', '-').split('-')
                        if len(parts) == 3:
                            rec_date = datetime(int(parts[0]), int(parts[1]), int(parts[2])).date()

                    # 只保留收文日期不是今天的记录
                    if rec_date != today:
                        records.append(rec)
                    else:
                        self.log(f"排除当日记录：收文编号={rec.get('收文编号', '')}, 收文日期={rec_date_str}")
                except:
                    # 如果日期解析失败，保留该记录（避免过度过滤）
                    records.append(rec)

            self.log(f"查询结果数量: {len(all_records)} 条，排除当日后: {len(records)} 条")
            
            # 调试：输出前3条记录的来文字号
            if records:
                for i, rec in enumerate(records[:3]):
                    ref_in_db = rec.get('来文字号', '')
                    # 检查是否真正包含搜索关键字
                    contains = ref_number in ref_in_db
                    self.log(f"结果{i+1}: 来文字号='{ref_in_db}' (长度:{len(ref_in_db)}, 包含查询字符:{contains})")
            
            if not records:
                messagebox.showinfo("查询结果", f"未找到包含 '{ref_number}' 的记录")
                return
            
            # 如果找到记录，显示查询结果窗口
            self.show_reference_search_results(records, ref_number)
            
        except Exception as e:
            messagebox.showerror("错误", f"查询失败: {e}")
            self.log(f"来文字号查询失败: {e}")
            import traceback
            traceback.print_exc()
    
    def show_reference_search_results(self, records, search_keyword):
        """显示来文字号查询结果窗口（支持多条记录切换）"""
        # 创建弹出窗口
        result_window = tk.Toplevel(self.root)
        result_window.title(f"查询结果 - 找到 {len(records)} 条记录")
        
        # 计算窗口位置：放在屏幕右半部分（编辑区域位置）
        window_width = 850
        window_height = 700
        
        # 获取主窗口位置
        root_x = self.root.winfo_x()
        root_y = self.root.winfo_y()
        root_width = self.root.winfo_width()
        root_height = self.root.winfo_height()
        
        # 计算右侧位置（主窗口右半部分）
        x = root_x + root_width // 2 + 50
        y = root_y + 50
        
        result_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
        result_window.transient(self.root)
        result_window.grab_set()
        
        # 当前显示的记录索引
        current_index = [0]  # 使用列表以便在闭包中修改
        
        def show_record(index):
            """显示指定索引的记录（使用表格形式，支持多行）"""
            record = records[index]
            
            # 清空并重新启用文本框
            detail_text.config(state='normal')
            detail_text.delete(1.0, tk.END)
            
            # 删除所有已存在的tag
            for tag in detail_text.tag_names():
                detail_text.tag_delete(tag)
            
            # 定义字段及其显示顺序
            fields = [
                ('收文编号', record.get('收文编号', '')),
                ('来文单位', record.get('来文单位', '')),
                ('来文字号', record.get('来文字号', '')),
                ('收文日期', record.get('收文日期', '')),
                ('来文类型', record.get('来文类型', '')),
                ('收文途径', record.get('收文途径', '')),
                ('文件标题', record.get('文件标题', '')),
                ('领导批示', record.get('领导批示', '')),
                ('承办意见', record.get('承办意见', '')),
                ('办理情况', record.get('办理情况', ''))
            ]
            
            # 配置文本标签样式
            detail_text.tag_configure('field_name', font=(ui_font('微软雅黑'), 12, 'bold'), 
                                     foreground='#2c3e50')
            detail_text.tag_configure('field_value', font=(ui_font('仿宋_GB2312'), 12), 
                                     foreground='#34495e')
            detail_text.tag_configure('separator', font=(ui_font('微软雅黑'), 8), 
                                     foreground='#bdc3c7')
            
            # 插入表格数据
            for i, (field_name, field_value) in enumerate(fields):
                if i > 0:
                    # 添加分隔线
                    detail_text.insert(tk.END, '─' * 100 + '\n', 'separator')
                
                # 格式化字段名（固定宽度，右对齐）
                formatted_field_name = f'{field_name:>8}:  '
                detail_text.insert(tk.END, formatted_field_name, 'field_name')
                
                # 插入字段值（处理多行内容）
                if field_value:
                    # 处理多行内容
                    lines = str(field_value).split('\n')
                    for line_idx, line in enumerate(lines):
                        if line_idx > 0:
                            # 后续行需要缩进对齐
                            detail_text.insert(tk.END, '\n' + ' ' * 12)
                        detail_text.insert(tk.END, line, 'field_value')
                
                detail_text.insert(tk.END, '\n')
            
            # 禁用文本框编辑（但仍可选择和复制）
            detail_text.config(state='disabled')
            
            # 更新按钮状态
            prev_btn.config(state='normal' if index > 0 else 'disabled')
            next_btn.config(state='normal' if index < len(records) - 1 else 'disabled')
            counter_label.config(text=f"{index + 1} / {len(records)}")
        
        def show_prev():
            """显示上一条记录"""
            if current_index[0] > 0:
                current_index[0] -= 1
                show_record(current_index[0])
        
        def show_next():
            """显示下一条记录"""
            if current_index[0] < len(records) - 1:
                current_index[0] += 1
                show_record(current_index[0])
        
        # 顶部信息栏
        top_frame = ttk.Frame(result_window)
        top_frame.pack(fill=tk.X, padx=15, pady=10)
        
        ttk.Label(top_frame, text=f"搜索关键字: {search_keyword}", 
                 font=(ui_font('微软雅黑'), 11, 'bold')).pack(side=tk.LEFT)
        
        ttk.Label(top_frame, text=f"共找到 {len(records)} 条记录", 
                 font=(ui_font('微软雅黑'), 11), foreground='#27ae60').pack(side=tk.RIGHT)
        
        # 导航按钮区域
        nav_frame = ttk.Frame(result_window)
        nav_frame.pack(fill=tk.X, padx=15, pady=8)
        
        prev_btn = ttk.Button(nav_frame, text="◀ 上一条", command=show_prev, 
                             bootstyle="info-outline", width=12)
        prev_btn.pack(side=tk.LEFT, padx=5)
        
        counter_label = ttk.Label(nav_frame, text="1 / 1", 
                                 font=(ui_font('微软雅黑'), 12, 'bold'))
        counter_label.pack(side=tk.LEFT, expand=True)
        
        next_btn = ttk.Button(nav_frame, text="下一条 ▶", command=show_next, 
                             bootstyle="info-outline", width=12)
        next_btn.pack(side=tk.RIGHT, padx=5)
        
        # 表格显示区域
        detail_frame = ttk.Frame(result_window)
        detail_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
        
        # 创建Text组件（支持多行显示和文本选择）
        detail_text = tk.Text(detail_frame, wrap=tk.WORD, 
                             font=(ui_font('仿宋_GB2312'), 12),
                             relief=tk.SOLID, borderwidth=2,
                             padx=10, pady=10,
                             bg='#f8f9fa')
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(detail_frame, orient='vertical', command=detail_text.yview)
        detail_text.configure(yscrollcommand=scrollbar.set)
        
        # 布局
        detail_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 底部按钮
        bottom_frame = ttk.Frame(result_window)
        bottom_frame.pack(fill=tk.X, padx=15, pady=10)
        
        ttk.Button(bottom_frame, text="关闭", command=result_window.destroy,
                  bootstyle="secondary", width=12).pack(side=tk.RIGHT)
        
        # 显示第一条记录
        show_record(0)


def main():
    """主函数"""
    # 设置日志记录（在程序启动时立即开始记录，即使闪退也能看到日志）
    import logging
    import sys

    # 跨平台启动引导：切到应用目录、释放默认配置
    pc.ensure_app_cwd()
    pc.seed_user_data()

    # 命令行批处理模式：启动命令行.sh 传 --cli，等价于 Windows 的 python main.py
    if '--cli' in sys.argv[1:]:
        sys.argv.remove('--cli')
        import main as _cli
        _cli.main()
        return

    # 使用固定的日志文件名（每次启动自动覆盖）
    log_filename = get_appdata_path('app_startup.log')
    _handlers = [logging.FileHandler(log_filename, mode='w', encoding='utf-8')]
    # PyInstaller console=False 时 sys.stdout / sys.stderr 均为 None，
    # 不能把 None 交给 StreamHandler，否则日志初始化就会抛异常且看不到任何提示。
    _stream = sys.stdout or sys.stderr
    if _stream is not None:
        _handlers.append(logging.StreamHandler(_stream))
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=_handlers
    )
    
    try:
        logging.info("=" * 60)
        logging.info("投诉批办单管理系统 v1.38.2 启动中...")
        logging.info("=" * 60)
        logging.info(f"Python版本: {sys.version}")
        logging.info(f"平台: {sys.platform}")
        logging.info(f"日志文件: {log_filename}")
        
        # 检查系统依赖
        logging.info("检查系统依赖...")
        try:
            import tkinter
            logging.info("✓ tkinter可用")
        except Exception as e:
            logging.error(f"✗ tkinter不可用: {e}")
            raise
        
        try:
            import ttkbootstrap
            logging.info(f"✓ ttkbootstrap可用 (版本: {ttkbootstrap.__version__ if hasattr(ttkbootstrap, '__version__') else '未知'})")
        except Exception as e:
            logging.error(f"✗ ttkbootstrap不可用: {e}")
            raise
        
        # 延迟隐藏控制台窗口（确保日志已经写入）
        def hide_console_delayed():
            """延迟隐藏控制台，确保日志文件创建成功"""
            import time
            time.sleep(0.5)  # 等待日志写入
            try:
                import ctypes
                if IS_WINDOWS:
                    whnd = ctypes.windll.kernel32.GetConsoleWindow()
                    if whnd != 0:
                        ctypes.windll.user32.ShowWindow(whnd, 0)  # 0 = SW_HIDE
                        logging.info("✓ 控制台窗口已隐藏")
            except Exception as e:
                logging.warning(f"隐藏控制台窗口失败: {e}")
        
        # 在后台线程中隐藏控制台
        import threading
        threading.Thread(target=hide_console_delayed, daemon=True).start()
        
        # 加载保存的主题设置
        logging.info("加载主题配置...")
        theme_name = 'litera'  # 默认主题
        try:
            import json
            config_file = get_appdata_path('window_config.json')
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    theme_name = config.get('theme', 'litera')
                    logging.info(f"✓ 加载主题: {theme_name}")
            else:
                logging.info("配置文件不存在，使用默认主题")
        except Exception as e:
            logging.warning(f"加载主题配置失败: {e}，使用默认主题")
            theme_name = 'litera'
        
        # 使用ttkbootstrap创建主窗口
        logging.info("创建主窗口...")
        logging.info(f"主题: {theme_name}")
        try:
            root = ttk.Window(
                title="投诉批办单管理系统 v1.38.2",
                themename=theme_name,
                size=(1400, 900),
                resizable=(True, True)
            )
            logging.info("✓ 主窗口创建成功")
            
            logging.info("初始化GUI...")
            app = ComplaintManagementGUI(root)
            logging.info("✓ GUI初始化成功")
            
            logging.info("进入主循环...")
            root.mainloop()
            
        except Exception as e:
            logging.error(f"✗ 启动程序失败: {e}", exc_info=True)
            # 显示错误对话框
            try:
                import tkinter.messagebox as msgbox
                msgbox.showerror(
                    "启动失败", 
                    f"程序启动失败！\n\n错误信息：{e}\n\n请查看日志文件：{log_filename}"
                )
            except:
                pass
            raise
            
    except Exception as e:
        logging.critical(f"✗ 程序崩溃: {e}", exc_info=True)
        # 保留日志文件供分析
        logging.info(f"日志已保存到: {log_filename}")
        raise


if __name__ == '__main__':
    main()
