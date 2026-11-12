"""
投诉批办单管理系统 - 规则管理界面
提供拟办规则的可视化配置和管理功能
"""

import os
import json
import sys
import tkinter as tk
from tkinter import messagebox, scrolledtext
# 与 gui_app.py 共用同一套 ttkbootstrap 1.x / 2.x 兼容处理（PanedWindow 别名）
from ttk_compat import ttk
from ttkbootstrap.constants import *
from rule_engine import RuleEngine, get_resource_path, get_user_data_path
from platform_compat import ui_font


class RuleManagementWindow:
    """规则管理主窗口"""

    def __init__(self, parent):
        """
        初始化规则管理窗口

        Args:
            parent: 父窗口
        """
        self.parent = parent
        self.engine = RuleEngine()

        # 创建顶级窗口
        self.window = tk.Toplevel(parent)
        self.window.title("拟办规则设置")
        self.window.geometry("1200x700")

        # 当前选中的规则
        self.current_rule = None
        self.current_rule_type = None  # 'special', 'multi_dept', 'single_dept', 'title_modifier'

        # 创建界面
        self.create_widgets()

        # 加载规则列表
        self.load_rules()

        # 绑定窗口关闭事件
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        """创建界面控件"""
        # 主容器
        main_frame = ttk.Frame(self.window, padding=10)
        main_frame.pack(fill=BOTH, expand=YES)

        # 标题
        title_label = ttk.Label(
            main_frame,
            text="📋 拟办规则配置",
            font=(ui_font('微软雅黑'), 16, 'bold'),
            bootstyle="inverse-primary"
        )
        title_label.pack(fill=X, pady=(0, 10))

        # 标签页容器
        self.notebook = ttk.Notebook(main_frame, bootstyle="primary")
        self.notebook.pack(fill=BOTH, expand=YES, pady=(0, 10))

        # 创建四个标签页
        self.create_single_dept_tab()
        self.create_multi_dept_tab()
        self.create_special_tab()
        self.create_modifier_tab()

        # 底部按钮区
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=X)

        # 左侧按钮
        left_buttons = ttk.Frame(button_frame)
        left_buttons.pack(side=LEFT)

        ttk.Button(
            left_buttons,
            text="📥 导入规则",
            command=self.import_rules,
            bootstyle="info-outline",
            width=12
        ).pack(side=LEFT, padx=5)

        ttk.Button(
            left_buttons,
            text="📤 导出规则",
            command=self.export_rules,
            bootstyle="info-outline",
            width=12
        ).pack(side=LEFT, padx=5)

        ttk.Button(
            left_buttons,
            text="🔄 重置为默认",
            command=self.reset_to_default,
            bootstyle="warning-outline",
            width=15
        ).pack(side=LEFT, padx=5)

        # 右侧按钮
        right_buttons = ttk.Frame(button_frame)
        right_buttons.pack(side=RIGHT)

        ttk.Button(
            right_buttons,
            text="💾 保存",
            command=self.save_rules,
            bootstyle="success",
            width=12
        ).pack(side=LEFT, padx=5)

        ttk.Button(
            right_buttons,
            text="❌ 取消",
            command=self.on_closing,
            bootstyle="secondary",
            width=12
        ).pack(side=LEFT, padx=5)

    def create_single_dept_tab(self):
        """创建单科室规则标签页"""
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text="🏥 单科室规则")

        # 分割窗口：左侧规则列表，右侧详情
        paned = ttk.PanedWindow(tab, orient=HORIZONTAL)
        paned.pack(fill=BOTH, expand=YES)

        # 左侧：规则列表
        left_frame = ttk.Frame(paned)
        paned.add(left_frame, weight=1)

        # 列表标题
        list_title = ttk.Label(
            left_frame,
            text="规则列表",
            font=(ui_font('微软雅黑'), 12, 'bold')
        )
        list_title.pack(fill=X, pady=(0, 5))

        # 规则列表
        list_frame = ttk.Frame(left_frame)
        list_frame.pack(fill=BOTH, expand=YES)

        scrollbar = ttk.Scrollbar(list_frame, orient=VERTICAL)
        scrollbar.pack(side=RIGHT, fill=Y)

        self.single_dept_listbox = tk.Listbox(
            list_frame,
            yscrollcommand=scrollbar.set,
            font=(ui_font('微软雅黑'), 10),
            activestyle='none'
        )
        self.single_dept_listbox.pack(side=LEFT, fill=BOTH, expand=YES)
        scrollbar.config(command=self.single_dept_listbox.yview)

        self.single_dept_listbox.bind('<<ListboxSelect>>', self.on_single_dept_select)

        # 列表操作按钮
        list_buttons = ttk.Frame(left_frame)
        list_buttons.pack(fill=X, pady=(5, 0))

        ttk.Button(
            list_buttons,
            text="➕ 添加",
            command=lambda: self.add_rule('single_dept'),
            bootstyle="success-outline",
            width=10
        ).pack(side=LEFT, padx=2)

        ttk.Button(
            list_buttons,
            text="➖ 删除",
            command=lambda: self.delete_rule('single_dept'),
            bootstyle="danger-outline",
            width=10
        ).pack(side=LEFT, padx=2)

        ttk.Button(
            list_buttons,
            text="⬆ 上移",
            command=lambda: self.move_rule_up('single_dept'),
            bootstyle="info-outline",
            width=10
        ).pack(side=LEFT, padx=2)

        ttk.Button(
            list_buttons,
            text="⬇ 下移",
            command=lambda: self.move_rule_down('single_dept'),
            bootstyle="info-outline",
            width=10
        ).pack(side=LEFT, padx=2)

        # 右侧：规则详情
        right_frame = ttk.Frame(paned, padding=10)
        paned.add(right_frame, weight=2)

        # 详情标题
        detail_title = ttk.Label(
            right_frame,
            text="规则详情",
            font=(ui_font('微软雅黑'), 12, 'bold')
        )
        detail_title.pack(fill=X, pady=(0, 10))

        # 详情表单
        self.single_dept_detail_frame = ttk.Frame(right_frame)
        self.single_dept_detail_frame.pack(fill=BOTH, expand=YES)

        self.create_single_dept_form()

    def create_single_dept_form(self):
        """创建单科室规则表单"""
        frame = self.single_dept_detail_frame

        # 清空现有内容
        for widget in frame.winfo_children():
            widget.destroy()

        # 滚动区域
        canvas = tk.Canvas(frame)
        scrollbar = ttk.Scrollbar(frame, orient=VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor=NW)
        canvas.configure(yscrollcommand=scrollbar.set)

        # 规则ID（只读）
        ttk.Label(scrollable_frame, text="规则ID:", font=(ui_font('微软雅黑'), 10)).grid(
            row=0, column=0, sticky=W, padx=5, pady=5
        )
        self.single_dept_id_entry = ttk.Entry(scrollable_frame, state='readonly', width=40)
        self.single_dept_id_entry.grid(row=0, column=1, sticky=EW, padx=5, pady=5)

        # 规则名称
        ttk.Label(scrollable_frame, text="规则名称:", font=(ui_font('微软雅黑'), 10)).grid(
            row=1, column=0, sticky=W, padx=5, pady=5
        )
        self.single_dept_name_entry = ttk.Entry(scrollable_frame, width=40)
        self.single_dept_name_entry.grid(row=1, column=1, sticky=EW, padx=5, pady=5)

        # 规则描述
        ttk.Label(scrollable_frame, text="规则描述:", font=(ui_font('微软雅黑'), 10)).grid(
            row=2, column=0, sticky=W, padx=5, pady=5
        )
        self.single_dept_desc_entry = ttk.Entry(scrollable_frame, width=40)
        self.single_dept_desc_entry.grid(row=2, column=1, sticky=EW, padx=5, pady=5)

        # 优先级
        ttk.Label(scrollable_frame, text="优先级:", font=(ui_font('微软雅黑'), 10)).grid(
            row=3, column=0, sticky=W, padx=5, pady=5
        )
        self.single_dept_priority_spin = ttk.Spinbox(scrollable_frame, from_=20, to=29, width=10)
        self.single_dept_priority_spin.set(21)
        self.single_dept_priority_spin.grid(row=3, column=1, sticky=W, padx=5, pady=5)

        # 启用状态
        self.single_dept_enabled_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            scrollable_frame,
            text="启用此规则",
            variable=self.single_dept_enabled_var,
            bootstyle="success-round-toggle"
        ).grid(row=4, column=0, columnspan=2, sticky=W, padx=5, pady=5)

        # 触发条件
        ttk.Label(
            scrollable_frame,
            text="触发条件",
            font=(ui_font('微软雅黑'), 11, 'bold')
        ).grid(row=5, column=0, columnspan=2, sticky=W, padx=5, pady=(15, 5))

        # 文件名包含
        ttk.Label(scrollable_frame, text="文件名包含:", font=(ui_font('微软雅黑'), 10)).grid(
            row=6, column=0, sticky=W, padx=5, pady=5
        )
        self.single_dept_filename_entry = ttk.Entry(scrollable_frame, width=40)
        self.single_dept_filename_entry.grid(row=6, column=1, sticky=EW, padx=5, pady=5)
        ttk.Label(
            scrollable_frame,
            text="多个关键词用逗号分隔",
            font=(ui_font('微软雅黑'), 9),
            foreground='gray'
        ).grid(row=7, column=1, sticky=W, padx=5)

        # 文件名不包含
        ttk.Label(scrollable_frame, text="文件名不包含:", font=(ui_font('微软雅黑'), 10)).grid(
            row=8, column=0, sticky=W, padx=5, pady=5
        )
        self.single_dept_not_filename_entry = ttk.Entry(scrollable_frame, width=40)
        self.single_dept_not_filename_entry.grid(row=8, column=1, sticky=EW, padx=5, pady=5)

        # 填充内容
        ttk.Label(
            scrollable_frame,
            text="填充内容",
            font=(ui_font('微软雅黑'), 11, 'bold')
        ).grid(row=9, column=0, columnspan=2, sticky=W, padx=5, pady=(15, 5))

        # 拟办意见
        ttk.Label(scrollable_frame, text="拟办意见:", font=(ui_font('微软雅黑'), 10)).grid(
            row=10, column=0, sticky=NW, padx=5, pady=5
        )
        self.single_dept_opinion_text = scrolledtext.ScrolledText(
            scrollable_frame,
            width=40,
            height=6,
            font=(ui_font('微软雅黑'), 10),
            wrap=tk.WORD
        )
        self.single_dept_opinion_text.grid(row=10, column=1, sticky=EW, padx=5, pady=5)

        # 办理结果
        ttk.Label(scrollable_frame, text="办理结果:", font=(ui_font('微软雅黑'), 10)).grid(
            row=11, column=0, sticky=NW, padx=5, pady=5
        )
        self.single_dept_result_text = scrolledtext.ScrolledText(
            scrollable_frame,
            width=40,
            height=3,
            font=(ui_font('微软雅黑'), 10),
            wrap=tk.WORD
        )
        self.single_dept_result_text.grid(row=11, column=1, sticky=EW, padx=5, pady=5)

        # 按钮
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.grid(row=12, column=0, columnspan=2, pady=(15, 5))

        ttk.Button(
            button_frame,
            text="🧪 测试规则",
            command=self.test_current_rule,
            bootstyle="info",
            width=15
        ).pack(side=LEFT, padx=5)

        ttk.Button(
            button_frame,
            text="✅ 应用",
            command=lambda: self.apply_rule('single_dept'),
            bootstyle="success",
            width=15
        ).pack(side=LEFT, padx=5)

        scrollable_frame.columnconfigure(1, weight=1)

        canvas.pack(side=LEFT, fill=BOTH, expand=YES)
        scrollbar.pack(side=RIGHT, fill=Y)

    def create_multi_dept_tab(self):
        """创建多科室联合规则标签页"""
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text="🤝 多科室联合规则")

        # 分割窗口：左侧规则列表，右侧详情
        paned = ttk.PanedWindow(tab, orient=HORIZONTAL)
        paned.pack(fill=BOTH, expand=YES)

        # 左侧：规则列表
        left_frame = ttk.Frame(paned)
        paned.add(left_frame, weight=1)

        list_title = ttk.Label(
            left_frame,
            text="规则列表",
            font=(ui_font('微软雅黑'), 12, 'bold')
        )
        list_title.pack(fill=X, pady=(0, 5))

        # 规则列表
        list_frame = ttk.Frame(left_frame)
        list_frame.pack(fill=BOTH, expand=YES)

        scrollbar = ttk.Scrollbar(list_frame, orient=VERTICAL)
        scrollbar.pack(side=RIGHT, fill=Y)

        self.multi_dept_listbox = tk.Listbox(
            list_frame,
            yscrollcommand=scrollbar.set,
            font=(ui_font('微软雅黑'), 10),
            activestyle='none'
        )
        self.multi_dept_listbox.pack(side=LEFT, fill=BOTH, expand=YES)
        scrollbar.config(command=self.multi_dept_listbox.yview)

        self.multi_dept_listbox.bind('<<ListboxSelect>>', self.on_multi_dept_select)

        # 列表操作按钮
        list_buttons = ttk.Frame(left_frame)
        list_buttons.pack(fill=X, pady=(5, 0))

        ttk.Button(
            list_buttons,
            text="➕ 添加",
            command=lambda: self.add_rule('multi_dept'),
            bootstyle="success-outline",
            width=10
        ).pack(side=LEFT, padx=2)

        ttk.Button(
            list_buttons,
            text="➖ 删除",
            command=lambda: self.delete_rule('multi_dept'),
            bootstyle="danger-outline",
            width=10
        ).pack(side=LEFT, padx=2)

        # 右侧：规则详情
        right_frame = ttk.Frame(paned, padding=10)
        paned.add(right_frame, weight=2)

        detail_title = ttk.Label(
            right_frame,
            text="规则详情",
            font=(ui_font('微软雅黑'), 12, 'bold')
        )
        detail_title.pack(fill=X, pady=(0, 10))

        # 详情表单
        self.multi_dept_detail_frame = ttk.Frame(right_frame)
        self.multi_dept_detail_frame.pack(fill=BOTH, expand=YES)

        self.create_multi_dept_form()

    def create_multi_dept_form(self):
        """创建多科室联合规则表单"""
        frame = self.multi_dept_detail_frame

        # 清空现有内容
        for widget in frame.winfo_children():
            widget.destroy()

        # 滚动区域
        canvas = tk.Canvas(frame)
        scrollbar = ttk.Scrollbar(frame, orient=VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor=NW)
        canvas.configure(yscrollcommand=scrollbar.set)

        # 规则ID（只读）
        ttk.Label(scrollable_frame, text="规则ID:", font=(ui_font('微软雅黑'), 10)).grid(
            row=0, column=0, sticky=W, padx=5, pady=5
        )
        self.multi_dept_id_entry = ttk.Entry(scrollable_frame, state='readonly', width=40)
        self.multi_dept_id_entry.grid(row=0, column=1, sticky=EW, padx=5, pady=5)

        # 规则名称
        ttk.Label(scrollable_frame, text="规则名称:", font=(ui_font('微软雅黑'), 10)).grid(
            row=1, column=0, sticky=W, padx=5, pady=5
        )
        self.multi_dept_name_entry = ttk.Entry(scrollable_frame, width=40)
        self.multi_dept_name_entry.grid(row=1, column=1, sticky=EW, padx=5, pady=5)

        # 规则描述
        ttk.Label(scrollable_frame, text="规则描述:", font=(ui_font('微软雅黑'), 10)).grid(
            row=2, column=0, sticky=W, padx=5, pady=5
        )
        self.multi_dept_desc_entry = ttk.Entry(scrollable_frame, width=40)
        self.multi_dept_desc_entry.grid(row=2, column=1, sticky=EW, padx=5, pady=5)

        # 优先级
        ttk.Label(scrollable_frame, text="优先级:", font=(ui_font('微软雅黑'), 10)).grid(
            row=3, column=0, sticky=W, padx=5, pady=5
        )
        self.multi_dept_priority_spin = ttk.Spinbox(scrollable_frame, from_=10, to=19, width=10)
        self.multi_dept_priority_spin.set(10)
        self.multi_dept_priority_spin.grid(row=3, column=1, sticky=W, padx=5, pady=5)

        # 启用状态
        self.multi_dept_enabled_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            scrollable_frame,
            text="启用此规则",
            variable=self.multi_dept_enabled_var,
            bootstyle="success-round-toggle"
        ).grid(row=4, column=0, columnspan=2, sticky=W, padx=5, pady=5)

        # 触发条件
        ttk.Label(
            scrollable_frame,
            text="触发条件",
            font=(ui_font('微软雅黑'), 11, 'bold')
        ).grid(row=5, column=0, columnspan=2, sticky=W, padx=5, pady=(15, 5))

        # 科室关键词顺序
        ttk.Label(scrollable_frame, text="科室关键词顺序:", font=(ui_font('微软雅黑'), 10)).grid(
            row=6, column=0, sticky=NW, padx=5, pady=5
        )
        order_frame = ttk.Frame(scrollable_frame)
        order_frame.grid(row=6, column=1, sticky=EW, padx=5, pady=5)

        ttk.Label(order_frame, text="第一位（牵头）:").pack(anchor=W)
        self.multi_dept_first_entry = ttk.Entry(order_frame, width=20)
        self.multi_dept_first_entry.pack(fill=X, pady=2)

        ttk.Label(order_frame, text="第二位（配合）:").pack(anchor=W, pady=(10, 0))
        self.multi_dept_second_entry = ttk.Entry(order_frame, width=20)
        self.multi_dept_second_entry.pack(fill=X, pady=2)

        ttk.Label(
            scrollable_frame,
            text="多个关键词用逗号分隔",
            font=(ui_font('微软雅黑'), 9),
            foreground='gray'
        ).grid(row=7, column=1, sticky=W, padx=5)

        # 填充内容
        ttk.Label(
            scrollable_frame,
            text="填充内容",
            font=(ui_font('微软雅黑'), 11, 'bold')
        ).grid(row=8, column=0, columnspan=2, sticky=W, padx=5, pady=(15, 5))

        # 拟办意见
        ttk.Label(scrollable_frame, text="拟办意见:", font=(ui_font('微软雅黑'), 10)).grid(
            row=9, column=0, sticky=NW, padx=5, pady=5
        )
        self.multi_dept_opinion_text = scrolledtext.ScrolledText(
            scrollable_frame,
            width=40,
            height=6,
            font=(ui_font('微软雅黑'), 10),
            wrap=tk.WORD
        )
        self.multi_dept_opinion_text.grid(row=9, column=1, sticky=EW, padx=5, pady=5)

        # 办理结果
        ttk.Label(scrollable_frame, text="办理结果:", font=(ui_font('微软雅黑'), 10)).grid(
            row=10, column=0, sticky=NW, padx=5, pady=5
        )
        self.multi_dept_result_text = scrolledtext.ScrolledText(
            scrollable_frame,
            width=40,
            height=3,
            font=(ui_font('微软雅黑'), 10),
            wrap=tk.WORD
        )
        self.multi_dept_result_text.grid(row=10, column=1, sticky=EW, padx=5, pady=5)

        # 按钮
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.grid(row=11, column=0, columnspan=2, pady=(15, 5))

        ttk.Button(
            button_frame,
            text="🧪 测试规则",
            command=self.test_current_rule,
            bootstyle="info",
            width=15
        ).pack(side=LEFT, padx=5)

        ttk.Button(
            button_frame,
            text="✅ 应用",
            command=lambda: self.apply_rule('multi_dept'),
            bootstyle="success",
            width=15
        ).pack(side=LEFT, padx=5)

        scrollable_frame.columnconfigure(1, weight=1)

        canvas.pack(side=LEFT, fill=BOTH, expand=YES)
        scrollbar.pack(side=RIGHT, fill=Y)

    def create_special_tab(self):
        """创建特殊情况规则标签页"""
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text="⚡ 特殊情况规则")

        # 分割窗口：左侧规则列表，右侧详情
        paned = ttk.PanedWindow(tab, orient=HORIZONTAL)
        paned.pack(fill=BOTH, expand=YES)

        # 左侧：规则列表
        left_frame = ttk.Frame(paned)
        paned.add(left_frame, weight=1)

        list_title = ttk.Label(
            left_frame,
            text="规则列表",
            font=(ui_font('微软雅黑'), 12, 'bold')
        )
        list_title.pack(fill=X, pady=(0, 5))

        # 规则列表
        list_frame = ttk.Frame(left_frame)
        list_frame.pack(fill=BOTH, expand=YES)

        scrollbar = ttk.Scrollbar(list_frame, orient=VERTICAL)
        scrollbar.pack(side=RIGHT, fill=Y)

        self.special_listbox = tk.Listbox(
            list_frame,
            yscrollcommand=scrollbar.set,
            font=(ui_font('微软雅黑'), 10),
            activestyle='none'
        )
        self.special_listbox.pack(side=LEFT, fill=BOTH, expand=YES)
        scrollbar.config(command=self.special_listbox.yview)

        self.special_listbox.bind('<<ListboxSelect>>', self.on_special_select)

        # 列表操作按钮
        list_buttons = ttk.Frame(left_frame)
        list_buttons.pack(fill=X, pady=(5, 0))

        ttk.Button(
            list_buttons,
            text="➕ 添加",
            command=lambda: self.add_rule('special'),
            bootstyle="success-outline",
            width=10
        ).pack(side=LEFT, padx=2)

        ttk.Button(
            list_buttons,
            text="➖ 删除",
            command=lambda: self.delete_rule('special'),
            bootstyle="danger-outline",
            width=10
        ).pack(side=LEFT, padx=2)

        # 右侧：规则详情
        right_frame = ttk.Frame(paned, padding=10)
        paned.add(right_frame, weight=2)

        detail_title = ttk.Label(
            right_frame,
            text="规则详情",
            font=(ui_font('微软雅黑'), 12, 'bold')
        )
        detail_title.pack(fill=X, pady=(0, 10))

        # 详情表单
        self.special_detail_frame = ttk.Frame(right_frame)
        self.special_detail_frame.pack(fill=BOTH, expand=YES)

        self.create_special_form()

    def create_special_form(self):
        """创建特殊情况规则表单"""
        frame = self.special_detail_frame

        # 清空现有内容
        for widget in frame.winfo_children():
            widget.destroy()

        # 滚动区域
        canvas = tk.Canvas(frame)
        scrollbar = ttk.Scrollbar(frame, orient=VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor=NW)
        canvas.configure(yscrollcommand=scrollbar.set)

        # 规则ID（只读）
        ttk.Label(scrollable_frame, text="规则ID:", font=(ui_font('微软雅黑'), 10)).grid(
            row=0, column=0, sticky=W, padx=5, pady=5
        )
        self.special_id_entry = ttk.Entry(scrollable_frame, state='readonly', width=40)
        self.special_id_entry.grid(row=0, column=1, sticky=EW, padx=5, pady=5)

        # 规则名称
        ttk.Label(scrollable_frame, text="规则名称:", font=(ui_font('微软雅黑'), 10)).grid(
            row=1, column=0, sticky=W, padx=5, pady=5
        )
        self.special_name_entry = ttk.Entry(scrollable_frame, width=40)
        self.special_name_entry.grid(row=1, column=1, sticky=EW, padx=5, pady=5)

        # 规则描述
        ttk.Label(scrollable_frame, text="规则描述:", font=(ui_font('微软雅黑'), 10)).grid(
            row=2, column=0, sticky=W, padx=5, pady=5
        )
        self.special_desc_entry = ttk.Entry(scrollable_frame, width=40)
        self.special_desc_entry.grid(row=2, column=1, sticky=EW, padx=5, pady=5)

        # 优先级
        ttk.Label(scrollable_frame, text="优先级:", font=(ui_font('微软雅黑'), 10)).grid(
            row=3, column=0, sticky=W, padx=5, pady=5
        )
        self.special_priority_spin = ttk.Spinbox(scrollable_frame, from_=1, to=9, width=10)
        self.special_priority_spin.set(5)
        self.special_priority_spin.grid(row=3, column=1, sticky=W, padx=5, pady=5)

        # 启用状态
        self.special_enabled_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            scrollable_frame,
            text="启用此规则",
            variable=self.special_enabled_var,
            bootstyle="success-round-toggle"
        ).grid(row=4, column=0, columnspan=2, sticky=W, padx=5, pady=5)

        # 规则类型
        ttk.Label(scrollable_frame, text="规则类型:", font=(ui_font('微软雅黑'), 10)).grid(
            row=5, column=0, sticky=W, padx=5, pady=5
        )
        self.special_type_var = tk.StringVar(value="replace")
        type_frame = ttk.Frame(scrollable_frame)
        type_frame.grid(row=5, column=1, sticky=W, padx=5, pady=5)

        ttk.Radiobutton(
            type_frame,
            text="完全替换",
            variable=self.special_type_var,
            value="replace",
            bootstyle="info",
            command=self.on_special_type_change
        ).pack(side=LEFT, padx=(0, 10))

        ttk.Radiobutton(
            type_frame,
            text="修正规则",
            variable=self.special_type_var,
            value="modifier",
            bootstyle="info",
            command=self.on_special_type_change
        ).pack(side=LEFT)

        ttk.Label(
            scrollable_frame,
            text="完全替换：直接替换填充内容；修正规则：在其他规则基础上修改",
            font=(ui_font('微软雅黑'), 9),
            foreground='gray'
        ).grid(row=6, column=1, sticky=W, padx=5)

        # 触发条件
        ttk.Label(
            scrollable_frame,
            text="触发条件",
            font=(ui_font('微软雅黑'), 11, 'bold')
        ).grid(row=7, column=0, columnspan=2, sticky=W, padx=5, pady=(15, 5))

        # 标题关键词
        ttk.Label(scrollable_frame, text="标题包含:", font=(ui_font('微软雅黑'), 10)).grid(
            row=8, column=0, sticky=W, padx=5, pady=5
        )
        self.special_title_kw_entry = ttk.Entry(scrollable_frame, width=40)
        self.special_title_kw_entry.grid(row=8, column=1, sticky=EW, padx=5, pady=5)

        # 文件名关键词
        ttk.Label(scrollable_frame, text="文件名包含:", font=(ui_font('微软雅黑'), 10)).grid(
            row=9, column=0, sticky=W, padx=5, pady=5
        )
        self.special_filename_kw_entry = ttk.Entry(scrollable_frame, width=40)
        self.special_filename_kw_entry.grid(row=9, column=1, sticky=EW, padx=5, pady=5)

        ttk.Label(
            scrollable_frame,
            text="多个关键词用逗号分隔；至少填写一项",
            font=(ui_font('微软雅黑'), 9),
            foreground='gray'
        ).grid(row=10, column=1, sticky=W, padx=5)

        # 公共场所识别码管理区域（仅对公共场所规则显示）
        self.public_place_frame = ttk.LabelFrame(scrollable_frame, text="公共场所识别码", padding=10)
        self.public_place_frame.grid(row=11, column=0, columnspan=2, sticky=EW, padx=5, pady=10)
        self.public_place_frame.grid_remove()  # 默认隐藏

        # 识别码列表
        list_container = ttk.Frame(self.public_place_frame)
        list_container.pack(fill=BOTH, expand=YES)

        self.public_place_listbox = tk.Listbox(list_container, height=8, font=(ui_font('微软雅黑'), 10))
        self.public_place_listbox.pack(side=LEFT, fill=BOTH, expand=YES)

        list_scrollbar = ttk.Scrollbar(list_container, orient=VERTICAL, command=self.public_place_listbox.yview)
        list_scrollbar.pack(side=RIGHT, fill=Y)
        self.public_place_listbox.config(yscrollcommand=list_scrollbar.set)

        # 识别码操作按钮
        btn_frame = ttk.Frame(self.public_place_frame)
        btn_frame.pack(fill=X, pady=(10, 0))

        self.public_place_add_entry = ttk.Entry(btn_frame, width=20)
        self.public_place_add_entry.pack(side=LEFT, padx=(0, 5))

        ttk.Button(btn_frame, text="➕ 添加", command=self.add_public_place_keyword,
                  bootstyle="success-outline", width=8).pack(side=LEFT, padx=2)
        ttk.Button(btn_frame, text="➖ 删除", command=self.remove_public_place_keyword,
                  bootstyle="danger-outline", width=8).pack(side=LEFT, padx=2)

        # 填充内容区域（根据规则类型显示或隐藏）
        self.output_content_frame = ttk.LabelFrame(scrollable_frame, text="填充内容", padding=10)
        self.output_content_frame.grid(row=12, column=0, columnspan=2, sticky=EW, padx=5, pady=10)

        # 拟办意见
        ttk.Label(self.output_content_frame, text="拟办意见:", font=(ui_font('微软雅黑'), 10)).grid(
            row=0, column=0, sticky=NW, padx=5, pady=5
        )
        self.special_opinion_text = scrolledtext.ScrolledText(
            self.output_content_frame,
            width=40,
            height=6,
            font=(ui_font('微软雅黑'), 10),
            wrap=tk.WORD
        )
        self.special_opinion_text.grid(row=0, column=1, sticky=EW, padx=5, pady=5)

        # 办理结果
        ttk.Label(self.output_content_frame, text="办理结果:", font=(ui_font('微软雅黑'), 10)).grid(
            row=1, column=0, sticky=NW, padx=5, pady=5
        )
        self.special_result_text = scrolledtext.ScrolledText(
            self.output_content_frame,
            width=40,
            height=3,
            font=(ui_font('微软雅黑'), 10),
            wrap=tk.WORD
        )
        self.special_result_text.grid(row=1, column=1, sticky=EW, padx=5, pady=5)

        self.output_content_frame.columnconfigure(1, weight=1)

        # 修正规则区域（modifier类型）
        self.modifier_content_frame = ttk.LabelFrame(scrollable_frame, text="拟办意见修正规则", padding=10)
        self.modifier_content_frame.grid(row=13, column=0, columnspan=2, sticky=EW, padx=5, pady=10)
        self.modifier_content_frame.grid_remove()  # 默认隐藏

        self.modifier_rules_text = scrolledtext.ScrolledText(
            self.modifier_content_frame,
            width=60,
            height=8,
            font=(ui_font('微软雅黑'), 9),
            wrap=tk.WORD
        )
        self.modifier_rules_text.pack(fill=BOTH, expand=YES)

        ttk.Label(
            self.modifier_content_frame,
            text="格式示例：删除\"，请法监科亦阅\" | 替换\"妥处\"为\"办理\"",
            font=(ui_font('微软雅黑'), 9),
            foreground='blue'
        ).pack(anchor=W, pady=(5, 0))

        # 按钮
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.grid(row=14, column=0, columnspan=2, pady=(15, 5))

        ttk.Button(
            button_frame,
            text="🧪 测试规则",
            command=self.test_current_rule,
            bootstyle="info",
            width=15
        ).pack(side=LEFT, padx=5)

        ttk.Button(
            button_frame,
            text="✅ 应用",
            command=lambda: self.apply_rule('special'),
            bootstyle="success",
            width=15
        ).pack(side=LEFT, padx=5)

        scrollable_frame.columnconfigure(1, weight=1)

        canvas.pack(side=LEFT, fill=BOTH, expand=YES)
        scrollbar.pack(side=RIGHT, fill=Y)

    def create_modifier_tab(self):
        """创建标题修正规则标签页"""
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text="📝 标题修正规则")

        # 分割窗口：左侧规则列表，右侧详情
        paned = ttk.PanedWindow(tab, orient=HORIZONTAL)
        paned.pack(fill=BOTH, expand=YES)

        # 左侧：规则列表
        left_frame = ttk.Frame(paned)
        paned.add(left_frame, weight=1)

        list_title = ttk.Label(
            left_frame,
            text="规则列表",
            font=(ui_font('微软雅黑'), 12, 'bold')
        )
        list_title.pack(fill=X, pady=(0, 5))

        # 规则列表
        list_frame = ttk.Frame(left_frame)
        list_frame.pack(fill=BOTH, expand=YES)

        scrollbar = ttk.Scrollbar(list_frame, orient=VERTICAL)
        scrollbar.pack(side=RIGHT, fill=Y)

        self.modifier_listbox = tk.Listbox(
            list_frame,
            yscrollcommand=scrollbar.set,
            font=(ui_font('微软雅黑'), 10),
            activestyle='none'
        )
        self.modifier_listbox.pack(side=LEFT, fill=BOTH, expand=YES)
        scrollbar.config(command=self.modifier_listbox.yview)

        self.modifier_listbox.bind('<<ListboxSelect>>', self.on_modifier_select)

        # 列表操作按钮
        list_buttons = ttk.Frame(left_frame)
        list_buttons.pack(fill=X, pady=(5, 0))

        ttk.Button(
            list_buttons,
            text="➕ 添加",
            command=lambda: self.add_rule('modifier'),
            bootstyle="success-outline",
            width=10
        ).pack(side=LEFT, padx=2)

        ttk.Button(
            list_buttons,
            text="➖ 删除",
            command=lambda: self.delete_rule('modifier'),
            bootstyle="danger-outline",
            width=10
        ).pack(side=LEFT, padx=2)

        # 右侧：规则详情
        right_frame = ttk.Frame(paned, padding=10)
        paned.add(right_frame, weight=2)

        detail_title = ttk.Label(
            right_frame,
            text="规则详情",
            font=(ui_font('微软雅黑'), 12, 'bold')
        )
        detail_title.pack(fill=X, pady=(0, 10))

        # 详情表单
        self.modifier_detail_frame = ttk.Frame(right_frame)
        self.modifier_detail_frame.pack(fill=BOTH, expand=YES)

        self.create_modifier_form()

    def create_modifier_form(self):
        """创建标题修正规则表单"""
        frame = self.modifier_detail_frame

        # 清空现有内容
        for widget in frame.winfo_children():
            widget.destroy()

        # 滚动区域
        canvas = tk.Canvas(frame)
        scrollbar = ttk.Scrollbar(frame, orient=VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor=NW)
        canvas.configure(yscrollcommand=scrollbar.set)

        # 规则ID（只读）
        ttk.Label(scrollable_frame, text="规则ID:", font=(ui_font('微软雅黑'), 10)).grid(
            row=0, column=0, sticky=W, padx=5, pady=5
        )
        self.modifier_id_entry = ttk.Entry(scrollable_frame, state='readonly', width=40)
        self.modifier_id_entry.grid(row=0, column=1, sticky=EW, padx=5, pady=5)

        # 规则名称
        ttk.Label(scrollable_frame, text="规则名称:", font=(ui_font('微软雅黑'), 10)).grid(
            row=1, column=0, sticky=W, padx=5, pady=5
        )
        self.modifier_name_entry = ttk.Entry(scrollable_frame, width=40)
        self.modifier_name_entry.grid(row=1, column=1, sticky=EW, padx=5, pady=5)

        # 规则描述
        ttk.Label(scrollable_frame, text="规则描述:", font=(ui_font('微软雅黑'), 10)).grid(
            row=2, column=0, sticky=W, padx=5, pady=5
        )
        self.modifier_desc_entry = ttk.Entry(scrollable_frame, width=40)
        self.modifier_desc_entry.grid(row=2, column=1, sticky=EW, padx=5, pady=5)

        # 启用状态
        self.modifier_enabled_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            scrollable_frame,
            text="启用此规则",
            variable=self.modifier_enabled_var,
            bootstyle="success-round-toggle"
        ).grid(row=3, column=0, columnspan=2, sticky=W, padx=5, pady=5)

        # 触发条件
        ttk.Label(
            scrollable_frame,
            text="触发条件",
            font=(ui_font('微软雅黑'), 11, 'bold')
        ).grid(row=4, column=0, columnspan=2, sticky=W, padx=5, pady=(15, 5))

        # 文件名关键词
        ttk.Label(scrollable_frame, text="文件名包含:", font=(ui_font('微软雅黑'), 10)).grid(
            row=5, column=0, sticky=W, padx=5, pady=5
        )
        self.modifier_filename_kw_entry = ttk.Entry(scrollable_frame, width=40)
        self.modifier_filename_kw_entry.grid(row=5, column=1, sticky=EW, padx=5, pady=5)

        ttk.Label(
            scrollable_frame,
            text="多个关键词用逗号分隔",
            font=(ui_font('微软雅黑'), 9),
            foreground='gray'
        ).grid(row=6, column=1, sticky=W, padx=5)

        # 修正操作
        ttk.Label(
            scrollable_frame,
            text="修正操作",
            font=(ui_font('微软雅黑'), 11, 'bold')
        ).grid(row=7, column=0, columnspan=2, sticky=W, padx=5, pady=(15, 5))

        # 修正类型
        ttk.Label(scrollable_frame, text="修正类型:", font=(ui_font('微软雅黑'), 10)).grid(
            row=8, column=0, sticky=W, padx=5, pady=5
        )
        self.modifier_action_var = tk.StringVar(value="append")
        action_frame = ttk.Frame(scrollable_frame)
        action_frame.grid(row=8, column=1, sticky=W, padx=5, pady=5)

        ttk.Radiobutton(
            action_frame,
            text="追加",
            variable=self.modifier_action_var,
            value="append",
            bootstyle="info"
        ).pack(side=LEFT, padx=(0, 10))

        ttk.Radiobutton(
            action_frame,
            text="前置",
            variable=self.modifier_action_var,
            value="prepend",
            bootstyle="info"
        ).pack(side=LEFT, padx=(0, 10))

        ttk.Radiobutton(
            action_frame,
            text="替换",
            variable=self.modifier_action_var,
            value="replace",
            bootstyle="info"
        ).pack(side=LEFT)

        # 修正文本
        ttk.Label(scrollable_frame, text="修正文本:", font=(ui_font('微软雅黑'), 10)).grid(
            row=9, column=0, sticky=NW, padx=5, pady=5
        )
        self.modifier_text_entry = scrolledtext.ScrolledText(
            scrollable_frame,
            width=40,
            height=4,
            font=(ui_font('微软雅黑'), 10),
            wrap=tk.WORD
        )
        self.modifier_text_entry.grid(row=9, column=1, sticky=EW, padx=5, pady=5)

        # 替换规则
        replace_frame = ttk.LabelFrame(scrollable_frame, text="替换规则（仅替换模式）", padding=10)
        replace_frame.grid(row=10, column=0, columnspan=2, sticky=EW, padx=5, pady=10)

        ttk.Label(replace_frame, text="查找文本:", font=(ui_font('微软雅黑'), 10)).grid(
            row=0, column=0, sticky=W, padx=5, pady=5
        )
        self.modifier_find_entry = ttk.Entry(replace_frame, width=35)
        self.modifier_find_entry.grid(row=0, column=1, sticky=EW, padx=5, pady=5)

        ttk.Label(replace_frame, text="替换为:", font=(ui_font('微软雅黑'), 10)).grid(
            row=1, column=0, sticky=W, padx=5, pady=5
        )
        self.modifier_replace_entry = ttk.Entry(replace_frame, width=35)
        self.modifier_replace_entry.grid(row=1, column=1, sticky=EW, padx=5, pady=5)

        replace_frame.columnconfigure(1, weight=1)

        # 按钮
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.grid(row=11, column=0, columnspan=2, pady=(15, 5))

        ttk.Button(
            button_frame,
            text="🧪 测试规则",
            command=self.test_current_rule,
            bootstyle="info",
            width=15
        ).pack(side=LEFT, padx=5)

        ttk.Button(
            button_frame,
            text="✅ 应用",
            command=lambda: self.apply_rule('modifier'),
            bootstyle="success",
            width=15
        ).pack(side=LEFT, padx=5)

        scrollable_frame.columnconfigure(1, weight=1)

        canvas.pack(side=LEFT, fill=BOTH, expand=YES)
        scrollbar.pack(side=RIGHT, fill=Y)

    def load_rules(self):
        """加载规则到列表"""
        self.load_single_dept_rules()
        self.load_multi_dept_rules()
        self.load_special_rules()
        self.load_modifier_rules()

    def load_single_dept_rules(self):
        """加载单科室规则列表"""
        self.single_dept_listbox.delete(0, tk.END)

        rules = self.engine.rules.get('single_dept_rules', [])
        # 按优先级排序
        rules = sorted(rules, key=lambda x: x.get('priority', 999))

        for rule in rules:
            enabled = "☑" if rule.get('enabled', True) else "☐"
            name = rule.get('name', rule.get('id', '未命名'))
            display_text = f"{enabled} {name}"
            self.single_dept_listbox.insert(tk.END, display_text)

    def load_multi_dept_rules(self):
        """加载多科室联合规则列表"""
        self.multi_dept_listbox.delete(0, tk.END)

        rules = self.engine.rules.get('multi_dept_rules', [])
        # 按优先级排序
        rules = sorted(rules, key=lambda x: x.get('priority', 999))

        for rule in rules:
            enabled = "☑" if rule.get('enabled', True) else "☐"
            name = rule.get('name', rule.get('id', '未命名'))
            display_text = f"{enabled} {name}"
            self.multi_dept_listbox.insert(tk.END, display_text)

    def load_special_rules(self):
        """加载特殊情况规则列表"""
        self.special_listbox.delete(0, tk.END)

        rules = self.engine.rules.get('special_rules', [])
        # 按优先级排序
        rules = sorted(rules, key=lambda x: x.get('priority', 999))

        for rule in rules:
            enabled = "☑" if rule.get('enabled', True) else "☐"
            name = rule.get('name', rule.get('id', '未命名'))
            display_text = f"{enabled} {name}"
            self.special_listbox.insert(tk.END, display_text)

    def load_modifier_rules(self):
        """加载标题修正规则列表"""
        self.modifier_listbox.delete(0, tk.END)

        rules = self.engine.rules.get('title_modifiers', [])
        # 按优先级排序
        rules = sorted(rules, key=lambda x: x.get('priority', 999))

        for rule in rules:
            enabled = "☑" if rule.get('enabled', True) else "☐"
            name = rule.get('name', rule.get('id', '未命名'))
            display_text = f"{enabled} {name}"
            self.modifier_listbox.insert(tk.END, display_text)

    def on_single_dept_select(self, event):
        """单科室规则选择事件"""
        selection = self.single_dept_listbox.curselection()
        if not selection:
            return

        index = selection[0]
        rules = sorted(
            self.engine.rules.get('single_dept_rules', []),
            key=lambda x: x.get('priority', 999)
        )

        if index < len(rules):
            self.current_rule = rules[index]
            self.current_rule_type = 'single_dept'
            self.display_single_dept_rule(self.current_rule)

    def on_multi_dept_select(self, event):
        """多科室联合规则选择事件"""
        selection = self.multi_dept_listbox.curselection()
        if not selection:
            return

        index = selection[0]
        rules = sorted(
            self.engine.rules.get('multi_dept_rules', []),
            key=lambda x: x.get('priority', 999)
        )

        if index < len(rules):
            self.current_rule = rules[index]
            self.current_rule_type = 'multi_dept'
            self.display_multi_dept_rule(self.current_rule)

    def on_special_select(self, event):
        """特殊情况规则选择事件"""
        selection = self.special_listbox.curselection()
        if not selection:
            return

        index = selection[0]
        rules = sorted(
            self.engine.rules.get('special_rules', []),
            key=lambda x: x.get('priority', 999)
        )

        if index < len(rules):
            self.current_rule = rules[index]
            self.current_rule_type = 'special'
            self.display_special_rule(self.current_rule)

    def on_modifier_select(self, event):
        """标题修正规则选择事件"""
        selection = self.modifier_listbox.curselection()
        if not selection:
            return

        index = selection[0]
        rules = sorted(
            self.engine.rules.get('title_modifiers', []),
            key=lambda x: x.get('priority', 999)
        )

        if index < len(rules):
            self.current_rule = rules[index]
            self.current_rule_type = 'modifier'
            self.display_modifier_rule(self.current_rule)

    def display_single_dept_rule(self, rule):
        """显示单科室规则详情"""
        # 清空表单
        self.single_dept_id_entry.configure(state='normal')
        self.single_dept_id_entry.delete(0, tk.END)
        self.single_dept_name_entry.delete(0, tk.END)
        self.single_dept_desc_entry.delete(0, tk.END)
        self.single_dept_filename_entry.delete(0, tk.END)
        self.single_dept_not_filename_entry.delete(0, tk.END)
        self.single_dept_opinion_text.delete(1.0, tk.END)
        self.single_dept_result_text.delete(1.0, tk.END)

        # 填充数据
        self.single_dept_id_entry.insert(0, rule.get('id', ''))
        self.single_dept_id_entry.configure(state='readonly')
        self.single_dept_name_entry.insert(0, rule.get('name', ''))
        self.single_dept_desc_entry.insert(0, rule.get('description', ''))
        self.single_dept_priority_spin.set(rule.get('priority', 21))
        self.single_dept_enabled_var.set(rule.get('enabled', True))

        # 条件
        conditions = rule.get('conditions', {})
        filename_contains = conditions.get('filename_contains', [])
        if filename_contains:
            self.single_dept_filename_entry.insert(0, ', '.join(filename_contains))

        filename_not_contains = conditions.get('filename_not_contains', [])
        if filename_not_contains:
            self.single_dept_not_filename_entry.insert(0, ', '.join(filename_not_contains))

        # 输出
        output = rule.get('output', {})
        self.single_dept_opinion_text.insert(1.0, output.get('拟办意见', ''))
        self.single_dept_result_text.insert(1.0, output.get('办理结果', ''))

    def display_multi_dept_rule(self, rule):
        """显示多科室联合规则详情"""
        # 清空表单
        self.multi_dept_id_entry.configure(state='normal')
        self.multi_dept_id_entry.delete(0, tk.END)
        self.multi_dept_name_entry.delete(0, tk.END)
        self.multi_dept_desc_entry.delete(0, tk.END)
        self.multi_dept_first_entry.delete(0, tk.END)
        self.multi_dept_second_entry.delete(0, tk.END)
        self.multi_dept_opinion_text.delete(1.0, tk.END)
        if hasattr(self, 'multi_dept_result_text'):
            self.multi_dept_result_text.delete(1.0, tk.END)

        # 填充数据
        self.multi_dept_id_entry.insert(0, rule.get('id', ''))
        self.multi_dept_id_entry.configure(state='readonly')
        self.multi_dept_name_entry.insert(0, rule.get('name', ''))
        self.multi_dept_desc_entry.insert(0, rule.get('description', ''))
        self.multi_dept_priority_spin.set(rule.get('priority', 11))
        self.multi_dept_enabled_var.set(rule.get('enabled', True))

        # 条件 - 科室关键词顺序
        conditions = rule.get('conditions', {})
        dept_keywords_order = conditions.get('dept_keywords_order', [])
        if len(dept_keywords_order) >= 1:
            self.multi_dept_first_entry.insert(0, ', '.join(dept_keywords_order[0]))
        if len(dept_keywords_order) >= 2:
            self.multi_dept_second_entry.insert(0, ', '.join(dept_keywords_order[1]))

        # 输出
        output = rule.get('output', {})
        self.multi_dept_opinion_text.insert(1.0, output.get('拟办意见', ''))
        if hasattr(self, 'multi_dept_result_text'):
            self.multi_dept_result_text.insert(1.0, output.get('办理结果', ''))

    def display_special_rule(self, rule):
        """显示特殊情况规则详情"""
        # 清空表单
        self.special_id_entry.configure(state='normal')
        self.special_id_entry.delete(0, tk.END)
        self.special_name_entry.delete(0, tk.END)
        self.special_desc_entry.delete(0, tk.END)
        self.special_title_kw_entry.delete(0, tk.END)
        self.special_filename_kw_entry.delete(0, tk.END)
        self.special_opinion_text.delete(1.0, tk.END)
        self.special_result_text.delete(1.0, tk.END)
        self.modifier_rules_text.delete(1.0, tk.END)
        self.public_place_listbox.delete(0, tk.END)

        # 填充数据
        self.special_id_entry.insert(0, rule.get('id', ''))
        self.special_id_entry.configure(state='readonly')
        self.special_name_entry.insert(0, rule.get('name', ''))
        self.special_desc_entry.insert(0, rule.get('description', ''))
        self.special_priority_spin.set(rule.get('priority', 1))
        self.special_enabled_var.set(rule.get('enabled', True))

        # 判断规则类型：有modifier字段的是修正规则，否则是替换规则
        has_modifier = 'modifier' in rule
        rule_type = 'modifier' if has_modifier else 'replace'
        self.special_type_var.set(rule_type)

        # 条件
        conditions = rule.get('conditions', {})

        # 标题关键词
        title_contains = conditions.get('file_title_contains', [])
        if title_contains:
            self.special_title_kw_entry.insert(0, ', '.join(title_contains))

        # 文件名关键词
        filename_contains = conditions.get('filename_contains', [])
        filename_contains_any = conditions.get('filename_contains_any', [])
        file_title_contains_any = conditions.get('file_title_contains_any', [])

        if filename_contains:
            self.special_filename_kw_entry.insert(0, ', '.join(filename_contains))
        elif filename_contains_any:
            self.special_filename_kw_entry.insert(0, ', '.join(filename_contains_any))

        # 检查是否是公共场所规则
        is_public_place_rule = rule.get('id') == 'special_public_place'

        if is_public_place_rule and file_title_contains_any:
            # 显示公共场所识别码管理界面
            self.public_place_frame.grid()
            for keyword in file_title_contains_any:
                self.public_place_listbox.insert(tk.END, keyword)
        else:
            self.public_place_frame.grid_remove()

        # 根据规则类型显示不同的内容
        if has_modifier:
            # 修正规则：显示modifier内容
            self.output_content_frame.grid_remove()
            self.modifier_content_frame.grid()

            modifier = rule.get('modifier', {})
            opinion_modifiers = modifier.get('拟办意见', [])

            # 格式化显示modifier规则
            modifier_text = []
            for mod in opinion_modifiers:
                action = mod.get('action', '')
                if action == 'remove':
                    modifier_text.append(f"删除：\"{mod.get('text', '')}\"")
                elif action == 'replace':
                    modifier_text.append(f"替换：\"{mod.get('old', '')}\" → \"{mod.get('new', '')}\"")
                elif action == 'append':
                    modifier_text.append(f"追加：\"{mod.get('text', '')}\"")
                elif action == 'prepend':
                    modifier_text.append(f"前置：\"{mod.get('text', '')}\"")

            self.modifier_rules_text.insert(1.0, '\n'.join(modifier_text))
        else:
            # 替换规则：显示output内容
            self.output_content_frame.grid()
            self.modifier_content_frame.grid_remove()

            output = rule.get('output', {})
            self.special_opinion_text.insert(1.0, output.get('拟办意见', ''))
            self.special_result_text.insert(1.0, output.get('办理结果', ''))

    def on_special_type_change(self):
        """特殊规则类型变化时的处理"""
        rule_type = self.special_type_var.get()

        if rule_type == 'modifier':
            # 修正规则：隐藏输出内容，显示修正规则
            self.output_content_frame.grid_remove()
            self.modifier_content_frame.grid()
        else:
            # 替换规则：显示输出内容，隐藏修正规则
            self.output_content_frame.grid()
            self.modifier_content_frame.grid_remove()

    def add_public_place_keyword(self):
        """添加公共场所识别码"""
        keyword = self.public_place_add_entry.get().strip()
        if not keyword:
            messagebox.showwarning("提示", "请输入要添加的识别码")
            return

        # 检查是否已存在
        existing = self.public_place_listbox.get(0, tk.END)
        if keyword in existing:
            messagebox.showwarning("提示", f"识别码 \"{keyword}\" 已存在")
            return

        # 添加到列表
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

            messagebox.showinfo("成功", f"已添加识别码：{keyword}\n配置已自动保存")
        else:
            messagebox.showinfo("成功", f"已添加识别码：{keyword}")

    def remove_public_place_keyword(self):
        """删除公共场所识别码"""
        selection = self.public_place_listbox.curselection()
        if not selection:
            messagebox.showwarning("提示", "请先选择要删除的识别码")
            return

        keyword = self.public_place_listbox.get(selection[0])
        if messagebox.askyesno("确认", f"确定要删除识别码 \"{keyword}\" 吗？"):
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

                messagebox.showinfo("成功", f"已删除识别码：{keyword}\n配置已自动保存")
            else:
                messagebox.showinfo("成功", f"已删除识别码：{keyword}")

    def display_modifier_rule(self, rule):
        """显示标题修正规则详情"""
        # 清空表单
        self.modifier_id_entry.configure(state='normal')
        self.modifier_id_entry.delete(0, tk.END)
        self.modifier_name_entry.delete(0, tk.END)
        self.modifier_desc_entry.delete(0, tk.END)
        self.modifier_filename_kw_entry.delete(0, tk.END)
        self.modifier_text_entry.delete(1.0, tk.END)
        self.modifier_find_entry.delete(0, tk.END)
        self.modifier_replace_entry.delete(0, tk.END)

        # 填充数据
        self.modifier_id_entry.insert(0, rule.get('id', ''))
        self.modifier_id_entry.configure(state='readonly')
        self.modifier_name_entry.insert(0, rule.get('name', ''))
        self.modifier_desc_entry.insert(0, rule.get('description', ''))
        self.modifier_priority_spin.set(rule.get('priority', 1))
        self.modifier_enabled_var.set(rule.get('enabled', True))

        # 条件
        conditions = rule.get('conditions', {})
        filename_contains = conditions.get('filename_contains', [])
        if filename_contains:
            self.modifier_filename_kw_entry.insert(0, ', '.join(filename_contains))

        # action 字段
        action = rule.get('action', {})
        if action:
            # 修正类型
            action_type = action.get('type', 'append')
            self.modifier_action_var.set(action_type)

            # 修正文本
            if action_type in ['append', 'prepend']:
                text = action.get('text', '')
                self.modifier_text_entry.insert(1.0, text)
            elif action_type == 'replace':
                old_text = action.get('old', '')
                new_text = action.get('new', '')
                self.modifier_find_entry.insert(0, old_text)
                self.modifier_replace_entry.insert(0, new_text)

    def apply_rule(self, rule_type):
        """应用规则修改"""
        if rule_type == 'single_dept':
            self.apply_single_dept_rule()
        elif rule_type == 'multi_dept':
            self.apply_multi_dept_rule()
        elif rule_type == 'special':
            self.apply_special_rule()
        elif rule_type == 'modifier':
            self.apply_modifier_rule()

    def apply_single_dept_rule(self):
        """应用单科室规则修改"""
        if not self.current_rule:
            messagebox.showwarning("提示", "请先选择要修改的规则")
            return

        # 收集表单数据
        self.current_rule['name'] = self.single_dept_name_entry.get().strip()
        self.current_rule['description'] = self.single_dept_desc_entry.get().strip()
        self.current_rule['priority'] = int(self.single_dept_priority_spin.get())
        self.current_rule['enabled'] = self.single_dept_enabled_var.get()

        # 条件
        filename_contains = self.single_dept_filename_entry.get().strip()
        if filename_contains:
            self.current_rule['conditions']['filename_contains'] = [
                kw.strip() for kw in filename_contains.split(',') if kw.strip()
            ]

        filename_not_contains = self.single_dept_not_filename_entry.get().strip()
        if filename_not_contains:
            self.current_rule['conditions']['filename_not_contains'] = [
                kw.strip() for kw in filename_not_contains.split(',') if kw.strip()
            ]
        else:
            self.current_rule['conditions'].pop('filename_not_contains', None)

        # 输出
        self.current_rule['output']['拟办意见'] = self.single_dept_opinion_text.get(1.0, tk.END).strip()
        self.current_rule['output']['办理结果'] = self.single_dept_result_text.get(1.0, tk.END).strip()

        # 保存配置
        self.engine.save_rules()

        # 刷新列表
        self.load_single_dept_rules()
        messagebox.showinfo("成功", "规则已应用")

    def apply_multi_dept_rule(self):
        """应用多科室联合规则修改"""
        if not self.current_rule:
            messagebox.showwarning("提示", "请先选择要修改的规则")
            return

        # 收集表单数据
        self.current_rule['name'] = self.multi_dept_name_entry.get().strip()
        self.current_rule['description'] = self.multi_dept_desc_entry.get().strip()
        self.current_rule['priority'] = int(self.multi_dept_priority_spin.get())
        self.current_rule['enabled'] = self.multi_dept_enabled_var.get()

        # 条件 - 科室关键词顺序
        dept_keywords_order = []
        first_keywords = self.multi_dept_first_entry.get().strip()
        if first_keywords:
            dept_keywords_order.append([kw.strip() for kw in first_keywords.split(',') if kw.strip()])

        second_keywords = self.multi_dept_second_entry.get().strip()
        if second_keywords:
            dept_keywords_order.append([kw.strip() for kw in second_keywords.split(',') if kw.strip()])

        if dept_keywords_order:
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

        filename_not_contains = self.multi_dept_not_filename_entry.get().strip()
        if filename_not_contains:
            self.current_rule['conditions']['filename_not_contains'] = [
                kw.strip() for kw in filename_not_contains.split(',') if kw.strip()
            ]
        else:
            self.current_rule['conditions'].pop('filename_not_contains', None)

        # 输出
        self.current_rule['output']['拟办意见'] = self.multi_dept_opinion_text.get(1.0, tk.END).strip()

        # 如果有办理结果字段
        if hasattr(self, 'multi_dept_result_text'):
            self.current_rule['output']['办理结果'] = self.multi_dept_result_text.get(1.0, tk.END).strip()

        # 保存配置
        self.engine.save_rules()

        # 刷新列表
        self.load_multi_dept_rules()
        messagebox.showinfo("成功", "规则已应用")

    def apply_special_rule(self):
        """应用特殊情况规则修改"""
        if not self.current_rule:
            messagebox.showwarning("提示", "请先选择要修改的规则")
            return

        # 收集表单数据
        self.current_rule['name'] = self.special_name_entry.get().strip()
        self.current_rule['description'] = self.special_desc_entry.get().strip()
        self.current_rule['priority'] = int(self.special_priority_spin.get())
        self.current_rule['enabled'] = self.special_enabled_var.get()

        # 规则类型
        rule_type = self.special_type_var.get()

        # 条件
        conditions = {}

        # 标题关键词
        title_contains = self.special_title_kw_entry.get().strip()
        if title_contains:
            conditions['file_title_contains'] = [
                kw.strip() for kw in title_contains.split(',') if kw.strip()
            ]

        # 文件名关键词
        filename_contains = self.special_filename_kw_entry.get().strip()
        if filename_contains:
            # 检查是否是公共场所规则
            is_public_place_rule = self.current_rule.get('id') == 'special_public_place'
            if is_public_place_rule:
                conditions['filename_contains_any'] = [
                    kw.strip() for kw in filename_contains.split(',') if kw.strip()
                ]
                # 从列表框获取公共场所识别码
                public_place_keywords = list(self.public_place_listbox.get(0, tk.END))
                if public_place_keywords:
                    conditions['file_title_contains_any'] = public_place_keywords
            else:
                conditions['filename_contains'] = [
                    kw.strip() for kw in filename_contains.split(',') if kw.strip()
                ]

        self.current_rule['conditions'] = conditions

        # 根据规则类型保存不同的内容
        if rule_type == 'modifier':
            # 修正规则：保存modifier字段
            modifier_text = self.modifier_rules_text.get(1.0, tk.END).strip()
            modifiers = []

            # 解析修正规则文本
            for line in modifier_text.split('\n'):
                line = line.strip()
                if not line:
                    continue

                if line.startswith('删除：'):
                    text = line[3:].strip().strip('"')
                    modifiers.append({'action': 'remove', 'text': text})
                elif line.startswith('替换：'):
                    parts = line[3:].split('→')
                    if len(parts) == 2:
                        old = parts[0].strip().strip('"')
                        new = parts[1].strip().strip('"')
                        modifiers.append({'action': 'replace', 'old': old, 'new': new})
                elif line.startswith('追加：'):
                    text = line[3:].strip().strip('"')
                    modifiers.append({'action': 'append', 'text': text})
                elif line.startswith('前置：'):
                    text = line[3:].strip().strip('"')
                    modifiers.append({'action': 'prepend', 'text': text})

            self.current_rule['modifier'] = {'拟办意见': modifiers}
            # 删除output字段
            self.current_rule.pop('output', None)
        else:
            # 替换规则：保存output字段
            output = {}
            output['拟办意见'] = self.special_opinion_text.get(1.0, tk.END).strip()
            output['办理结果'] = self.special_result_text.get(1.0, tk.END).strip()
            self.current_rule['output'] = output
            # 删除modifier字段
            self.current_rule.pop('modifier', None)

        # 保存配置
        self.engine.save_rules()

        # 刷新列表
        self.load_special_rules()
        messagebox.showinfo("成功", "规则已应用")

    def apply_modifier_rule(self):
        """应用标题修正规则修改"""
        if not self.current_rule:
            messagebox.showwarning("提示", "请先选择要修改的规则")
            return

        # 收集表单数据
        self.current_rule['name'] = self.modifier_name_entry.get().strip()
        self.current_rule['description'] = self.modifier_desc_entry.get().strip()
        self.current_rule['priority'] = int(self.modifier_priority_spin.get())
        self.current_rule['enabled'] = self.modifier_enabled_var.get()

        # 条件
        keyword = self.modifier_filename_kw_entry.get().strip()
        if keyword:
            self.current_rule['conditions']['filename_contains'] = [
                kw.strip() for kw in keyword.split(',') if kw.strip()
            ]

        # action - 根据类型构建
        action_type = self.modifier_action_var.get()
        action = {
            'field': '文件标题',
            'type': action_type
        }

        if action_type in ['append', 'prepend']:
            text = self.modifier_text_entry.get(1.0, tk.END).strip()
            action['text'] = text
        elif action_type == 'replace':
            old_text = self.modifier_find_entry.get().strip()
            new_text = self.modifier_replace_entry.get().strip()
            action['old'] = old_text
            action['new'] = new_text

        self.current_rule['action'] = action

        # 保存配置
        self.engine.save_rules()

        # 刷新列表
        self.load_modifier_rules()
        messagebox.showinfo("成功", "规则已应用")

    def add_rule(self, rule_type):
        """添加新规则"""
        if rule_type == 'single_dept':
            self.add_single_dept_rule()
        elif rule_type == 'multi_dept':
            self.add_multi_dept_rule()
        elif rule_type == 'special':
            self.add_special_rule()
        elif rule_type == 'modifier':
            self.add_modifier_rule()
        else:
            messagebox.showinfo("提示", "该规则类型的添加功能将在后续版本中实现")

    def add_single_dept_rule(self):
        """添加新的单科室规则"""
        # 创建对话框
        dialog = tk.Toplevel(self.window)
        dialog.title("添加单科室规则")
        dialog.geometry("500x300")
        dialog.transient(self.window)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding=20)
        frame.pack(fill=BOTH, expand=YES)

        ttk.Label(frame, text="新建单科室规则", font=(ui_font('微软雅黑'), 12, 'bold')).pack(pady=(0, 15))

        # 规则ID
        id_frame = ttk.Frame(frame)
        id_frame.pack(fill=X, pady=5)
        ttk.Label(id_frame, text="规则ID:", width=12).pack(side=LEFT)
        id_entry = ttk.Entry(id_frame)
        id_entry.pack(side=LEFT, fill=X, expand=YES, padx=(5, 0))
        ttk.Label(frame, text="建议格式: single_xxx", font=(ui_font('微软雅黑'), 9), foreground='gray').pack(anchor=W, padx=(80, 0))

        # 规则名称
        name_frame = ttk.Frame(frame)
        name_frame.pack(fill=X, pady=5)
        ttk.Label(name_frame, text="规则名称:", width=12).pack(side=LEFT)
        name_entry = ttk.Entry(name_frame)
        name_entry.pack(side=LEFT, fill=X, expand=YES, padx=(5, 0))

        # 关键词
        keyword_frame = ttk.Frame(frame)
        keyword_frame.pack(fill=X, pady=5)
        ttk.Label(keyword_frame, text="文件名关键词:", width=12).pack(side=LEFT)
        keyword_entry = ttk.Entry(keyword_frame)
        keyword_entry.pack(side=LEFT, fill=X, expand=YES, padx=(5, 0))

        # 优先级
        priority_frame = ttk.Frame(frame)
        priority_frame.pack(fill=X, pady=5)
        ttk.Label(priority_frame, text="优先级:", width=12).pack(side=LEFT)
        priority_spin = ttk.Spinbox(priority_frame, from_=20, to=29, width=10)
        priority_spin.set(21)
        priority_spin.pack(side=LEFT, padx=(5, 0))

        def create_rule():
            rule_id = id_entry.get().strip()
            rule_name = name_entry.get().strip()
            keyword = keyword_entry.get().strip()

            if not rule_id or not rule_name or not keyword:
                messagebox.showwarning("提示", "请填写所有必填项", parent=dialog)
                return

            # 检查ID是否重复
            existing_ids = [r['id'] for r in self.engine.rules.get('single_dept_rules', [])]
            if rule_id in existing_ids:
                messagebox.showerror("错误", f"规则ID '{rule_id}' 已存在", parent=dialog)
                return

            # 创建新规则
            new_rule = {
                'id': rule_id,
                'name': rule_name,
                'description': f"{rule_name}相关投诉",
                'priority': int(priority_spin.get()),
                'enabled': True,
                'conditions': {
                    'filename_contains': [keyword]
                },
                'output': {
                    '拟办意见': '',
                    '办理结果': ''
                }
            }

            # 添加到规则列表
            if 'single_dept_rules' not in self.engine.rules:
                self.engine.rules['single_dept_rules'] = []
            self.engine.rules['single_dept_rules'].append(new_rule)

            # 刷新列表
            self.load_single_dept_rules()

            dialog.destroy()
            messagebox.showinfo("成功", f"已添加规则: {rule_name}")

        # 按钮
        button_frame = ttk.Frame(frame)
        button_frame.pack(pady=(20, 0))

        ttk.Button(button_frame, text="创建", command=create_rule, bootstyle="success", width=12).pack(side=LEFT, padx=5)
        ttk.Button(button_frame, text="取消", command=dialog.destroy, bootstyle="secondary", width=12).pack(side=LEFT, padx=5)

    def add_multi_dept_rule(self):
        """添加新的多科室联合规则"""
        dialog = tk.Toplevel(self.window)
        dialog.title("添加多科室联合规则")
        dialog.geometry("600x500")
        dialog.transient(self.window)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding=20)
        frame.pack(fill=BOTH, expand=YES)

        ttk.Label(frame, text="新建多科室联合规则", font=(ui_font('微软雅黑'), 12, 'bold')).pack(pady=(0, 15))

        # 规则ID
        id_frame = ttk.Frame(frame)
        id_frame.pack(fill=X, pady=5)
        ttk.Label(id_frame, text="规则ID:", width=15).pack(side=LEFT)
        id_entry = ttk.Entry(id_frame)
        id_entry.pack(side=LEFT, fill=X, expand=YES, padx=(5, 0))
        ttk.Label(frame, text="建议格式: multi_xxx_xxx", font=(ui_font('微软雅黑'), 9), foreground='gray').pack(anchor=W, padx=(105, 0))

        # 规则名称
        name_frame = ttk.Frame(frame)
        name_frame.pack(fill=X, pady=5)
        ttk.Label(name_frame, text="规则名称:", width=15).pack(side=LEFT)
        name_entry = ttk.Entry(name_frame)
        name_entry.pack(side=LEFT, fill=X, expand=YES, padx=(5, 0))

        # 关键词（科室1）
        keyword1_frame = ttk.Frame(frame)
        keyword1_frame.pack(fill=X, pady=5)
        ttk.Label(keyword1_frame, text="科室1关键词:", width=15).pack(side=LEFT)
        keyword1_entry = ttk.Entry(keyword1_frame)
        keyword1_entry.pack(side=LEFT, fill=X, expand=YES, padx=(5, 0))

        # 关键词（科室2）
        keyword2_frame = ttk.Frame(frame)
        keyword2_frame.pack(fill=X, pady=5)
        ttk.Label(keyword2_frame, text="科室2关键词:", width=15).pack(side=LEFT)
        keyword2_entry = ttk.Entry(keyword2_frame)
        keyword2_entry.pack(side=LEFT, fill=X, expand=YES, padx=(5, 0))

        # 优先级
        priority_frame = ttk.Frame(frame)
        priority_frame.pack(fill=X, pady=5)
        ttk.Label(priority_frame, text="优先级:", width=15).pack(side=LEFT)
        priority_spin = ttk.Spinbox(priority_frame, from_=10, to=19, width=10)
        priority_spin.set(10)
        priority_spin.pack(side=LEFT, padx=(5, 0))
        ttk.Label(priority_frame, text="(10-19，数字越小优先级越高)", font=(ui_font('微软雅黑'), 9), foreground='gray').pack(side=LEFT, padx=(10, 0))

        # 拟办意见模板
        ttk.Label(frame, text="拟办意见模板:", font=(ui_font('微软雅黑'), 10)).pack(anchor=W, pady=(10, 2))
        opinion_text = scrolledtext.ScrolledText(frame, width=60, height=4, font=(ui_font('微软雅黑'), 9), wrap=tk.WORD)
        opinion_text.pack(fill=BOTH, expand=YES, pady=(0, 5))
        # 默认模板
        opinion_text.insert(1.0, "请莉莉同志阅，请吕宁同志阅示，请高鹏同志亦阅。\n建议：请{科室1}牵头、{科室2}配合办理，请法监科亦阅。")

        # 办理结果模板
        ttk.Label(frame, text="办理结果模板:", font=(ui_font('微软雅黑'), 10)).pack(anchor=W, pady=(5, 2))
        result_text = scrolledtext.ScrolledText(frame, width=60, height=2, font=(ui_font('微软雅黑'), 9), wrap=tk.WORD)
        result_text.pack(fill=BOTH, expand=YES, pady=(0, 5))
        # 默认模板
        result_text.insert(1.0, "信息中心已转{科室1}、{科室2}。")

        ttk.Label(frame, text="提示：使用 {科室1}、{科室2} 作为占位符", font=(ui_font('微软雅黑'), 9), foreground='blue').pack(anchor=W)

        def create_rule():
            rule_id = id_entry.get().strip()
            rule_name = name_entry.get().strip()
            keyword1 = keyword1_entry.get().strip()
            keyword2 = keyword2_entry.get().strip()

            if not rule_id or not rule_name or not keyword1 or not keyword2:
                messagebox.showwarning("提示", "请填写所有必填项", parent=dialog)
                return

            # 检查ID是否重复
            existing_ids = [r['id'] for r in self.engine.rules.get('multi_dept_rules', [])]
            if rule_id in existing_ids:
                messagebox.showerror("错误", f"规则ID '{rule_id}' 已存在", parent=dialog)
                return

            # 创建新规则 - 按照现有规则结构
            new_rule = {
                'id': rule_id,
                'priority': int(priority_spin.get()),
                'enabled': True,
                'name': rule_name,
                'description': f"{rule_name}联合办理",
                'conditions': {
                    'filename_contains_all': [keyword1, keyword2],
                    'order': [keyword1, keyword2]
                },
                'output': {
                    '拟办意见': opinion_text.get(1.0, tk.END).strip(),
                    '办理结果': result_text.get(1.0, tk.END).strip()
                }
            }

            # 添加到规则列表
            if 'multi_dept_rules' not in self.engine.rules:
                self.engine.rules['multi_dept_rules'] = []
            self.engine.rules['multi_dept_rules'].append(new_rule)

            # 保存配置
            self.engine.save_rules()

            # 刷新列表
            self.load_multi_dept_rules()

            dialog.destroy()
            messagebox.showinfo("成功", f"已添加规则: {rule_name}")

        # 按钮
        button_frame = ttk.Frame(frame)
        button_frame.pack(pady=(10, 0))

        ttk.Button(button_frame, text="创建", command=create_rule, bootstyle="success", width=12).pack(side=LEFT, padx=5)
        ttk.Button(button_frame, text="取消", command=dialog.destroy, bootstyle="secondary", width=12).pack(side=LEFT, padx=5)

    def add_special_rule(self):
        """添加新的特殊情况规则"""
        dialog = tk.Toplevel(self.window)
        dialog.title("添加特殊情况规则")
        dialog.geometry("600x550")
        dialog.transient(self.window)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding=20)
        frame.pack(fill=BOTH, expand=YES)

        ttk.Label(frame, text="新建特殊情况规则", font=(ui_font('微软雅黑'), 12, 'bold')).pack(pady=(0, 15))

        # 规则ID
        id_frame = ttk.Frame(frame)
        id_frame.pack(fill=X, pady=5)
        ttk.Label(id_frame, text="规则ID:", width=15).pack(side=LEFT)
        id_entry = ttk.Entry(id_frame)
        id_entry.pack(side=LEFT, fill=X, expand=YES, padx=(5, 0))
        ttk.Label(frame, text="建议格式: special_xxx_xxx", font=(ui_font('微软雅黑'), 9), foreground='gray').pack(anchor=W, padx=(105, 0))

        # 规则名称
        name_frame = ttk.Frame(frame)
        name_frame.pack(fill=X, pady=5)
        ttk.Label(name_frame, text="规则名称:", width=15).pack(side=LEFT)
        name_entry = ttk.Entry(name_frame)
        name_entry.pack(side=LEFT, fill=X, expand=YES, padx=(5, 0))

        # 标题关键词
        title_keyword_frame = ttk.Frame(frame)
        title_keyword_frame.pack(fill=X, pady=5)
        ttk.Label(title_keyword_frame, text="标题关键词:", width=15).pack(side=LEFT)
        title_keyword_entry = ttk.Entry(title_keyword_frame)
        title_keyword_entry.pack(side=LEFT, fill=X, expand=YES, padx=(5, 0))
        ttk.Label(frame, text="多个关键词用逗号分隔", font=(ui_font('微软雅黑'), 9), foreground='gray').pack(anchor=W, padx=(105, 0))

        # 文件名关键词
        filename_keyword_frame = ttk.Frame(frame)
        filename_keyword_frame.pack(fill=X, pady=5)
        ttk.Label(filename_keyword_frame, text="文件名关键词:", width=15).pack(side=LEFT)
        filename_keyword_entry = ttk.Entry(filename_keyword_frame)
        filename_keyword_entry.pack(side=LEFT, fill=X, expand=YES, padx=(5, 0))
        ttk.Label(frame, text="可选，多个关键词用逗号分隔", font=(ui_font('微软雅黑'), 9), foreground='gray').pack(anchor=W, padx=(105, 0))

        # 排除关键词
        exclude_frame = ttk.Frame(frame)
        exclude_frame.pack(fill=X, pady=5)
        ttk.Label(exclude_frame, text="排除关键词:", width=15).pack(side=LEFT)
        exclude_entry = ttk.Entry(exclude_frame)
        exclude_entry.pack(side=LEFT, fill=X, expand=YES, padx=(5, 0))
        ttk.Label(frame, text="可选，包含这些词的将被排除", font=(ui_font('微软雅黑'), 9), foreground='gray').pack(anchor=W, padx=(105, 0))

        # 优先级
        priority_frame = ttk.Frame(frame)
        priority_frame.pack(fill=X, pady=5)
        ttk.Label(priority_frame, text="优先级:", width=15).pack(side=LEFT)
        priority_spin = ttk.Spinbox(priority_frame, from_=1, to=9, width=10)
        priority_spin.set(1)
        priority_spin.pack(side=LEFT, padx=(5, 0))
        ttk.Label(priority_frame, text="(1-9，数字越小优先级越高)", font=(ui_font('微软雅黑'), 9), foreground='gray').pack(side=LEFT, padx=(10, 0))

        # 拟办意见模板
        ttk.Label(frame, text="拟办意见模板:", font=(ui_font('微软雅黑'), 10)).pack(anchor=W, pady=(10, 2))
        opinion_text = scrolledtext.ScrolledText(frame, width=60, height=4, font=(ui_font('微软雅黑'), 9), wrap=tk.WORD)
        opinion_text.pack(fill=BOTH, expand=YES, pady=(0, 5))

        # 办理结果模板
        ttk.Label(frame, text="办理结果模板:", font=(ui_font('微软雅黑'), 10)).pack(anchor=W, pady=(5, 2))
        result_text = scrolledtext.ScrolledText(frame, width=60, height=2, font=(ui_font('微软雅黑'), 9), wrap=tk.WORD)
        result_text.pack(fill=BOTH, expand=YES, pady=(0, 5))

        def create_rule():
            rule_id = id_entry.get().strip()
            rule_name = name_entry.get().strip()
            title_keywords = title_keyword_entry.get().strip()
            filename_keywords = filename_keyword_entry.get().strip()

            if not rule_id or not rule_name or not title_keywords:
                messagebox.showwarning("提示", "请至少填写规则ID、名称和标题关键词", parent=dialog)
                return

            # 检查ID是否重复
            existing_ids = [r['id'] for r in self.engine.rules.get('special_rules', [])]
            if rule_id in existing_ids:
                messagebox.showerror("错误", f"规则ID '{rule_id}' 已存在", parent=dialog)
                return

            # 构建条件
            conditions = {
                'file_title_contains': [kw.strip() for kw in title_keywords.split(',') if kw.strip()]
            }

            if filename_keywords:
                conditions['filename_contains'] = [kw.strip() for kw in filename_keywords.split(',') if kw.strip()]

            exclude_keywords = exclude_entry.get().strip()
            if exclude_keywords:
                conditions['filename_not_contains'] = [kw.strip() for kw in exclude_keywords.split(',') if kw.strip()]

            # 创建新规则 - 按照现有规则结构
            new_rule = {
                'id': rule_id,
                'priority': int(priority_spin.get()),
                'enabled': True,
                'name': rule_name,
                'description': f"{rule_name}特殊处理",
                'conditions': conditions,
                'output': {
                    '拟办意见': opinion_text.get(1.0, tk.END).strip(),
                    '办理结果': result_text.get(1.0, tk.END).strip()
                }
            }

            # 添加到规则列表
            if 'special_rules' not in self.engine.rules:
                self.engine.rules['special_rules'] = []
            self.engine.rules['special_rules'].append(new_rule)

            # 保存配置
            self.engine.save_rules()

            # 刷新列表
            self.load_special_rules()

            dialog.destroy()
            messagebox.showinfo("成功", f"已添加规则: {rule_name}")

        # 按钮
        button_frame = ttk.Frame(frame)
        button_frame.pack(pady=(10, 0))

        ttk.Button(button_frame, text="创建", command=create_rule, bootstyle="success", width=12).pack(side=LEFT, padx=5)
        ttk.Button(button_frame, text="取消", command=dialog.destroy, bootstyle="secondary", width=12).pack(side=LEFT, padx=5)

    def add_modifier_rule(self):
        """添加新的标题修正规则"""
        dialog = tk.Toplevel(self.window)
        dialog.title("添加标题修正规则")
        dialog.geometry("600x400")
        dialog.transient(self.window)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding=20)
        frame.pack(fill=BOTH, expand=YES)

        ttk.Label(frame, text="新建标题修正规则", font=(ui_font('微软雅黑'), 12, 'bold')).pack(pady=(0, 15))

        # 规则ID
        id_frame = ttk.Frame(frame)
        id_frame.pack(fill=X, pady=5)
        ttk.Label(id_frame, text="规则ID:", width=15).pack(side=LEFT)
        id_entry = ttk.Entry(id_frame)
        id_entry.pack(side=LEFT, fill=X, expand=YES, padx=(5, 0))
        ttk.Label(frame, text="建议格式: modifier_xxx", font=(ui_font('微软雅黑'), 9), foreground='gray').pack(anchor=W, padx=(105, 0))

        # 规则名称
        name_frame = ttk.Frame(frame)
        name_frame.pack(fill=X, pady=5)
        ttk.Label(name_frame, text="规则名称:", width=15).pack(side=LEFT)
        name_entry = ttk.Entry(name_frame)
        name_entry.pack(side=LEFT, fill=X, expand=YES, padx=(5, 0))

        # 触发关键词
        keyword_frame = ttk.Frame(frame)
        keyword_frame.pack(fill=X, pady=5)
        ttk.Label(keyword_frame, text="触发关键词:", width=15).pack(side=LEFT)
        keyword_entry = ttk.Entry(keyword_frame)
        keyword_entry.pack(side=LEFT, fill=X, expand=YES, padx=(5, 0))
        ttk.Label(frame, text="文件名包含这些词时应用此规则，逗号分隔", font=(ui_font('微软雅黑'), 9), foreground='gray').pack(anchor=W, padx=(105, 0))

        # 修正类型
        ttk.Label(frame, text="修正类型:", font=(ui_font('微软雅黑'), 10, 'bold')).pack(anchor=W, pady=(10, 5))

        type_var = tk.StringVar(value="append")
        type_frame = ttk.Frame(frame)
        type_frame.pack(fill=X, pady=5, padx=(105, 0))

        ttk.Radiobutton(type_frame, text="追加文本", variable=type_var, value="append_if_not_exists").pack(side=LEFT, padx=(0, 15))
        ttk.Radiobutton(type_frame, text="替换文本", variable=type_var, value="replace").pack(side=LEFT, padx=(0, 15))
        ttk.Radiobutton(type_frame, text="删除文本", variable=type_var, value="remove").pack(side=LEFT)

        # 修正文本
        ttk.Label(frame, text="修正文本:", font=(ui_font('微软雅黑'), 10)).pack(anchor=W, pady=(10, 2))
        ttk.Label(frame, text="追加模式：要追加的文本 | 删除模式：要删除的文本", font=(ui_font('微软雅黑'), 9), foreground='gray').pack(anchor=W, padx=(0, 0))
        text_entry = ttk.Entry(frame, width=60)
        text_entry.pack(fill=X, pady=(2, 5))

        # 替换模式专用字段
        replace_frame = ttk.LabelFrame(frame, text="替换模式设置", padding=10)
        replace_frame.pack(fill=X, pady=(10, 5))

        ttk.Label(replace_frame, text="查找文本:", font=(ui_font('微软雅黑'), 10)).grid(row=0, column=0, sticky=W, padx=5, pady=5)
        old_text_entry = ttk.Entry(replace_frame, width=40)
        old_text_entry.grid(row=0, column=1, sticky=EW, padx=5, pady=5)

        ttk.Label(replace_frame, text="替换为:", font=(ui_font('微软雅黑'), 10)).grid(row=1, column=0, sticky=W, padx=5, pady=5)
        new_text_entry = ttk.Entry(replace_frame, width=40)
        new_text_entry.grid(row=1, column=1, sticky=EW, padx=5, pady=5)

        replace_frame.columnconfigure(1, weight=1)
        ttk.Label(frame, text="注意：替换模式下使用上方的查找/替换字段", font=(ui_font('微软雅黑'), 9), foreground='blue').pack(anchor=W)

        def create_rule():
            rule_id = id_entry.get().strip()
            rule_name = name_entry.get().strip()
            keywords = keyword_entry.get().strip()
            action_type = type_var.get()

            if not rule_id or not rule_name or not keywords:
                messagebox.showwarning("提示", "请填写规则ID、名称和触发关键词", parent=dialog)
                return

            # 检查ID是否重复
            existing_ids = [r['id'] for r in self.engine.rules.get('title_modifiers', [])]
            if rule_id in existing_ids:
                messagebox.showerror("错误", f"规则ID '{rule_id}' 已存在", parent=dialog)
                return

            # 构建action
            action = {
                'field': '文件标题',
                'type': action_type
            }

            if action_type == 'append_if_not_exists':
                text = text_entry.get().strip()
                if not text:
                    messagebox.showwarning("提示", "请填写要追加的文本", parent=dialog)
                    return
                action['text'] = text
            elif action_type == 'remove':
                text = text_entry.get().strip()
                if not text:
                    messagebox.showwarning("提示", "请填写要删除的文本", parent=dialog)
                    return
                action['text'] = text
            elif action_type == 'replace':
                old_text = old_text_entry.get().strip()
                new_text = new_text_entry.get().strip()
                if not old_text:
                    messagebox.showwarning("提示", "请填写查找文本", parent=dialog)
                    return
                action['old'] = old_text
                action['new'] = new_text

            # 创建新规则 - 按照现有规则结构
            new_rule = {
                'id': rule_id,
                'enabled': True,
                'name': rule_name,
                'description': f"{rule_name}",
                'conditions': {
                    'filename_contains': [kw.strip() for kw in keywords.split(',') if kw.strip()]
                },
                'action': action
            }

            # 添加到规则列表
            if 'title_modifiers' not in self.engine.rules:
                self.engine.rules['title_modifiers'] = []
            self.engine.rules['title_modifiers'].append(new_rule)

            # 保存配置
            self.engine.save_rules()

            # 刷新列表
            self.load_modifier_rules()

            dialog.destroy()
            messagebox.showinfo("成功", f"已添加规则: {rule_name}")

        # 按钮
        button_frame = ttk.Frame(frame)
        button_frame.pack(pady=(10, 0))

        ttk.Button(button_frame, text="创建", command=create_rule, bootstyle="success", width=12).pack(side=LEFT, padx=5)
        ttk.Button(button_frame, text="取消", command=dialog.destroy, bootstyle="secondary", width=12).pack(side=LEFT, padx=5)
        ttk.Label(id_frame, text="规则ID:", width=12).pack(side=LEFT)
        id_entry = ttk.Entry(id_frame)
        id_entry.pack(side=LEFT, fill=X, expand=YES, padx=(5, 0))
        ttk.Label(frame, text="建议格式: modifier_xxx", font=(ui_font('微软雅黑'), 9), foreground='gray').pack(anchor=W, padx=(80, 0))

        # 规则名称
        name_frame = ttk.Frame(frame)
        name_frame.pack(fill=X, pady=5)
        ttk.Label(name_frame, text="规则名称:", width=12).pack(side=LEFT)
        name_entry = ttk.Entry(name_frame)
        name_entry.pack(side=LEFT, fill=X, expand=YES, padx=(5, 0))

        # 匹配模式
        pattern_frame = ttk.Frame(frame)
        pattern_frame.pack(fill=X, pady=5)
        ttk.Label(pattern_frame, text="匹配模式:", width=12).pack(side=LEFT)
        pattern_entry = ttk.Entry(pattern_frame)
        pattern_entry.pack(side=LEFT, fill=X, expand=YES, padx=(5, 0))

        # 替换文本
        replacement_frame = ttk.Frame(frame)
        replacement_frame.pack(fill=X, pady=5)
        ttk.Label(replacement_frame, text="替换文本:", width=12).pack(side=LEFT)
        replacement_entry = ttk.Entry(replacement_frame)
        replacement_entry.pack(side=LEFT, fill=X, expand=YES, padx=(5, 0))

        # 优先级
        priority_frame = ttk.Frame(frame)
        priority_frame.pack(fill=X, pady=5)
        ttk.Label(priority_frame, text="优先级:", width=12).pack(side=LEFT)
        priority_spin = ttk.Spinbox(priority_frame, from_=1, to=99, width=10)
        priority_spin.set(1)
        priority_spin.pack(side=LEFT, padx=(5, 0))

        def create_rule():
            rule_id = id_entry.get().strip()
            rule_name = name_entry.get().strip()
            pattern = pattern_entry.get().strip()
            replacement = replacement_entry.get().strip()

            if not rule_id or not rule_name or not pattern:
                messagebox.showwarning("提示", "请填写所有必填项", parent=dialog)
                return

            # 检查ID是否重复
            existing_ids = [r['id'] for r in self.engine.rules.get('title_modifiers', [])]
            if rule_id in existing_ids:
                messagebox.showerror("错误", f"规则ID '{rule_id}' 已存在", parent=dialog)
                return

            # 创建新规则
            new_rule = {
                'id': rule_id,
                'name': rule_name,
                'description': f"标题修正: {pattern} -> {replacement}",
                'priority': int(priority_spin.get()),
                'enabled': True,
                'pattern': pattern,
                'replacement': replacement
            }

            # 添加到规则列表
            if 'title_modifiers' not in self.engine.rules:
                self.engine.rules['title_modifiers'] = []
            self.engine.rules['title_modifiers'].append(new_rule)

            # 刷新列表
            self.load_modifier_rules()

            dialog.destroy()
            messagebox.showinfo("成功", f"已添加规则: {rule_name}")

        # 按钮
        button_frame = ttk.Frame(frame)
        button_frame.pack(pady=(20, 0))

        ttk.Button(button_frame, text="创建", command=create_rule, bootstyle="success", width=12).pack(side=LEFT, padx=5)
        ttk.Button(button_frame, text="取消", command=dialog.destroy, bootstyle="secondary", width=12).pack(side=LEFT, padx=5)

    def delete_rule(self, rule_type):
        """删除规则"""
        if rule_type == 'single_dept':
            self.delete_single_dept_rule()
        elif rule_type == 'multi_dept':
            self.delete_multi_dept_rule()
        elif rule_type == 'special':
            self.delete_special_rule()
        elif rule_type == 'modifier':
            self.delete_modifier_rule()
        else:
            messagebox.showinfo("提示", "该规则类型的删除功能将在后续版本中实现")

    def delete_single_dept_rule(self):
        """删除单科室规则"""
        selection = self.single_dept_listbox.curselection()
        if not selection:
            messagebox.showwarning("提示", "请先选择要删除的规则")
            return

        index = selection[0]
        rules = sorted(
            self.engine.rules.get('single_dept_rules', []),
            key=lambda x: x.get('priority', 999)
        )

        if index >= len(rules):
            return

        rule = rules[index]
        rule_name = rule.get('name', rule.get('id'))

        if messagebox.askyesno("确认删除", f"确定要删除规则 '{rule_name}' 吗？\n\n此操作不可撤销。"):
            # 从列表中移除
            self.engine.rules['single_dept_rules'].remove(rule)

            # 刷新列表
            self.load_single_dept_rules()

            # 清空详情
            self.current_rule = None
            self.create_single_dept_form()

            messagebox.showinfo("成功", f"已删除规则: {rule_name}")

    def delete_multi_dept_rule(self):
        """删除多科室联合规则"""
        selection = self.multi_dept_listbox.curselection()
        if not selection:
            messagebox.showwarning("提示", "请先选择要删除的规则")
            return

        index = selection[0]
        rules = sorted(
            self.engine.rules.get('multi_dept_rules', []),
            key=lambda x: x.get('priority', 999)
        )

        if index >= len(rules):
            return

        rule = rules[index]
        rule_name = rule.get('name', rule.get('id'))

        if messagebox.askyesno("确认删除", f"确定要删除规则 '{rule_name}' 吗？\n\n此操作不可撤销。"):
            self.engine.rules['multi_dept_rules'].remove(rule)
            self.load_multi_dept_rules()
            self.current_rule = None
            messagebox.showinfo("成功", f"已删除规则: {rule_name}")

    def delete_special_rule(self):
        """删除特殊情况规则"""
        selection = self.special_listbox.curselection()
        if not selection:
            messagebox.showwarning("提示", "请先选择要删除的规则")
            return

        index = selection[0]
        rules = sorted(
            self.engine.rules.get('special_rules', []),
            key=lambda x: x.get('priority', 999)
        )

        if index >= len(rules):
            return

        rule = rules[index]
        rule_name = rule.get('name', rule.get('id'))

        if messagebox.askyesno("确认删除", f"确定要删除规则 '{rule_name}' 吗？\n\n此操作不可撤销。"):
            self.engine.rules['special_rules'].remove(rule)
            self.load_special_rules()
            self.current_rule = None
            messagebox.showinfo("成功", f"已删除规则: {rule_name}")

    def delete_modifier_rule(self):
        """删除标题修正规则"""
        selection = self.modifier_listbox.curselection()
        if not selection:
            messagebox.showwarning("提示", "请先选择要删除的规则")
            return

        index = selection[0]
        rules = sorted(
            self.engine.rules.get('title_modifiers', []),
            key=lambda x: x.get('priority', 999)
        )

        if index >= len(rules):
            return

        rule = rules[index]
        rule_name = rule.get('name', rule.get('id'))

        if messagebox.askyesno("确认删除", f"确定要删除规则 '{rule_name}' 吗？\n\n此操作不可撤销。"):
            self.engine.rules['title_modifiers'].remove(rule)
            self.load_modifier_rules()
            self.current_rule = None
            messagebox.showinfo("成功", f"已删除规则: {rule_name}")

    def move_rule_up(self, rule_type):
        """上移规则"""
        if rule_type == 'single_dept':
            self.move_single_dept_rule(-1)
        else:
            messagebox.showinfo("提示", "该规则类型的移动功能将在后续版本中实现")

    def move_rule_down(self, rule_type):
        """下移规则"""
        if rule_type == 'single_dept':
            self.move_single_dept_rule(1)
        else:
            messagebox.showinfo("提示", "该规则类型的移动功能将在后续版本中实现")

    def move_single_dept_rule(self, direction):
        """
        移动单科室规则

        Args:
            direction: -1向上移动，1向下移动
        """
        selection = self.single_dept_listbox.curselection()
        if not selection:
            messagebox.showwarning("提示", "请先选择要移动的规则")
            return

        index = selection[0]
        rules = sorted(
            self.engine.rules.get('single_dept_rules', []),
            key=lambda x: x.get('priority', 999)
        )

        if index >= len(rules):
            return

        # 检查边界
        new_index = index + direction
        if new_index < 0 or new_index >= len(rules):
            return

        # 交换优先级
        current_rule = rules[index]
        target_rule = rules[new_index]

        current_priority = current_rule.get('priority', 21)
        target_priority = target_rule.get('priority', 21)

        current_rule['priority'] = target_priority
        target_rule['priority'] = current_priority

        # 刷新列表
        self.load_single_dept_rules()

        # 重新选中移动后的规则
        self.single_dept_listbox.selection_clear(0, tk.END)
        self.single_dept_listbox.selection_set(new_index)
        self.single_dept_listbox.see(new_index)

    def test_current_rule(self):
        """测试当前规则"""
        # 创建测试窗口
        test_window = RuleTestWindow(self.window, self.engine)

    def save_rules(self):
        """保存规则配置"""
        if self.engine.save_rules():
            messagebox.showinfo("成功", "规则配置已保存")
        else:
            messagebox.showerror("错误", "保存规则配置失败")

    def import_rules(self):
        """导入规则配置"""
        from tkinter import filedialog

        filename = filedialog.askopenfilename(
            parent=self.window,
            title="选择规则配置文件",
            filetypes=[("JSON配置文件", "*.json"), ("所有文件", "*.*")]
        )

        if not filename:
            return

        try:
            import shutil
            from datetime import datetime

            # 备份当前配置
            if os.path.exists(self.engine.config_file):
                backup_file = self.engine.config_file + f".backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                shutil.copy(self.engine.config_file, backup_file)
                print(f"已备份当前配置: {backup_file}")

            # 读取导入文件
            with open(filename, 'r', encoding='utf-8-sig') as f:
                import_config = json.load(f)

            # 验证配置格式
            if 'rules' not in import_config or 'keywords' not in import_config:
                messagebox.showerror("错误", "配置文件格式不正确")
                return

            # 复制配置文件
            shutil.copy(filename, self.engine.config_file)

            # 重新加载规则
            self.engine.load_rules()

            # 刷新界面
            self.load_rules()

            messagebox.showinfo("成功", "规则配置已导入\n\n原配置已备份")

        except Exception as e:
            messagebox.showerror("错误", f"导入规则配置失败:\n{e}")

    def export_rules(self):
        """导出规则配置"""
        from tkinter import filedialog
        from datetime import datetime

        default_filename = f"filling_rules_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"

        filename = filedialog.asksaveasfilename(
            parent=self.window,
            title="导出规则配置",
            defaultextension=".json",
            initialfile=default_filename,
            filetypes=[("JSON配置文件", "*.json"), ("所有文件", "*.*")]
        )

        if not filename:
            return

        try:
            import shutil

            # 先保存当前配置
            if not self.engine.save_rules():
                messagebox.showerror("错误", "保存当前配置失败")
                return

            # 复制配置文件
            shutil.copy(self.engine.config_file, filename)

            messagebox.showinfo("成功", f"规则配置已导出到:\n{filename}")

        except Exception as e:
            messagebox.showerror("错误", f"导出规则配置失败:\n{e}")

    def check_conflicts(self):
        """检查规则冲突"""
        conflicts = []

        # 获取所有规则
        all_rules = []

        # 添加特殊情况规则
        for rule in self.engine.rules.get('special_rules', []):
            if rule.get('enabled', True):
                all_rules.append({
                    'type': 'special',
                    'rule': rule,
                    'priority': rule.get('priority', 1)
                })

        # 添加多科室联合规则
        for rule in self.engine.rules.get('multi_dept_rules', []):
            if rule.get('enabled', True):
                all_rules.append({
                    'type': 'multi_dept',
                    'rule': rule,
                    'priority': rule.get('priority', 11)
                })

        # 添加单科室规则
        for rule in self.engine.rules.get('single_dept_rules', []):
            if rule.get('enabled', True):
                all_rules.append({
                    'type': 'single_dept',
                    'rule': rule,
                    'priority': rule.get('priority', 21)
                })

        # 按优先级排序
        all_rules.sort(key=lambda x: x['priority'])

        # 检查关键词冲突
        for i in range(len(all_rules)):
            for j in range(i + 1, len(all_rules)):
                rule1 = all_rules[i]
                rule2 = all_rules[j]

                # 获取两个规则的关键词
                keywords1 = set(rule1['rule'].get('conditions', {}).get('filename_contains', []))
                keywords2 = set(rule2['rule'].get('conditions', {}).get('filename_contains', []))

                # 检查是否有交集
                common_keywords = keywords1 & keywords2

                if common_keywords:
                    # 检查排除关键词
                    exclude1 = set(rule1['rule'].get('conditions', {}).get('filename_not_contains', []))
                    exclude2 = set(rule2['rule'].get('conditions', {}).get('filename_not_contains', []))

                    # 如果一个规则的包含关键词在另一个规则的排除关键词中，则不冲突
                    if keywords1 & exclude2 or keywords2 & exclude1:
                        continue

                    conflicts.append({
                        'rule1': rule1,
                        'rule2': rule2,
                        'common_keywords': list(common_keywords),
                        'severity': 'high' if rule1['priority'] == rule2['priority'] else 'medium'
                    })

        # 显示冲突检测结果
        self.show_conflict_results(conflicts)

    def show_conflict_results(self, conflicts):
        """显示冲突检测结果"""
        # 创建结果窗口
        result_window = tk.Toplevel(self.window)
        result_window.title("规则冲突检测结果")
        result_window.geometry("900x600")
        result_window.transient(self.window)

        frame = ttk.Frame(result_window, padding=20)
        frame.pack(fill=BOTH, expand=YES)

        # 标题
        title_frame = ttk.Frame(frame)
        title_frame.pack(fill=X, pady=(0, 15))

        if conflicts:
            ttk.Label(
                title_frame,
                text=f"⚠️ 检测到 {len(conflicts)} 个潜在冲突",
                font=(ui_font('微软雅黑'), 14, 'bold'),
                foreground='#F97316'
            ).pack(side=LEFT)
        else:
            ttk.Label(
                title_frame,
                text="✓ 未检测到规则冲突",
                font=(ui_font('微软雅黑'), 14, 'bold'),
                foreground='#22C55E'
            ).pack(side=LEFT)

        # 说明文本
        info_frame = ttk.Frame(frame)
        info_frame.pack(fill=X, pady=(0, 10))

        info_text = """
冲突说明：
• 高严重度（红色）：优先级相同的规则使用了相同的关键词，可能导致不确定的匹配结果
• 中严重度（橙色）：不同优先级的规则使用了相同的关键词，高优先级规则会先匹配
• 建议：为冲突规则添加排除关键词或调整优先级来避免不期望的匹配行为
        """

        ttk.Label(
            info_frame,
            text=info_text.strip(),
            font=(ui_font('微软雅黑'), 9),
            foreground='#64748B',
            justify=LEFT
        ).pack(anchor=W)

        # 冲突列表
        list_frame = ttk.Frame(frame)
        list_frame.pack(fill=BOTH, expand=YES, pady=(10, 0))

        # 创建滚动文本框
        text_scroll = ttk.Scrollbar(list_frame)
        text_scroll.pack(side=RIGHT, fill=Y)

        conflict_text = tk.Text(
            list_frame,
            wrap=tk.WORD,
            font=(ui_font('Consolas'), 10),
            yscrollcommand=text_scroll.set
        )
        conflict_text.pack(side=LEFT, fill=BOTH, expand=YES)
        text_scroll.config(command=conflict_text.yview)

        # 配置标签样式
        conflict_text.tag_config('high', foreground='#DC2626', font=(ui_font('微软雅黑'), 10, 'bold'))
        conflict_text.tag_config('medium', foreground='#F97316', font=(ui_font('微软雅黑'), 10, 'bold'))
        conflict_text.tag_config('rule_name', foreground='#0EA5E9', font=(ui_font('微软雅黑'), 10, 'bold'))
        conflict_text.tag_config('keyword', foreground='#8B5CF6', font=(ui_font('Consolas'), 10))
        conflict_text.tag_config('priority', foreground='#64748B', font=(ui_font('微软雅黑'), 9))

        if conflicts:
            for idx, conflict in enumerate(conflicts, 1):
                rule1 = conflict['rule1']['rule']
                rule2 = conflict['rule2']['rule']
                severity = conflict['severity']
                common = ', '.join(conflict['common_keywords'])

                # 严重度标记
                severity_text = "【高严重度】" if severity == 'high' else "【中严重度】"
                conflict_text.insert(tk.END, f"\n{idx}. {severity_text}\n", severity)

                # 规则1
                conflict_text.insert(tk.END, "规则1: ", 'priority')
                conflict_text.insert(tk.END, f"{rule1.get('name', rule1.get('id'))}", 'rule_name')
                conflict_text.insert(tk.END, f" (优先级: {conflict['rule1']['priority']}, ", 'priority')
                conflict_text.insert(tk.END, f"类型: {conflict['rule1']['type']})\n", 'priority')

                # 规则2
                conflict_text.insert(tk.END, "规则2: ", 'priority')
                conflict_text.insert(tk.END, f"{rule2.get('name', rule2.get('id'))}", 'rule_name')
                conflict_text.insert(tk.END, f" (优先级: {conflict['rule2']['priority']}, ", 'priority')
                conflict_text.insert(tk.END, f"类型: {conflict['rule2']['type']})\n", 'priority')

                # 冲突关键词
                conflict_text.insert(tk.END, "共同关键词: ", 'priority')
                conflict_text.insert(tk.END, f"{common}\n", 'keyword')

                # 建议
                if severity == 'high':
                    conflict_text.insert(tk.END, "⚠️ 建议: 这两个规则优先级相同且匹配条件重叠，建议调整优先级或添加排除关键词\n")
                else:
                    conflict_text.insert(tk.END, f"ℹ️ 说明: 当文件名同时匹配这两个规则时，优先级 {conflict['rule1']['priority']} 的规则会先匹配\n")

                conflict_text.insert(tk.END, "-" * 80 + "\n")
        else:
            conflict_text.insert(tk.END, "\n✓ 所有规则都没有冲突，配置良好！\n\n")
            conflict_text.insert(tk.END, "规则引擎会按照以下顺序匹配：\n")
            conflict_text.insert(tk.END, "1. 特殊情况规则（优先级 1-9）\n")
            conflict_text.insert(tk.END, "2. 多科室联合规则（优先级 10-19）\n")
            conflict_text.insert(tk.END, "3. 单科室规则（优先级 20-29）\n")

        conflict_text.config(state=tk.DISABLED)

        # 关闭按钮
        button_frame = ttk.Frame(frame)
        button_frame.pack(pady=(15, 0))

        ttk.Button(
            button_frame,
            text="关闭",
            command=result_window.destroy,
            bootstyle="secondary",
            width=15
        ).pack()

    def reset_to_default(self):
        """重置为默认配置"""
        if not messagebox.askyesno("确认", "确定要重置为默认配置吗？\n\n当前配置将被覆盖！\n（原配置会自动备份）"):
            return

        try:
            import shutil
            from datetime import datetime

            # 备份当前配置
            if os.path.exists(self.engine.config_file):
                backup_file = self.engine.config_file + f".backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                shutil.copy(self.engine.config_file, backup_file)
                print(f"已备份当前配置: {backup_file}")

            # 复制默认配置
            default_file = get_resource_path(os.path.join('AppData', 'filling_rules_default.json'))
            if not os.path.exists(default_file):
                messagebox.showerror("错误", "默认配置文件不存在")
                return

            shutil.copy(default_file, self.engine.config_file)

            # 重新加载规则
            self.engine.load_rules()

            # 刷新界面
            self.load_rules()

            messagebox.showinfo("成功", "已重置为默认配置\n\n原配置已备份")

        except Exception as e:
            messagebox.showerror("错误", f"重置配置失败:\n{e}")

    def on_closing(self):
        """关闭窗口"""
        self.window.destroy()


class RuleTestWindow:
    """规则测试窗口"""

    def __init__(self, parent, engine):
        """
        初始化测试窗口

        Args:
            parent: 父窗口
            engine: 规则引擎实例
        """
        self.parent = parent
        self.engine = engine

        # 创建窗口
        self.window = tk.Toplevel(parent)
        self.window.title("规则测试")
        self.window.geometry("700x600")

        self.create_widgets()

    def create_widgets(self):
        """创建界面控件"""
        main_frame = ttk.Frame(self.window, padding=20)
        main_frame.pack(fill=BOTH, expand=YES)

        # 标题
        ttk.Label(
            main_frame,
            text="🧪 规则测试",
            font=(ui_font('微软雅黑'), 14, 'bold')
        ).pack(fill=X, pady=(0, 15))

        # 测试输入
        input_frame = ttk.LabelFrame(main_frame, text="测试输入", padding=10)
        input_frame.pack(fill=X, pady=(0, 10))

        ttk.Label(input_frame, text="文件名:", font=(ui_font('微软雅黑'), 10)).grid(
            row=0, column=0, sticky=W, padx=5, pady=5
        )
        self.filename_entry = ttk.Entry(input_frame, width=50)
        self.filename_entry.grid(row=0, column=1, sticky=EW, padx=5, pady=5)
        self.filename_entry.insert(0, "热线工单_医政.doc")

        ttk.Label(input_frame, text="文件标题:", font=(ui_font('微软雅黑'), 10)).grid(
            row=1, column=0, sticky=W, padx=5, pady=5
        )
        self.title_entry = ttk.Entry(input_frame, width=50)
        self.title_entry.grid(row=1, column=1, sticky=EW, padx=5, pady=5)
        self.title_entry.insert(0, "关于XX医院收费问题的投诉")

        input_frame.columnconfigure(1, weight=1)

        # 测试按钮
        ttk.Button(
            main_frame,
            text="▶ 开始测试",
            command=self.run_test,
            bootstyle="success",
            width=20
        ).pack(pady=10)

        # 测试结果
        result_frame = ttk.LabelFrame(main_frame, text="测试结果", padding=10)
        result_frame.pack(fill=BOTH, expand=YES)

        self.result_text = scrolledtext.ScrolledText(
            result_frame,
            width=60,
            height=15,
            font=(ui_font('微软雅黑'), 10),
            wrap=tk.WORD
        )
        self.result_text.pack(fill=BOTH, expand=YES)

    def run_test(self):
        """执行规则测试"""
        filename = self.filename_entry.get().strip()
        title = self.title_entry.get().strip()

        if not filename:
            messagebox.showwarning("提示", "请输入文件名")
            return

        # 清空结果
        self.result_text.delete(1.0, tk.END)

        # 执行测试
        result = self.engine.test_rule(filename, title)

        # 显示结果
        if result['matched']:
            self.result_text.insert(tk.END, "✓ 匹配成功\n\n", "success")
            self.result_text.insert(tk.END, "匹配结果：\n", "bold")
            self.result_text.insert(tk.END, "=" * 50 + "\n\n")

            output = result['output']
            for key, value in output.items():
                self.result_text.insert(tk.END, f"{key}：\n", "bold")
                self.result_text.insert(tk.END, f"{value}\n\n")
        else:
            self.result_text.insert(tk.END, "✗ 未匹配任何规则\n\n", "error")
            self.result_text.insert(tk.END, result['message'])

        # 配置标签样式
        self.result_text.tag_config("success", foreground="green", font=(ui_font('微软雅黑'), 11, 'bold'))
        self.result_text.tag_config("error", foreground="red", font=(ui_font('微软雅黑'), 11, 'bold'))
        self.result_text.tag_config("bold", font=(ui_font('微软雅黑'), 10, 'bold'))


if __name__ == '__main__':
    # 测试代码
    root = ttk.Window(themename="litera")
    root.withdraw()  # 隐藏主窗口

    app = RuleManagementWindow(root)

    root.mainloop()
