#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
投诉批办单自动生成应用
自动处理PDF和Word文件，生成规范的投诉批办单
"""

import os
import json
import glob
import shutil
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path
import pdfplumber
from docx import Document
import re
import platform_compat as pc
from platform_compat import (IS_WINDOWS, get_app_dir, get_appdata_path,
                             get_resource_path, get_user_data_path, ui_font)

# 导入新的规则引擎
try:
    from rule_engine import RuleEngine, KeywordMatcher
    USE_NEW_RULE_ENGINE = True
    print("使用新规则引擎")
except ImportError:
    USE_NEW_RULE_ENGINE = False
    print("警告: 无法导入新规则引擎，使用旧规则系统")


class CounterManager:
    """编号管理器：持久化存储文件编号（支持按年份管理）"""
    
    def __init__(self, config_file='counter.json', year=None):
        """
        初始化编号管理器
        
        Args:
            config_file: 配置文件名（默认counter.json，用于向后兼容）
            year: 年份（4位数字），如果指定则使用counter-YYYY.json格式
        """
        if year:
            # 使用年份特定的配置文件
            config_file = f'counter-{year}.json'
        
        self.config_file = get_appdata_path(config_file)
        self.year = year
        self.file_number = 1
        self.receipt_number = 1
        self.load()
    
    def load(self):
        """从JSON文件加载编号"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8-sig') as f:
                    data = json.load(f)
                    self.file_number = data.get('file_number', 1)
                    self.receipt_number = data.get('receipt_number', 1)
                print(f"加载编号配置: 文件编号={self.file_number}, 收文编号={self.receipt_number}")
            except Exception as e:
                print(f"加载配置文件失败: {e}，使用默认编号")
        else:
            print(f"配置文件不存在，使用初始编号: {self.config_file}")
    
    def save(self):
        """保存编号到JSON文件"""
        try:
            data = {
                'file_number': self.file_number,
                'receipt_number': self.receipt_number
            }
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            print(f"保存编号配置成功")
        except Exception as e:
            print(f"保存配置文件失败: {e}")
    
    def get_next_file_number(self):
        """获取下一个文件编号并递增"""
        current = self.file_number
        self.file_number += 1
        self.save()
        return current
    
    def get_next_receipt_number(self):
        """获取下一个收文编号并递增"""
        current = self.receipt_number
        self.receipt_number += 1
        self.save()
        return current
    
    def set_numbers(self, file_number, receipt_number):
        """设置编号（用于编号管理界面）"""
        self.file_number = file_number
        self.receipt_number = receipt_number
        self.save()
    
    def rollback_file_number(self):
        """回滚文件编号（处理失败时使用）"""
        if self.file_number > 1:
            self.file_number -= 1
            self.save()
    
    def rollback_receipt_number(self):
        """回滚收文编号（处理失败时使用）"""
        if self.receipt_number > 1:
            self.receipt_number -= 1
            self.save()


class WordMLExtractor:
    """
    Word 2003 XML（WordML）提取器。

    现场发现的真实数据：部分「热线工单」导出文件扩展名叫 .docx / .doc，
    文件内容却是以 ``<?xml ...?>`` 开头的 WordML 单文件（w:wordDocument），
    不是 OOXML 的 ZIP 包。python-docx 只能读 ZIP，遇到这类文件直接
    PackageNotFoundError，界面就报「提取数据失败」。本类按解析后的表格
    行列提取字段，不依赖 Word / LibreOffice，麒麟、统信上都能直接跑。
    """

    W = '{http://schemas.microsoft.com/office/word/2003/wordml}'

    @staticmethod
    def is_wordml(path):
        """按文件头判断，不看扩展名（扩展名经常是错的）。"""
        try:
            with open(path, 'rb') as f:
                head = f.read(512).lstrip().lower()
            return head.startswith(b'<?xml') and b'worddocument' in head
        except Exception:
            return False

    @classmethod
    def _tables(cls, path):
        """返回 [[[单元格文本, ...], ...], ...]：表 -> 行 -> 列。"""
        with open(path, 'rb') as f:
            raw = f.read()
        root = ET.fromstring(raw)
        result = []
        for tbl in root.iter(cls.W + 'tbl'):
            rows = []
            for tr in tbl.findall(cls.W + 'tr'):
                cells = []
                for tc in tr.findall(cls.W + 'tc'):
                    text = ''.join(t.text or '' for t in tc.iter(cls.W + 't'))
                    cells.append(text.replace('\r', '').replace('\x07', '').strip())
                rows.append(cells)
            if rows:
                result.append(rows)
        return result

    @classmethod
    def extract_table_data(cls, path):
        """
        从 WordML 工单表提取字段，输出兼容已有生成器的键：
            来文字号 / 文件标题 / 标题 / 来源 / 姓名 / 诉求内容 / 联系电话 / 事发地
        找不到任何可用字段时返回 None（让调用方继续走 LO 转换兜底）。
        """
        try:
            tables = cls._tables(path)
        except Exception as e:
            print(f'WordML 解析失败 {path}: {e}')
            return None

        if not tables:
            print(f'WordML 没有表格: {path}')
            return None

        rows = tables[0]
        data = {}

        # 1) 标签 -> 值 的键值对（同一行里「标签」后面紧跟它的值）
        labels = {
            '工单编号': '工单编号', '诉求标题': '诉求标题', '来源渠道': '来源渠道',
            '姓名': '姓名', '诉求内容': '诉求内容', '事发地': '事发地',
            '联系电话': '联系电话', '工单类型': '工单类型', '责任部门': '责任部门',
            '登记时间': '登记时间', '办理期限': '办理期限',
        }
        for row in rows:
            for i, cell in enumerate(row):
                key = labels.get(cell)
                if key and i + 1 < len(row) and row[i + 1]:
                    data.setdefault(key, row[i + 1])
                # 兼容「标签：值」写在同一格里
                for cn in labels:
                    if cell.startswith(cn + '：') or cell.startswith(cn + ':'):
                        val = cell.split('：', 1)[-1].split(':', 1)[-1].strip()
                        if val:
                            data.setdefault(cn, val)

        # 2) 供 WordGenerator.generate_from_word / generate_from_doc 使用的标准键
        title = data.get('诉求标题') or ''
        content = data.get('诉求内容') or ''
        number = data.get('工单编号') or ''
        source = data.get('来源渠道') or ''
        name = data.get('姓名') or ''

        # 标题为空时用诉求内容首句兜底，避免生成空白批办单
        if not title and content:
            title = content.split('。')[0][:120]

        data['来文字号'] = number
        data['文件标题'] = title
        data['标题'] = title
        data['来源'] = source
        data['姓名'] = name

        if not (title or number or name):
            print(f'WordML 未识别到有效字段: {path}')
            return None

        print(f'WordML提取成功: 工单编号={number}, 标题={title}, 来源={source}, 姓名={name}')
        return data


def is_wordml_file(path):
    """模块级便捷函数。"""
    return WordMLExtractor.is_wordml(path)


def prefer_wps():
    """
    是否优先用 WPS 转换老 .doc。

    环境变量 COMPLAINTS_WPS_PREFERENCE：
        auto  默认 —— 检测到 WPS 就用，没装就走 LibreOffice
        wps   强制 WPS（没装则失败，便于现场验证 WPS 通道）
        off   禁用 WPS，只走 LibreOffice
    也可在 AppData/config.json 里写 {"wps_preference": "auto|wps|off"}。
    """
    pref = (os.environ.get('COMPLAINTS_WPS_PREFERENCE') or '').strip().lower()
    if pref not in ('auto', 'wps', 'off'):
        pref = 'auto'
        try:
            cfg = os.path.join(get_appdata_path('config.json'))
            if os.path.isfile(cfg):
                with open(cfg, 'r', encoding='utf-8-sig') as f:
                    val = (json.load(f).get('wps_preference') or '').strip().lower()
                if val in ('auto', 'wps', 'off'):
                    pref = val
        except Exception:
            pass
    if pref == 'off':
        return False
    try:
        import wps_compat
        if wps_compat.wps_available():
            return True
    except Exception:
        pass
    return pref == 'wps'


def wps_status():
    """给启动日志/诊断脚本用的一行说明。"""
    try:
        import wps_compat
        return wps_compat.diagnose()
    except Exception as e:
        return 'WPS: 模块不可用 (%s)' % e


class PDFExtractor:
    """PDF数据提取器"""
    
    @staticmethod
    def extract_table_data(pdf_path):
        """
        从PDF提取表格数据（自动检测表格在第一页或第二页）
        返回: dict包含提取的数据
        """
        try:
            with pdfplumber.open(pdf_path) as pdf:
                if len(pdf.pages) == 0:
                    print(f"警告: {pdf_path} 没有页面")
                    return None

                # 依次尝试每一页，找到有表格的页面
                table = None
                for page_idx, page in enumerate(pdf.pages):
                    tables = page.extract_tables()
                    if tables and len(tables) > 0:
                        table = tables[0]
                        print(f"PDF在第{page_idx + 1}页找到表格")
                        break

                if table is None:
                    print(f"警告: {pdf_path} 所有页面均未找到表格")
                    return None

                # 提取数据（注意：索引从0开始）
                data = {}

                # 第1行第2列 → 来文字号
                # 实际数据可能在第3列（索引2）
                if len(table) > 0 and len(table[0]) > 2:
                    # 优先取第3列，如果为空则取第2列
                    data['来文字号'] = table[0][2] if table[0][2] else (table[0][1] if len(table[0]) > 1 and table[0][1] else "")
                elif len(table) > 0 and len(table[0]) > 1:
                    data['来文字号'] = table[0][1] if table[0][1] else ""
                else:
                    data['来文字号'] = ""

                # 第3行第3列 → 文件标题
                if len(table) > 2 and len(table[2]) > 2:
                    data['文件标题'] = table[2][2] if table[2][2] else ""
                else:
                    data['文件标题'] = ""

                # pdfplumber 的空单元格会是 None，这里统一成字符串：
                # 下游规则引擎做 `in` 判断、生成器做 str 操作，None 会直接抛异常，
                # 表现为「PDF 不填拟办意见/办理情况」。
                for _k in ('来文字号', '文件标题'):
                    _v = data.get(_k)
                    data[_k] = '' if _v is None else str(_v).strip()

                print(f"PDF提取成功: 来文字号={data['来文字号']}, 文件标题={data['文件标题']}")
                return data

        except Exception as e:
            print(f"提取PDF数据失败 {pdf_path}: {e}")
            return None


class DOCExtractor:
    """
    DOC文件数据提取器。
    Windows 用 Word COM（与旧版完全一致）；统信/麒麟用 LibreOffice/WPS 把 .doc
    转成 .docx 后用 python-docx 读取表格。扩展名为 .doc 但实际是 OOXML 的文件
    （系统内“重命名为.doc”功能产生的）直接用 python-docx 读取。
    """

    _word_instance = None
    _com_initialized = False

    @staticmethod
    def _clean(text):
        return (text or '').replace('\r\x07', '').replace('\x07', '').strip()

    @staticmethod
    def _from_document(doc):
        """从 python-docx Document 对象提取第一张表的来文字号/文件标题。"""
        if not doc.tables:
            return None
        table = doc.tables[0]
        data = {'来文字号': '', '文件标题': ''}
        try:
            data['来文字号'] = DOCExtractor._clean(table.cell(0, 1).text)
        except Exception:
            pass
        try:
            data['文件标题'] = DOCExtractor._clean(table.cell(2, 2).text)
        except Exception:
            pass
        # 两字段都空说明不是目标表格，交给后续 WPS / LibreOffice 兜底
        return data if (data['来文字号'] or data['文件标题']) else None

    @staticmethod
    def _from_rows(rows):
        """从二维数组（WPS 转换后读出的表格）取同样的两个字段。"""
        if not rows:
            return None

        def cell(r, c):
            try:
                return DOCExtractor._clean(rows[r][c])
            except Exception:
                return ''

        data = {'来文字号': cell(0, 1), '文件标题': cell(2, 2)}
        return data if (data['来文字号'] or data['文件标题']) else None
    @classmethod
    def get_word_instance(cls):
        """获取或创建Word实例（仅 Windows COM 分支使用）。"""
        if not IS_WINDOWS:
            return None
        if cls._word_instance is None:
            import pythoncom
            import win32com.client
            if not cls._com_initialized:
                pythoncom.CoInitialize()
                cls._com_initialized = True
            cls._word_instance = win32com.client.Dispatch("Word.Application")
            cls._word_instance.Visible = False
            cls._word_instance.DisplayAlerts = 0
        return cls._word_instance

    @classmethod
    def release_word_instance(cls):
        """释放Word实例（批量处理完成后调用；Linux 下为空操作）。"""
        if cls._word_instance is not None:
            try:
                cls._word_instance.Quit()
            except Exception:
                pass
            cls._word_instance = None
        if cls._com_initialized:
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass
            cls._com_initialized = False

    @classmethod
    def extract_table_data(cls, doc_path, reuse_word=True):
        """
        从DOC文件提取表格数据（与PDF提取相同的字段）。
        Args:
            doc_path: DOC文件路径
            reuse_word: 是否复用Word实例（批量处理时设为True以提高性能）
        Returns:
            dict: {'来文字号': ..., '文件标题': ...}，失败返回 None
        """
        doc_path = os.path.abspath(doc_path)

        # 0) 扩展名不可信：先按文件头识别 WordML（Word 2003 XML 单文件）
        if WordMLExtractor.is_wordml(doc_path):
            data = WordMLExtractor.extract_table_data(doc_path)
            if data:
                return data

        # 1) 兼容「.docx 重命名为 .doc」的文件：本质是 OOXML，python-docx 可直接读
        try:
            d = Document(doc_path)
            data = cls._from_document(d)
            if data:
                print(f"DOC提取成功(OXML): 来文字号={data.get('来文字号','')}, 文件标题={data.get('文件标题','')}")
                return data
        except Exception:
            pass

        # 2) Windows：Word COM（保留旧版行为）
        if IS_WINDOWS:
            word = None
            doc = None
            should_quit = False
            try:
                if reuse_word:
                    word = cls.get_word_instance()
                else:
                    import pythoncom
                    import win32com.client
                    pythoncom.CoInitialize()
                    word = win32com.client.Dispatch("Word.Application")
                    word.Visible = False
                    word.DisplayAlerts = 0
                    should_quit = True

                doc = word.Documents.Open(doc_path)
                data = {}
                if doc.Tables.Count > 0:
                    table = doc.Tables(1)
                    try:
                        data['来文字号'] = cls._clean(table.Cell(1, 2).Range.Text)
                    except Exception:
                        data['来文字号'] = ''
                    try:
                        data['文件标题'] = cls._clean(table.Cell(3, 3).Range.Text)
                    except Exception:
                        data['文件标题'] = ''
                # 关键：COM 打开成功但表格为空/全是空字符串时，不能把空 dict
                # 当成「提取成功」返回，否则上层不会继续尝试 WPS / LibreOffice，
                # 现场就会表现为「提取数据失败」。只有拿到实质内容才提前返回。
                if data and (data.get('来文字号') or data.get('文件标题')):
                    print(f"DOC提取成功(COM): 来文字号={data.get('来文字号','')}, 文件标题={data.get('文件标题','')}")
                    return data
                print("DOC(COM) 未取到有效字段，继续尝试 WPS / LibreOffice")
            except Exception as e:
                print(f"DOC文件提取失败 {doc_path}: {e}")
            finally:
                if doc is not None:
                    try:
                        doc.Close(False)
                    except Exception:
                        pass
                if should_quit and word is not None:
                    try:
                        import pythoncom
                        word.Quit()
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass

        # 3) Linux：WPS（统信/麒麟预装）优先，LibreOffice 兜底，转成 .docx 后读取。
        #    pc.doc_to_docx 内部已按 WPS -> LO 的顺序尝试；这里额外暴露 WPS 直读分支，
        #    让「直接用 WPS 提取数据」不依赖 LibreOffice 是否安装。
        tmp_dir = tempfile.mkdtemp(prefix="docext_")
        try:
            # 3a) WPS 专用通道：转换 + 读表格一步完成
            if prefer_wps():
                try:
                    import wps_compat
                    rows = wps_compat.extract_table_with_wps(
                        doc_path,
                        reader=None,
                        out_dir=tmp_dir,
                    )
                    if rows and rows.get('_rows'):
                        data = cls._from_rows(rows['_rows'])
                        if data:
                            print(f"DOC提取成功(WPS): 来文字号={data.get('来文字号','')}, "
                                  f"文件标题={data.get('文件标题','')}")
                            return data
                except Exception as e:
                    print(f"WPS 提取失败 {doc_path}: {e}")

            # 3b) 通用通道（WPS 或 LibreOffice 转换）
            converted = pc.doc_to_docx(doc_path, tmp_dir)
            if converted:
                d = Document(converted)
                data = cls._from_document(d)
                if data:
                    print(f"DOC提取成功(转换): 来文字号={data.get('来文字号','')}, 文件标题={data.get('文件标题','')}")
                    return data
        except Exception as e:
            print(f"DOC文件提取失败 {doc_path}: {e}")
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        return None


class WordExtractor:
    """Word文档数据提取器"""
    
    @staticmethod
    def extract_table_data(docx_path):
        """
        从Word文档提取表格数据
        返回: dict包含提取的数据
        """
        # 扩展名不可信：现场大量「热线工单」导出为 Word 2003 XML（WordML），
        # 文件名却叫 .docx。python-docx 读不了，必须单独解析。
        if WordMLExtractor.is_wordml(docx_path):
            return WordMLExtractor.extract_table_data(docx_path)

        try:
            doc = Document(docx_path)
            
            if len(doc.tables) == 0:
                print(f"警告: {docx_path} 没有找到表格")
                return None
            
            # 取第一个表格
            table = doc.tables[0]
            
            data = {}
            
            # 第3行第2列（"来源"标签右侧的单元格，示例中为"来电"）
            if len(table.rows) > 2 and len(table.rows[2].cells) > 1:
                data['来源'] = table.rows[2].cells[1].text.strip()
            else:
                data['来源'] = ""
            
            # 第4行第2列提取"姓名："和"电话："之间的内容
            # 尝试第4行（索引3）
            data['姓名'] = ""
            for row_idx in [3, 2]:
                if len(table.rows) > row_idx and len(table.rows[row_idx].cells) > 1:
                    cell_text = table.rows[row_idx].cells[1].text.strip()
                    # 提取"姓名："后、"电话："前的内容
                    if "姓名：" in cell_text and "电话：" in cell_text:
                        start_idx = cell_text.find("姓名：") + 3
                        end_idx = cell_text.find("电话：")
                        data['姓名'] = cell_text[start_idx:end_idx].strip()
                        break
                    elif "电话：" in cell_text or "电话:" in cell_text:
                        # 如果只有"电话："，提取其前面的内容
                        separator = "电话：" if "电话：" in cell_text else "电话:"
                        data['姓名'] = cell_text.split(separator)[0].strip()
                        break
            
            print(f"Word提取成功: 来源={data['来源']}, 姓名={data['姓名']}")
            return data
            
        except Exception as e:
            print(f"提取Word数据失败 {docx_path}: {e}")
            return None
    
    @staticmethod
    def extract_complaint_data(docx_path):
        """
        从批办单Word文档中提取完整数据（10个字段）
        专门用于提取生成的批办单中的数据
        返回: dict包含提取的所有批办单字段数据
        """
        try:
            doc = Document(docx_path)
            
            if len(doc.tables) == 0:
                print(f"警告: {docx_path} 没有找到表格")
                return None
            
            # 取第一个表格
            table = doc.tables[0]
            
            data = {}
            
            # 批办单表格位置映射（行索引，列索引）
            field_mapping = {
                '来文单位': (0, 2),
                '来文字号': (1, 2),
                '收文日期': (1, 9),
                '来文类型': (2, 2),
                '收文途径': (2, 9),
                '收文编号': (3, 9),
                '文件标题': (4, 2),
                '领导批示': (5, 2),
                '承办意见': (6, 2),      # 拟办意见
                '办理情况': (9, 2),
            }
            
            # 提取所有字段
            for field_name, (row, col) in field_mapping.items():
                try:
                    if row < len(table.rows) and col < len(table.rows[row].cells):
                        cell_text = table.rows[row].cells[col].text.strip()
                        data[field_name] = cell_text
                    else:
                        data[field_name] = ''
                except Exception as e:
                    print(f"提取字段 {field_name} 失败: {e}")
                    data[field_name] = ''
            
            print(f"批办单提取成功: 收文编号={data.get('收文编号', '')}, 来文单位={data.get('来文单位', '')}, 文件标题={data.get('文件标题', '')}")
            return data
            
        except Exception as e:
            print(f"提取批办单数据失败 {docx_path}: {e}")
            return None



_DOC_LINE_PITCH = [312]     # 文档网格行距（twips），默认 A4 常见值


def _detect_doc_line_pitch(template_path):
    """
    从模板里读 docGrid 的 linePitch（twips），存入 _DOC_LINE_PITCH。

    为什么需要它：模板启用了「行网格」（w:docGrid w:type="lines"），
    Word 会把行高对齐到网格。LibreOffice 也支持网格，但如果段落行距被
    我们写成 auto，两边算法就分叉了。给「模板没显式行距」的段落补上
    以 linePitch 为准的固定行距，能让两边结果一致。
    """
    try:
        import re
        import zipfile
        with zipfile.ZipFile(template_path) as z:
            xml = z.read('word/document.xml').decode('utf-8', 'ignore')
        m = re.search(r'<w:docGrid[^/]*w:linePitch="(\d+)"', xml)
        if m:
            v = int(m.group(1))
            if v > 0:
                _DOC_LINE_PITCH[0] = v
                return v
    except Exception:
        pass
    return _DOC_LINE_PITCH[0]


# ---------------------------------------------------------------------------
# docGrid（行网格）处理 —— 解决「Word 一页、转 PDF 两页」
# ---------------------------------------------------------------------------
# 模板的 sectPr 里有 <w:docGrid w:type="lines" w:linePitch="312"/>。
#
# 这是 Word 的「文档网格」：每行对齐到 15.6pt 的网格上。
# Word 自己排版时会严格按网格吸附，我们这个模板就是按网格调出来的，
# 所以在 Word 里正好一页。
#
# LibreOffice 对网格的实现与 Word 不同：它把网格当作「最小行高」，
# 再叠加段落自身的行距，结果每行比 Word 高出零点几磅。模板页边距很紧
# （A4 上下各 1440 twips），十几行累积下来，表格后面的「备注」那行
# 就被挤到第 2 页。
#
# 实测（LibreOffice Kit 转 PDF，同一份内容只差这一行属性）：
#     带 docGrid  : 2 页
#     去掉 docGrid: 1 页      <-- 与 Word 记录的页数一致
# 连空模板本身都是这个结果（app.xml 记录 1 页，LO 转出来 2 页）。
#
# ⚠ 只删 docGrid 不够，会「表格变形」：
#   模板里很多表格行没有 trHeight，行高完全由网格撑起来（Word 下每行
#   约 64px）。一旦去掉网格，这些行就塌到内容的最小高度（约 41px），
#   整张表矮了一截、行高全变 —— 现场反馈「虽然一页，但表格完全变了样」。
#
#   解决办法：删 docGrid 的同时，给「原本没有 trHeight」的行补一个
#   atLeast 行高，值取该行在 Word 下由网格撑出的高度。实测补完后
#   LibreOffice 渲染的行高与 Word 逐行吻合（误差 ≤2px）。
#
# 可用 COMPLAINTS_KEEP_DOCGRID=1 关闭这个行为（万一某个模板确实依赖网格）。
_KEEP_DOCGRID = (os.environ.get('COMPLAINTS_KEEP_DOCGRID', '').strip() == '1')

# 删网格时给「无 trHeight」的行补的最小行高（twips）。
# 取网格行距 linePitch 的 2 倍：模板 linePitch=312，312*2=624，
# 实测 Word 下这类行高约 640~650 twips，取 640 最贴合（误差 <1px）。
_GRID_ROW_HEIGHT_TWIPS = 640
_GRID_ROW_HEIGHT_ALT = 650      # 行内含较大字号时用略高值（避免文字被压）

# 表格外段落的垂直间距补偿（twips）。
# 依据：用 Word（COM）导出基准 PDF，与 LibreOffice 输出逐项比对后标定。
#   Word 基准 : 顶留白 83.5pt / 标题→表格 27.5pt / 表格顶 132.0pt
#   未补偿时  : 顶留白 78.5pt / 标题→表格 19.0pt / 表格顶 118.5pt
# 补偿后与 Word 完全一致：顶留白 83.5 / 标题→表格 27.5 / 表格顶 132.0。
_TITLE_BEFORE_TWIPS = 100       # 标题段前 5pt（补顶部留白差）
_SPACER_LINE_TWIPS = 430        # 标题与表格之间空行的行高（补间距差）


def _compensate_outer_spacing(xml):
    """
    给表格**外**的段落补垂直间距，让 LibreOffice 渲染与 Word 一致。

    为什么只动表格外：表格内的行高由 trHeight 控制（已另行补齐）；
    表格外的标题和空行原本靠 docGrid 网格撑高，去掉网格后
    LibreOffice 会把它们压扁 —— 表现为「表格整体上移、顶部留白变大、
    标题与表格几乎贴在一起」。用 Word 导出的 PDF 做基准标定后，
    补上固定的段前距与空行行高即可精确对齐。

    返回 (新xml, 调整的段落数)。
    """
    import re

    a = xml.find('<w:tbl>')
    if a < 0:
        return xml, 0
    b = xml.find('</w:tbl>')
    if b < 0:
        return xml, 0
    b += len('</w:tbl>')
    pre, mid, post = xml[:a], xml[a:b], xml[b:]

    def _fix_first_paragraph(seg):
        """标题段（表格前第一个段落）补段前距。"""
        def f(m):
            p = m.group(0)
            if '<w:spacing' in p and 'w:before' in p:
                return p
            sp = '<w:spacing w:before="%d"/>' % _TITLE_BEFORE_TWIPS
            if '<w:spacing' in p:
                return re.sub(r'<w:spacing\b', '<w:spacing w:before="%d"' %
                              _TITLE_BEFORE_TWIPS, p, count=1)
            if '<w:pPr>' in p:
                return p.replace('<w:pPr>', '<w:pPr>' + sp, 1)
            return p.replace('<w:p>', '<w:p><w:pPr>' + sp + '</w:pPr>', 1)
        return re.sub(r'<w:p\b(?:(?!</w:p>).)*</w:p>', f, seg, count=1, flags=re.S)

    def _fix_spacers(seg):
        """表格前的空段落：把行高提到标定值（补偿网格消失的高度）。"""
        changed = [0]

        def f(m):
            p = m.group(0)
            if '<w:t' in p:      # 有文字的段落不动
                return p
            mline = re.search(r'<w:spacing[^>]*w:line="(\d+)"[^>]*/>', p)
            if mline:
                cur = int(mline.group(1))
                if cur >= _SPACER_LINE_TWIPS:
                    return p
                newp = re.sub(r'(<w:spacing[^>]*w:line=")\d+(")',
                              r'\g<1>%d\2' % _SPACER_LINE_TWIPS, p, count=1)
                # 行距规则必须是 exact/atLeast，否则按字号算
                if 'w:lineRule' not in newp:
                    newp = newp.replace('w:line="%d"' % _SPACER_LINE_TWIPS,
                                        'w:line="%d" w:lineRule="exact"' % _SPACER_LINE_TWIPS, 1)
                changed[0] += 1
                return newp
            sp = '<w:spacing w:line="%d" w:lineRule="exact"/>' % _SPACER_LINE_TWIPS
            if '<w:pPr>' in p:
                changed[0] += 1
                return p.replace('<w:pPr>', '<w:pPr>' + sp, 1)
            changed[0] += 1
            return p.replace('<w:p>', '<w:p><w:pPr>' + sp + '</w:pPr>', 1)

        return re.sub(r'<w:p\b(?:(?!</w:p>).)*</w:p>', f, seg, flags=re.S), changed[0]

    # 只处理标题段 + 表格前的空段落
    pre2 = _fix_first_paragraph(pre)
    pre2, n2 = _fix_spacers(pre2)
    return pre2 + mid + post, n2


def _strip_docgrid_in_xml(xml, fix_row_heights=True):
    """
    删掉 document.xml 里 sectPr 的 <w:docGrid .../>，
    并按需给「原本没有 trHeight」的表格行补最小行高。

    只动 docGrid 与缺失的 trHeight，其它 sectPr 属性（纸张、页边距、
    页眉页脚距离）以及已有的 trHeight 原样保留。

    返回 (新xml, 删除的docGrid个数, 补充的trHeight个数)。
    """
    import re

    pattern = re.compile(r'<w:docGrid\b[^>]*/>')
    grid_count = len(pattern.findall(xml))
    if grid_count:
        xml = pattern.sub('', xml)

    added = 0
    if fix_row_heights and grid_count:
        # 逐行处理：没有 trHeight 的行补上 atLeast 行高
        def _fix_row(m):
            nonlocal added
            row = m.group(0)
            # 该行是否含较大字号（sz >= 32 半磅 = 16pt 以上）→ 用稍高值
            szs = [int(x) for x in re.findall(r'<w:sz w:val="(\d+)"/>', row)]
            h = _GRID_ROW_HEIGHT_ALT if (szs and max(szs) >= 32) else _GRID_ROW_HEIGHT_TWIPS

            exist = re.search(r'<w:trHeight w:val="(\d+)"', row)
            if exist:
                # 已有 trHeight 但值很小（如 90 twips = 4.5pt）时，
                # Word 是靠网格把这一行撑到 64px 的；去掉网格后会塌掉。
                # 这里把它抬到网格行高，保证行高与 Word 一致。
                if int(exist.group(1)) < _GRID_ROW_HEIGHT_TWIPS:
                    row = re.sub(r'<w:trHeight[^/]*/>',
                                 '<w:trHeight w:val="%d" w:hRule="atLeast"/>' % h,
                                 row, count=1)
                    added += 1
                return row

            th = '<w:trHeight w:val="%d" w:hRule="atLeast"/>' % h
            if '<w:trPr>' in row:
                row = row.replace('<w:trPr>', '<w:trPr>' + th, 1)
            else:
                # 在 trPr 缺失时插入（紧跟 <w:tr ...> 开标签之后）
                m2 = re.match(r'(<w:tr\b[^>]*>)', row)
                if m2:
                    row = m2.group(1) + '<w:trPr>' + th + '</w:trPr>' + row[m2.end():]
            added += 1
            return row

        xml = re.sub(r'<w:tr\b.*?</w:tr>', _fix_row, xml, flags=re.S)

        # 表格外的垂直间距补偿（标题段前距 + 空段落行高）。
        # 否则 LibreOffice 会把标题和空行压扁，导致表格整体上移、
        # 顶部留白偏大、标题与表格几乎贴在一起（与 Word 不符）。
        xml, _n_sp = _compensate_outer_spacing(xml)
        added += _n_sp

    return xml, grid_count, added


def strip_docgrid(docx_path):
    """
    移除 docx 的 docGrid（行网格），并保持表格行高不变。

    为什么保存后再改：python-docx 没有暴露 docGrid，直接操作它的
    oxml 也能删，但保存后重新打包 document.xml 更稳（不依赖
    python-docx 内部结构，模板换版本也不会失效）。

    返回 (是否修改, 删除docGrid个数, 补充的trHeight个数)。
    """
    import os
    import shutil
    import tempfile
    import zipfile

    if _KEEP_DOCGRID:
        return False, 0, 0
    if not os.path.isfile(docx_path):
        return False, 0, 0

    tmp_fd, tmp_path = tempfile.mkstemp(suffix='.docx',
                                        dir=os.path.dirname(os.path.abspath(docx_path)))
    os.close(tmp_fd)
    removed = 0
    added = 0
    try:
        with zipfile.ZipFile(docx_path, 'r') as zin, \
                zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'word/document.xml':
                    text = data.decode('utf-8', 'ignore')
                    text, removed, added = _strip_docgrid_in_xml(text)
                    data = text.encode('utf-8')
                zout.writestr(item, data)
        if removed:
            shutil.move(tmp_path, docx_path)
            # 打印后清理
            tmp_path = None
            return True, removed, added
    except Exception:
        pass
    finally:
        if tmp_path and os.path.isfile(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass
    return False, 0, 0


class WordGenerator:
    """Word文档生成器"""
    def __init__(self, template_path='投诉批办单示例#1.docx', output_dir=None):
        # 模板优先从可写目录读取（用户可自行替换），否则从打包资源目录读取
        if not os.path.isabs(template_path):
            writable = os.path.join(get_app_dir(), template_path)
            self.template_path = writable if os.path.isfile(writable) else get_resource_path(template_path)
        else:
            self.template_path = template_path
        if output_dir is None:
            try:
                with open(get_appdata_path('path_config.json'), 'r', encoding='utf-8-sig') as f:
                    path_config = json.load(f)
                    output_dir = path_config.get('待打印', '待打印')
            except Exception:
                output_dir = '待打印'
        self.output_dir = output_dir
        os.makedirs(self.output_dir, exist_ok=True)
        # 读取模板的网格行距（docGrid linePitch），供固定行距兜底使用
        _detect_doc_line_pitch(self.template_path)
    
    @staticmethod
    def _doc_line_pitch():
        """
        读取当前文档的网格行距（docGrid linePitch，单位 twips）。
        用于给「模板里没有显式固定行距」的段落补一个固定行距，
        让 LibreOffice 的排版结果与 Word 的「按网格对齐」接近。
        """
        try:
            from docx.oxml.ns import qn
            # body 的 sectPr -> docGrid
            pass
        except Exception:
            pass
        return _DOC_LINE_PITCH[0]

    @staticmethod
    def set_cell_text_keep_format(cell, text, alignment=None, add_prefix_space=False, line_spacing_pt=None, font_size_pt=14):
        """
        设置单元格文本，统一使用仿宋GB2312字体
        同时设置中文字体和西文字体
        支持多行文本（使用多个段落，避免软回车问题）
        
        参数:
            cell: 单元格对象
            text: 要填充的文本
            alignment: 段落对齐方式（None=默认, 'center'=居中, 'left'=左对齐, 'right'=右对齐）
            add_prefix_space: 是否在每个段落前添加两个全角空格（2个汉字宽度）（通过首行缩进实现）
            line_spacing_pt: 固定行距（磅值），None则使用单倍行距
            font_size_pt: 字体大小（磅值），默认14pt（四号）
        """
        from docx.shared import Pt
        from docx.oxml.ns import qn
        from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING

        # 关键：不要用 cell.text = ''。
        # 那样会把模板段落里已有的 <w:spacing w:line="..." w:lineRule="exact"/>
        # 一起清掉，之后新建段落只能退化成单倍行距（w:line="240" w:lineRule="auto"）。
        # Word 对 auto 行距比较宽容，LibreOffice 会按实际行高排 —— 于是同一份
        # 内容在麒麟上被撑高、最后一行溢出到第 2 页（现场「一页变两页」）。
        # 正确做法：保留第一个段落元素，只清掉它里面的 run，段落属性原样保留。
        _keep_spacing = True

        def _strip_runs(paragraph):
            for r in list(paragraph.runs):
                r._element.getparent().remove(r._element)

        first_para = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
        # 记录模板自带的固定行距，多行时后续段落沿用
        _tpl_line = None
        _tpl_rule = None
        try:
            _pPr = first_para._p.pPr
            if _pPr is not None and _pPr.spacing_line is not None:
                _tpl_line = _pPr.spacing_line
                _tpl_rule = _pPr.spacing_lineRule
        except Exception:
            pass

        # 删掉除首个之外的段落（多段落残留会让单元格变高）
        for p in list(cell.paragraphs[1:]):
            p._element.getparent().remove(p._element)
        _strip_runs(first_para)

        # 调用方可能传 None（提取器缺字段），统一当空串处理，避免
        # 'NoneType' object has no attribute 'replace' 让整份文档生成失败
        if text is None:
            text = ''
        elif not isinstance(text, str):
            text = str(text)

        # 将文本按换行分割，并过滤掉空行（兼容全角空格/不间断空格）
        lines = text.replace('\r\n', '\n').replace('\r', '\n').split('\n')
        def _normalize_blank(s: str) -> str:
            return s.replace('\u3000', '').replace('\xa0', '').strip()
        lines = [ln for ln in lines if _normalize_blank(ln)]
        
        if not lines:
            return
        
        # 公共的段落格式设置函数
        def _format_paragraph(paragraph):
            pf = paragraph.paragraph_format
            # 注意：**不要**无条件写 space_before / space_after = 0。
            # 模板里那三个带行距的段落只有 <w:spacing w:line=".." w:lineRule="exact"/>，
            # 并没有 before/after。我们若额外写 before="0" after="0"，
            # 生成的 XML 就与模板不一致；更关键的是，一旦某个属性被显式写出，
            # 就会覆盖掉样式/默认值，行高计算随之偏移，表格就被撑开了。
            # 只在调用方明确要求时（line_spacing_pt）才设置整段行距。
            if line_spacing_pt is not None:
                pf.space_before = Pt(0)
                pf.space_after = Pt(0)
                pf.line_spacing = Pt(line_spacing_pt)
                # 明确指定行为：固定行距（兼容Word 2007）
                pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            elif _tpl_line:
                # 沿用模板自带的固定行距。
                # 注意单位：python-docx 的 spacing_line 返回的是 **EMU**
                # （1 pt = 12700 EMU），不是 twips，也不是磅。
                # 直接 Pt(spacing_line) 会得到天文数字（如 215900 -> 215900pt），
                # LibreOffice 读到这种行距就彻底撑爆页面。
                try:
                    _pt = _tpl_line / 12700.0          # EMU -> pt
                    if _pt <= 0 or _pt > 200:          # 异常值兜底
                        _pt = _DOC_LINE_PITCH[0] / 20.0
                    pf.line_spacing = Pt(_pt)
                    # _tpl_rule 是枚举，str() 是 'EXACTLY (4)'，不能直接比名字
                    try:
                        _is_exact = (_tpl_rule == WD_LINE_SPACING.EXACTLY)
                    except Exception:
                        _is_exact = ('EXACT' in str(_tpl_rule).upper())
                    pf.line_spacing_rule = (WD_LINE_SPACING.EXACTLY if _is_exact
                                            else WD_LINE_SPACING.AT_LEAST)
                except Exception:
                    pass
            else:
                # 模板该段落本来就没有设置行距 —— 那就**什么都不要设**。
                #
                # 这里曾经写过「用文档网格 linePitch 补一个固定行距」，
                # 结果把 40 个原本靠网格自动计算的段落全部钉死成 15.6pt
                # 固定行高，表格被整体撑开、行高变大 —— 现场反馈
                # 「虽然一页，但表格完全变了样」正是这个原因。
                #
                # 正确做法：模板没设就不设。行高交给 Word/LibreOffice 的
                # 默认规则（配合去掉 docGrid），两边排出来才一致。
                pass
            if alignment == 'center':
                pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif alignment == 'left':
                pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
            elif alignment == 'right':
                pf.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            # 使用首行缩进实现“空两字”
            if add_prefix_space:
                pf.first_line_indent = Pt(28)
        
        # 处理首段：复用已存在（已清空 run）的段落，保留其段落属性
        _format_paragraph(first_para)
        # 去掉前导的半角/全角/不间断空格
        # 写入 docx 的字体名必须是模板要求的「仿宋_GB2312」，和 Windows 版一致。
        # 注意不能用 ui_font()：它会把 Windows 字体映射成 Noto Serif CJK SC，
        # 导致麒麟上生成的批办单字体与模板/Windows 版不一致。
        run = first_para.add_run(lines[0].lstrip(' \t\u3000\u00A0'))
        run.font.size = Pt(font_size_pt)
        run.font.name = '仿宋_GB2312'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋_GB2312')
        run._element.rPr.rFonts.set(qn('w:ascii'), '仿宋_GB2312')
        run._element.rPr.rFonts.set(qn('w:hAnsi'), '仿宋_GB2312')
        
        # 处理其余段落
        for line in lines[1:]:
            para = cell.add_paragraph()
            _format_paragraph(para)
            r = para.add_run(line.lstrip(' \t\u3000\u00A0'))
            r.font.size = Pt(font_size_pt)
            r.font.name = '仿宋_GB2312'
            r._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋_GB2312')
            r._element.rPr.rFonts.set(qn('w:ascii'), '仿宋_GB2312')
            r._element.rPr.rFonts.set(qn('w:hAnsi'), '仿宋_GB2312')
    
    @staticmethod
    def set_cell_text_with_custom_format(cell, text, font_size_pt=14, line_spacing_pt=None, alignment=None):
        """
        设置单元格文本，允许自定义字号和行距
        专用于合并批办单时的来文字号格式设置
        
        参数:
            cell: 单元格对象
            text: 要填充的文本
            font_size_pt: 字体大小（磅值），默认14pt（四号）
            line_spacing_pt: 固定行距（磅值），None则使用单倍行距
            alignment: 段落对齐方式（None=默认, 'center'=居中, 'left'=左对齐, 'right'=右对齐）
        """
        from docx.shared import Pt
        from docx.oxml.ns import qn
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        
        # 清除单元格中所有段落
        for paragraph in cell.paragraphs:
            p = paragraph._element
            p.getparent().remove(p)
        
        # 将文本按 \r\n 分割成多行
        lines = text.replace('\r\n', '\n').replace('\r', '\n').split('\n')
        
        # 为每一行创建一个段落
        for i, line in enumerate(lines):
            paragraph = cell.add_paragraph()
            
            # 设置段落样式：紧凑行距，无段前段后间距
            paragraph_format = paragraph.paragraph_format
            paragraph_format.space_before = Pt(0)
            paragraph_format.space_after = Pt(0)
            
            # 设置行距
            if line_spacing_pt is not None:
                # 使用固定行距（磅值）
                paragraph_format.line_spacing = Pt(line_spacing_pt)
            else:
                # 使用单倍行距
                paragraph_format.line_spacing = 1.0
            
            # 设置对齐方式
            if alignment == 'center':
                paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif alignment == 'left':
                paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            elif alignment == 'right':
                paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            
            # 字体名固定为模板要求的「仿宋_GB2312」，不能用 ui_font() 做
            # Linux 映射，否则写进 docx 的会是 Noto Serif CJK SC。
            run = paragraph.add_run(line)
            # 设置字体大小（使用传入的参数）
            run.font.size = Pt(font_size_pt)
            # 设置西文字体（通过font.name）
            run.font.name = '仿宋_GB2312'
            # 设置完整的字体属性（确保中英文、数字、符号都使用仿宋）
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋_GB2312')
            run._element.rPr.rFonts.set(qn('w:ascii'), '仿宋_GB2312')
            run._element.rPr.rFonts.set(qn('w:hAnsi'), '仿宋_GB2312')
    
    @staticmethod
    def add_qrcode_to_document(doc, qr_data, size='small'):
        """
        在Word文档右下角添加二维码（悬浮于文字之上）
        如果已存在二维码，先删除旧的再添加新的
        
        参数:
            doc: Document 对象
            qr_data: 二维码内容（字典格式）
            size: 'large' (2.7cm) 或 'small' (1.7cm)
        """
        try:
            import qrcode
            from io import BytesIO
            from docx.shared import Inches
            from docx.oxml import parse_xml
            from docx.oxml.ns import qn
            
            # 构建二维码文本内容
            qr_text_parts = []
            for field, value in qr_data.items():
                if value:  # 只添加非空值
                    qr_text_parts.append(f"{field}：{value}")
            
            if not qr_text_parts:
                return  # 没有内容，不生成二维码
            
            qr_text = '\n'.join(qr_text_parts)
            
            # 生成二维码图片
            qr = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_L,
                box_size=10,
                border=1,
            )
            qr.add_data(qr_text)
            qr.make(fit=True)
            
            img = qr.make_image(fill_color="black", back_color="white")
            
            # 将图片保存到内存
            img_buffer = BytesIO()
            img.save(img_buffer, format='PNG')
            img_buffer.seek(0)
            
            # 确定二维码尺寸
            if size == 'large':
                qr_size_inch = Inches(2.7 / 2.54)
                qr_size_emu = 972000  # 2.7cm in EMU
                # 右下角位置 - 大尺寸（相对页面定位）
                pos_h = 5760000  # 16cm from left (原15cm + 1cm)
                pos_v = 9200000  # 25.5cm from top
            else:  # small
                qr_size_inch = Inches(1.7 / 2.54)
                qr_size_emu = 612000  # 1.7cm in EMU
                # 右下角位置 - 小尺寸（相对页面定位）
                pos_h = 6360000  # 17.7cm from left (原16.7cm + 1cm)
                pos_v = 9600000  # 26.7cm from top
            
            # 简化策略：总是在表格前插入一个新的空段落用于二维码
            # 1. 获取body并找到表格位置
            body = doc._element.body
            table_index = None
            for i, element in enumerate(body):
                if element.tag.endswith('}tbl'):
                    table_index = i
                    break
            
            if table_index is None:
                print("错误: 未找到表格")
                return
            
            # 2. 在表格前插入空段落
            para_xml = '''<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:pPr>
    <w:spacing w:before="0" w:after="0" w:line="1" w:lineRule="exact"/>
  </w:pPr>
</w:p>'''
            new_para_element = parse_xml(para_xml)
            body.insert(table_index, new_para_element)
            
            # 3. 找到新插入的段落对象
            # 新段落在body[table_index]位置
            # 计算它是第几个段落
            para_index = 0
            for i in range(table_index + 1):  # 只遍历到table_index
                if body[i].tag.endswith('}p'):
                    if i == table_index:
                        target_para = doc.paragraphs[para_index]
                        break
                    para_index += 1
            else:
                print("错误: 未找到目标段落")
                return
            
            # 在目标段落添加图片
            run = target_para.add_run()
            inline_shape = run.add_picture(img_buffer, width=qr_size_inch, height=qr_size_inch)
            
            # 获取图片的relationship ID和其他信息
            inline = inline_shape._inline
            blip = inline.graphic.graphicData.pic.blipFill.blip
            rId = blip.get(qn('r:embed'))
            # id和name没有命名空间前缀
            docPr_id = inline.docPr.get('id')
            docPr_name = inline.docPr.get('name')
            
            # 构建anchor XML（浮动定位）
            # behindDoc="0" 表示浮于文字上方
            # wrapNone 表示无文字环绕
            import random
            shape_id = random.randint(10000000, 99999999)
            
            anchor_xml = f'''<wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       distT="0" distB="0" distL="0" distR="0" 
       simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" 
       layoutInCell="0" allowOverlap="1">
  <wp:simplePos x="0" y="0"/>
  <wp:positionH relativeFrom="page">
    <wp:posOffset>{pos_h}</wp:posOffset>
  </wp:positionH>
  <wp:positionV relativeFrom="page">
    <wp:posOffset>{pos_v}</wp:posOffset>
  </wp:positionV>
  <wp:extent cx="{qr_size_emu}" cy="{qr_size_emu}"/>
  <wp:effectExtent l="0" t="0" r="0" b="0"/>
  <wp:wrapNone/>
  <wp:docPr id="{docPr_id}" name="{docPr_name}"/>
  <wp:cNvGraphicFramePr>
    <a:graphicFrameLocks noChangeAspect="1"/>
  </wp:cNvGraphicFramePr>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
      <pic:pic>
        <pic:nvPicPr>
          <pic:cNvPr id="{shape_id}" name="QRCode"/>
          <pic:cNvPicPr/>
        </pic:nvPicPr>
        <pic:blipFill>
          <a:blip r:embed="{rId}"/>
          <a:stretch>
            <a:fillRect/>
          </a:stretch>
        </pic:blipFill>
        <pic:spPr>
          <a:xfrm>
            <a:off x="0" y="0"/>
            <a:ext cx="{qr_size_emu}" cy="{qr_size_emu}"/>
          </a:xfrm>
          <a:prstGeom prst="rect">
            <a:avLst/>
          </a:prstGeom>
        </pic:spPr>
      </pic:pic>
    </a:graphicData>
  </a:graphic>
</wp:anchor>'''
            
            # 用anchor替换inline
            anchor = parse_xml(anchor_xml)
            drawing = inline.getparent()  # w:drawing元素
            # 创建新的w:drawing包含anchor
            new_drawing_xml = f'''<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                     xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                     xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
                     xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
{anchor_xml}
</w:drawing>'''
            new_drawing = parse_xml(new_drawing_xml)
            
            # 替换整个w:drawing元素
            run_element = drawing.getparent()  # w:r元素
            run_element.replace(drawing, new_drawing)
            
            print(f"  已添加二维码到文档 (尺寸: {size}, 浮动定位在页面: h={pos_h/360000:.1f}cm, v={pos_v/360000:.1f}cm)")
            
        except ImportError:
            print("  警告: qrcode 库未安装，跳过二维码生成")
        except Exception as e:
            print(f"  添加二维码失败: {e}")
            import traceback
            traceback.print_exc()
    
    def generate_from_pdf(self, pdf_data, file_number, receipt_number, output_path, source_filename=''):
        """
        从PDF数据生成Word文档
        source_filename: 源文件名，用于关键词匹配
        """
        try:
            # 加载模板
            doc = Document(self.template_path)
            
            if len(doc.tables) == 0:
                print("错误: 模板文件没有表格")
                return False
            
            # 获取第一个表格
            table = doc.tables[0]
            
            # 获取当前日期（格式：YYYY-M-D，去掉月和日的前导零）
            now = datetime.now()
            current_date = f"{now.year}-{now.month}-{now.day}"
            
            # 填充数据到对应标签后的单元格（标签右侧的单元格）
            # 注意：由于表格有单元格合并，实际列索引与视觉位置不同
            # 使用保持格式的方法填充
            
            # 来文单位 → 行[0]列[2]（"来文单位"标签右侧）- 添加前置空格
            self.set_cell_text_keep_format(table.rows[0].cells[2], "12345", add_prefix_space=True)
            
            # 来文字号 → 行[1]列[2]（"来文字号"标签右侧）- 居中，根据字符数调整字号
            lai_wen_zi_hao = pdf_data.get('来文字号', '')
            char_count = len(lai_wen_zi_hao)
            
            # 根据字符数确定字号
            if char_count >= 33:
                font_size = 9  # 小五号
            elif 29 <= char_count <= 32:
                font_size = 10.5  # 五号
            elif 25 <= char_count <= 28:
                font_size = 12  # 小四号
            else:
                font_size = 14  # 四号（默认）
            
            self.set_cell_text_keep_format(table.rows[1].cells[2], lai_wen_zi_hao, alignment='center', font_size_pt=font_size)
            
            # 收文日期 → 行[1]列[9]（"收文日期"标签右侧）- 居中
            self.set_cell_text_keep_format(table.rows[1].cells[9], current_date, alignment='center')
            
            # 来文类型 → 行[2]列[2]（"来文类型"标签右侧）- 居中
            self.set_cell_text_keep_format(table.rows[2].cells[2], "C", alignment='center')
            
            # 收文途径 → 行[2]列[9]（"收文途径"标签右侧）- 居中
            self.set_cell_text_keep_format(table.rows[2].cells[9], "信息中心", alignment='center')
            
            # 收文编号 → 行[3]列[9]（"收文编号"标签右侧）- 居中，格式：C-年份后两位-编号
            year_suffix = str(now.year)[-2:]  # 获取年份后两位
            self.set_cell_text_keep_format(table.rows[3].cells[9], f"C-{year_suffix}-{receipt_number}", alignment='center')
            
            # 文件标题 → 行[4]列[2]（"文件标题"标签右侧）- 添加前置空格
            self.set_cell_text_keep_format(table.rows[4].cells[2], pdf_data.get('文件标题', ''), add_prefix_space=True)
            
            # 根据文件名和文件标题关键词填充额外内容
            # 使用新的规则引擎
            keyword_content = None
            try:
                from rule_engine import RuleEngine
                engine = RuleEngine()
                result = engine.match_filling_rule(source_filename, pdf_data.get('文件标题', ''))
                if result:
                    keyword_content = {
                        '拟办意见': result.get('拟办意见', ''),
                        '办理结果': result.get('办理结果', '')
                    }
                    print(f"  [规则引擎] 已自动填充拟办意见和办理结果")
            except Exception as e:
                print(f"  规则引擎失败: {e}")
                import traceback
                traceback.print_exc()

            if keyword_content:
                # 拟办意见 → 行[6]列[2]（每个段落添加前置空格，行距设为28磅，兼容Word 2007）
                self.set_cell_text_keep_format(table.rows[6].cells[2], keyword_content['拟办意见'], add_prefix_space=True, line_spacing_pt=28)
                # 办理结果 → 行[9]列[2]（添加前置空格）
                self.set_cell_text_keep_format(table.rows[9].cells[2], keyword_content['办理结果'], add_prefix_space=True)
            
            # 保存文档
            doc.save(output_path)
            # 去掉 docGrid（行网格）：LibreOffice 按网格叠加行距，
            # 会比 Word 高出零点几磅，导致表格后的「备注」被挤到第 2 页。
            # 各段落行距已显式固定，去掉网格不影响版式。
            strip_docgrid(output_path)
            print(f"生成Word文档成功: {output_path}")
            return True
            
        except Exception as e:
            print(f"生成Word文档失败: {e}")
            return False
    
    def generate_from_doc(self, doc_data, file_number, receipt_number, output_path, source_filename=''):
        """
        从DOC文件数据生成Word文档（与PDF提取逻辑相同，但支持文件名关键词匹配）
        source_filename: 源文件名，用于关键词匹配
        """
        try:
            # 加载模板
            doc = Document(self.template_path)
            
            if len(doc.tables) == 0:
                print("错误: 模板文件没有表格")
                return False
            
            # 获取第一个表格
            table = doc.tables[0]
            
            # 获取当前日期（格式：YYYY-M-D，去掉月和日的前导零）
            now = datetime.now()
            current_date = f"{now.year}-{now.month}-{now.day}"
            
            # 获取文件标题，检查是否需要添加"（重办）"
            title = doc_data.get('文件标题', '')
            if '重办' in source_filename and title:
                if '（重办）' not in title:
                    title += '（重办）'
            
            # 填充基本数据（与PDF相同）
            self.set_cell_text_keep_format(table.rows[0].cells[2], "12345", add_prefix_space=True)
            
            # 来文字号 - 根据字符数调整字号
            lai_wen_zi_hao = doc_data.get('来文字号', '')
            char_count = len(lai_wen_zi_hao)
            
            # 根据字符数确定字号
            if char_count >= 33:
                font_size = 9  # 小五号
            elif 29 <= char_count <= 32:
                font_size = 10.5  # 五号
            elif 25 <= char_count <= 28:
                font_size = 12  # 小四号
            else:
                font_size = 14  # 四号（默认）
            
            self.set_cell_text_keep_format(table.rows[1].cells[2], lai_wen_zi_hao, alignment='center', font_size_pt=font_size)
            self.set_cell_text_keep_format(table.rows[1].cells[9], current_date, alignment='center')
            self.set_cell_text_keep_format(table.rows[2].cells[2], "C", alignment='center')
            self.set_cell_text_keep_format(table.rows[2].cells[9], "信息中心", alignment='center')
            # 收文编号格式：C-年份后两位-编号
            year_suffix = str(now.year)[-2:]  # 获取年份后两位
            self.set_cell_text_keep_format(table.rows[3].cells[9], f"C-{year_suffix}-{receipt_number}", alignment='center')
            self.set_cell_text_keep_format(table.rows[4].cells[2], title, add_prefix_space=True)
            
            # 根据文件名和文件标题关键词填充额外内容
            # 使用新的规则引擎
            keyword_content = None
            try:
                from rule_engine import RuleEngine
                engine = RuleEngine()
                result = engine.match_filling_rule(source_filename, title)
                if result:
                    keyword_content = {
                        '拟办意见': result.get('拟办意见', ''),
                        '办理结果': result.get('办理结果', '')
                    }
                    print(f"  [规则引擎] 已自动填充拟办意见和办理结果")
            except Exception as e:
                print(f"  规则引擎失败: {e}")
                import traceback
                traceback.print_exc()

            if keyword_content:
                # 拟办意见 → 行[6]列[2]（每个段落添加前置空格，行距设为28磅，兼容Word 2007）
                self.set_cell_text_keep_format(table.rows[6].cells[2], keyword_content['拟办意见'], add_prefix_space=True, line_spacing_pt=28)
                # 办理结果 → 行[9]列[2]（添加前置空格）
                self.set_cell_text_keep_format(table.rows[9].cells[2], keyword_content['办理结果'], add_prefix_space=True)
            
            # 保存文档
            doc.save(output_path)
            # 去掉 docGrid（行网格）：LibreOffice 按网格叠加行距，
            # 会比 Word 高出零点几磅，导致表格后的「备注」被挤到第 2 页。
            # 各段落行距已显式固定，去掉网格不影响版式。
            strip_docgrid(output_path)
            print(f"生成Word文档成功: {output_path}")
            return True
            
        except Exception as e:
            print(f"生成Word文档失败: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def generate_from_word(self, word_data, file_number, receipt_number, output_path, source_filename=''):
        """
        从Word数据生成Word文档
        source_filename: 源文件名，用于关键词匹配
        """
        try:
            # 加载模板
            doc = Document(self.template_path)
            
            if len(doc.tables) == 0:
                print("错误: 模板文件没有表格")
                return False
            
            # 获取第一个表格
            table = doc.tables[0]
            
            # 获取当前日期（格式：YYYY-M-D，去掉月和日的前导零）
            now = datetime.now()
            current_date = f"{now.year}-{now.month}-{now.day}"
            
            # 填充数据到对应标签后的单元格（标签右侧的单元格）
            # 注意：由于表格有单元格合并，实际列索引与视觉位置不同
            # 使用保持格式的方法填充
            
            # 三种源文件（PDF / .docx / .doc）走的是同一套工单格式，
            # 都有「工单编号」在表格第一行第二列，批次填充规则完全一致：
            # 来文单位=12345、收文途径=信息中心、来文字号=工单编号。
            # （历史遗留：本函数早期按「旧 docx 才填姓名/来源」写过，
            #   现场已确认不存在那种旧格式，故与 pdf/doc 分支统一。）
            self.set_cell_text_keep_format(table.rows[0].cells[2], "12345", add_prefix_space=True)

            # 收文日期 → 行[1]列[9]（"收文日期"标签右侧）- 居中
            self.set_cell_text_keep_format(table.rows[1].cells[9], current_date, alignment='center')

            # 来文字号 → 行[1]列[2]（"来文字号"标签右侧）- 居中，按字符数调字号
            # 值来自工单表格第一行第二列（WordMLExtractor 提取为「工单编号」，
            # 统一映射到「来文字号」）。
            lai_wen_zi_hao = word_data.get('来文字号') or word_data.get('工单编号', '')
            char_count = len(lai_wen_zi_hao)
            if char_count >= 33:
                font_size = 9
            elif 29 <= char_count <= 32:
                font_size = 10.5
            elif 25 <= char_count <= 28:
                font_size = 12
            else:
                font_size = 14
            self.set_cell_text_keep_format(table.rows[1].cells[2], lai_wen_zi_hao,
                                           alignment='center', font_size_pt=font_size)

            # 来文类型 → 行[2]列[2]（"来文类型"标签右侧）- 居中
            self.set_cell_text_keep_format(table.rows[2].cells[2], "C", alignment='center')

            # 收文途径 → 行[2]列[9]（"收文途径"标签右侧）- 居中
            self.set_cell_text_keep_format(table.rows[2].cells[9], "信息中心", alignment='center')
            
            # 收文编号 → 行[3]列[9]（"收文编号"标签右侧）- 居中，格式：C-年份后两位-编号
            year_suffix = str(now.year)[-2:]  # 获取年份后两位
            self.set_cell_text_keep_format(table.rows[3].cells[9], f"C-{year_suffix}-{receipt_number}", alignment='center')

            # 文件标题 → 行[4]列[2]（"文件标题"标签右侧）- 添加前置空格
            # 兼容两种提取器：WordML/热线工单用「标题」，旧版 docx 表用「文件标题」
            word_title = word_data.get('标题') or word_data.get('文件标题', '')
            self.set_cell_text_keep_format(table.rows[4].cells[2], word_title, add_prefix_space=True)

            # 根据文件名和文件标题关键词填充额外内容
            # 优先使用新的规则引擎
            keyword_content = None
            try:
                from rule_engine import RuleEngine
                engine = RuleEngine()
                result = engine.match_filling_rule(source_filename, word_data.get('标题', ''))
                if result:
                    keyword_content = {
                        '拟办意见': result.get('拟办意见', ''),
                        '办理结果': result.get('办理结果', '')
                    }
                    print(f"  [规则引擎] 已自动填充拟办意见和办理结果")
            except Exception as e:
                print(f"  规则引擎失败: {e}")
                import traceback
                traceback.print_exc()

            if keyword_content:
                # 拟办意见 → 行[6]列[2]（每个段落添加前置空格，行距设为28磅，兼容Word 2007）
                self.set_cell_text_keep_format(table.rows[6].cells[2], keyword_content['拟办意见'], add_prefix_space=True, line_spacing_pt=28)
                # 办理结果 → 行[9]列[2]（添加前置空格）
                self.set_cell_text_keep_format(table.rows[9].cells[2], keyword_content['办理结果'], add_prefix_space=True)

            # 保存文档
            doc.save(output_path)
            # 去掉 docGrid（行网格）：LibreOffice 按网格叠加行距，
            # 会比 Word 高出零点几磅，导致表格后的「备注」被挤到第 2 页。
            # 各段落行距已显式固定，去掉网格不影响版式。
            strip_docgrid(output_path)
            print(f"生成Word文档成功: {output_path}")
            return True
            
        except Exception as e:
            print(f"生成Word文档失败: {e}")
            return False


class DocumentProcessor:
    """文档处理主程序"""
    
    def __init__(self, output_dir=None):
        self.counter = CounterManager()
        if output_dir is None:
            try:
                with open(get_appdata_path('path_config.json'), 'r', encoding='utf-8-sig') as f:
                    path_config = json.load(f)
                    output_dir = path_config.get('待打印', '待打印')
            except Exception:
                output_dir = '待打印'
        self.output_dir = output_dir
        self.generator = WordGenerator(output_dir=output_dir)
        self.template_file = os.path.basename(self.generator.template_path)
        
        # 确保输出目录存在
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)
            print(f"创建输出目录: {self.output_dir}")
    
    def scan_files(self):
        """
        扫描当前目录的PDF、DOCX和DOC文件
        返回: (pdf_files, docx_files, doc_files)
        """
        current_dir = os.getcwd()
        
        # 查找所有PDF文件
        pdf_files = glob.glob(os.path.join(current_dir, '*.pdf'))
        
        # 查找所有DOCX文件
        docx_files = glob.glob(os.path.join(current_dir, '*.docx'))
        docx_files = [f for f in docx_files 
                     if not os.path.basename(f).startswith('~$')
                     and os.path.basename(f) != self.template_file
                     and not os.path.basename(f).startswith('投诉批办单#')]
        
        # 查找所有DOC文件
        doc_files = glob.glob(os.path.join(current_dir, '*.doc'))
        doc_files = [f for f in doc_files 
                    if not os.path.basename(f).startswith('~$')
                    and not os.path.basename(f).startswith('投诉批办单#')]
        
        print(f"\n扫描到 {len(pdf_files)} 个PDF文件, {len(docx_files)} 个DOCX文件, {len(doc_files)} 个DOC文件")
        
        return pdf_files, docx_files, doc_files
    
    def process_pdf(self, pdf_path):
        """处理单个PDF文件"""
        print(f"\n处理PDF: {os.path.basename(pdf_path)}")

        # 提取数据
        data = PDFExtractor.extract_table_data(pdf_path)
        if data is None:
            return False

        # 获取编号
        file_num = self.counter.get_next_file_number()
        receipt_num = self.counter.get_next_receipt_number()

        # 生成输出文件名
        output_filename = f"投诉批办单#{file_num}.docx"
        output_path = os.path.join(self.output_dir, output_filename)

        # 生成Word文档（传入源文件名用于关键词匹配）
        filename = os.path.basename(pdf_path)
        success = self.generator.generate_from_pdf(data, file_num, receipt_num, output_path, filename)

        return success
    
    def process_word(self, word_path):
        """处理单个DOCX文件"""
        print(f"\n处理DOCX: {os.path.basename(word_path)}")

        # 提取数据
        data = WordExtractor.extract_table_data(word_path)
        if data is None:
            return False

        # 获取编号
        file_num = self.counter.get_next_file_number()
        receipt_num = self.counter.get_next_receipt_number()

        # 生成输出文件名
        output_filename = f"投诉批办单#{file_num}.docx"
        output_path = os.path.join(self.output_dir, output_filename)

        # 生成Word文档（传入源文件名用于关键词匹配）
        filename = os.path.basename(word_path)
        success = self.generator.generate_from_word(data, file_num, receipt_num, output_path, filename)

        return success
    
    def process_doc(self, doc_path):
        """处理单个DOC文件（旧版Word格式）"""
        print(f"\n处理DOC: {os.path.basename(doc_path)}")
        
        # 提取数据
        data = DOCExtractor.extract_table_data(doc_path)
        if data is None:
            return False
        
        # 获取编号
        file_num = self.counter.get_next_file_number()
        receipt_num = self.counter.get_next_receipt_number()
        
        # 生成输出文件名
        output_filename = f"投诉批办单#{file_num}.docx"
        output_path = os.path.join(self.output_dir, output_filename)
        
        # 生成Word文档（传入源文件名用于关键词匹配）
        filename = os.path.basename(doc_path)
        success = self.generator.generate_from_doc(data, file_num, receipt_num, output_path, filename)
        
        return success
    
    def run(self):
        """运行主程序"""
        print("=" * 50)
        print("投诉批办单自动生成应用")
        print("=" * 50)
        
        # 扫描文件
        pdf_files, docx_files, doc_files = self.scan_files()
        
        # 处理统计
        total = len(pdf_files) + len(docx_files) + len(doc_files)
        success_count = 0
        fail_count = 0
        
        # 处理所有PDF文件
        for pdf_file in pdf_files:
            if self.process_pdf(pdf_file):
                success_count += 1
            else:
                fail_count += 1
        
        # 处理所有DOCX文件
        for docx_file in docx_files:
            if self.process_word(docx_file):
                success_count += 1
            else:
                fail_count += 1
        
        # 处理所有DOC文件
        for doc_file in doc_files:
            if self.process_doc(doc_file):
                success_count += 1
            else:
                fail_count += 1
        
        # 输出处理结果
        print("\n" + "=" * 50)
        print(f"处理完成!")
        print(f"总文件数: {total}")
        print(f"成功: {success_count}")
        print(f"失败: {fail_count}")
        print(f"输出目录: {os.path.abspath(self.output_dir)}")
        print("=" * 50)


def main():
    """程序入口（命令行模式，跨平台启动引导）"""
    pc.ensure_app_cwd()
    pc.seed_user_data()
    processor = DocumentProcessor()
    processor.run()


if __name__ == '__main__':
    main()

