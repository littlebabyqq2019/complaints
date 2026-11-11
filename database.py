# -*- coding: utf-8 -*-
"""
批办单数据库管理模块
使用SQLite存储批办单记录，支持检索和恢复
"""

import sqlite3
import os
from datetime import datetime
from typing import List, Dict, Optional, Tuple
from platform_compat import get_appdata_path


class ComplaintDatabase:
    """批办单数据库管理类"""
    
    def __init__(self, db_path=None):
        """
        初始化数据库连接
        
        Args:
            db_path: 数据库文件路径，默认使用AppData/complaints.db
        """
        if db_path is None:
            db_path = get_appdata_path('complaints.db')
        self.db_path = db_path
        self.init_database()
        self.optimize_database()
    
    def init_database(self):
        """初始化数据库表结构"""
        try:
            print(f"正在初始化数据库: {self.db_path}")
            
            # 确保数据库目录存在
            db_dir = os.path.dirname(self.db_path)
            if db_dir and not os.path.exists(db_dir):
                os.makedirs(db_dir)
                print(f"已创建数据库目录: {db_dir}")
            
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 检查表是否存在，如果存在则检查结构
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='complaints'")
            table_exists = cursor.fetchone() is not None
            
            if table_exists:
                # 检查主键是否是收文编号
                cursor.execute("PRAGMA table_info(complaints)")
                columns = cursor.fetchall()
                # columns格式: (cid, name, type, notnull, dflt_value, pk)
                primary_key_col = None
                for col in columns:
                    if col[5] == 1:  # pk=1表示主键
                        primary_key_col = col[1]
                        break
                
                if primary_key_col != '收文编号':
                    print("检测到旧数据库结构，开始迁移...")
                    # 备份旧数据
                    cursor.execute('SELECT * FROM complaints')
                    old_data = cursor.fetchall()
                    cursor.execute('PRAGMA table_info(complaints)')
                    old_columns = [col[1] for col in cursor.fetchall()]
                    
                    # 删除旧表
                    cursor.execute('DROP TABLE complaints')
                    print("已删除旧表结构")
                    
                    # 创建新表（收文编号作为主键）
                    cursor.execute('''
                        CREATE TABLE complaints (
                            收文编号 TEXT PRIMARY KEY,
                            来文单位 TEXT,
                            来文字号 TEXT,
                            收文日期 TEXT,
                            来文类型 TEXT,
                            收文途径 TEXT,
                            文件标题 TEXT,
                            领导批示 TEXT,
                            承办意见 TEXT,
                            办理情况 TEXT,
                            文件名 TEXT,
                            创建时间 TEXT,
                            更新时间 TEXT
                        )
                    ''')
                    print("已创建新表结构（收文编号作为主键）")
                    
                    # 迁移数据
                    if old_data:
                        # 找到收文编号列的索引
                        receipt_col_idx = old_columns.index('收文编号') if '收文编号' in old_columns else -1
                        migrated = 0
                        for row in old_data:
                            try:
                                # 构建新数据行
                                new_row = {}
                                for i, col_name in enumerate(old_columns):
                                    if col_name in ['收文编号', '来文单位', '来文字号', '收文日期', '来文类型', 
                                                   '收文途径', '文件标题', '领导批示', '承办意见', '办理情况', 
                                                   '文件名', '创建时间', '更新时间']:
                                        new_row[col_name] = row[i]
                                
                                # 插入新表
                                cursor.execute('''
                                    INSERT INTO complaints (收文编号, 来文单位, 来文字号, 收文日期, 来文类型, 
                                                          收文途径, 文件标题, 领导批示, 承办意见, 办理情况, 
                                                          文件名, 创建时间, 更新时间)
                                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                                ''', (
                                    new_row.get('收文编号', ''),
                                    new_row.get('来文单位', ''),
                                    new_row.get('来文字号', ''),
                                    new_row.get('收文日期', ''),
                                    new_row.get('来文类型', ''),
                                    new_row.get('收文途径', ''),
                                    new_row.get('文件标题', ''),
                                    new_row.get('领导批示', ''),
                                    new_row.get('承办意见', ''),
                                    new_row.get('办理情况', ''),
                                    new_row.get('文件名', ''),
                                    new_row.get('创建时间', ''),
                                    new_row.get('更新时间', '')
                                ))
                                migrated += 1
                            except Exception as e:
                                print(f"迁移记录失败: {e}")
                        print(f"数据迁移完成，成功迁移 {migrated}/{len(old_data)} 条记录")
                else:
                    print("数据库结构已是最新版本（收文编号作为主键）")
            else:
                # 创建批办单表（收文编号作为主键）
                cursor.execute('''
                    CREATE TABLE complaints (
                        收文编号 TEXT PRIMARY KEY,
                        来文单位 TEXT,
                        来文字号 TEXT,
                        收文日期 TEXT,
                        来文类型 TEXT,
                        收文途径 TEXT,
                        文件标题 TEXT,
                        领导批示 TEXT,
                        承办意见 TEXT,
                        办理情况 TEXT,
                        文件名 TEXT,
                        创建时间 TEXT,
                        更新时间 TEXT
                    )
                ''')
                print("数据库表创建成功（收文编号作为主键）")
            
            # 创建索引以加速检索（收文编号已是主键，无需额外索引）
            cursor.execute('''
                CREATE INDEX IF NOT EXISTS idx_date ON complaints(收文日期)
            ''')
            cursor.execute('''
                CREATE INDEX IF NOT EXISTS idx_unit ON complaints(来文单位)
            ''')
            cursor.execute('''
                CREATE INDEX IF NOT EXISTS idx_title ON complaints(文件标题)
            ''')
            print("数据库索引创建成功")
            
            conn.commit()
            conn.close()
            print(f"数据库初始化完成: {self.db_path}")
            
        except Exception as e:
            print(f"数据库初始化失败: {e}")
            import traceback
            traceback.print_exc()
            raise
    
    def add_complaint(self, data: Dict[str, str]) -> bool:
        """
        添加批办单记录
        
        Args:
            data: 批办单数据字典，包含所有字段
            
        Returns:
            bool: 是否添加成功
        """
        try:
            # 检查数据库文件是否存在，如果不存在则重新初始化
            if not os.path.exists(self.db_path):
                print(f"警告: 数据库文件不存在，正在重新创建: {self.db_path}")
                self.init_database()
            
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 检查表是否存在，如果不存在则重新初始化
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='complaints'")
            if not cursor.fetchone():
                conn.close()
                print("警告: complaints表不存在，正在重新初始化数据库")
                self.init_database()
                conn = sqlite3.connect(self.db_path)
                cursor = conn.cursor()
            
            now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            cursor.execute('''
                INSERT OR REPLACE INTO complaints (
                    来文单位, 来文字号, 收文日期, 来文类型, 收文途径, 收文编号,
                    文件标题, 领导批示, 承办意见, 办理情况, 文件名, 创建时间, 更新时间
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                data.get('来文单位', ''),
                data.get('来文字号', ''),
                data.get('收文日期', ''),
                data.get('来文类型', ''),
                data.get('收文途径', ''),
                data.get('收文编号', ''),
                data.get('文件标题', ''),
                data.get('领导批示', ''),
                data.get('承办意见', ''),
                data.get('办理情况', ''),
                data.get('文件名', ''),
                now,
                now
            ))
            
            conn.commit()
            conn.close()
            print(f"成功添加记录到数据库: {data.get('收文编号', 'N/A')}")
            return True
            
        except Exception as e:
            print(f"添加批办单记录失败: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def update_complaint(self, receipt_number: str, data: Dict[str, str]) -> bool:
        """
        更新批办单记录
        
        Args:
            receipt_number: 收文编号
            data: 要更新的数据字典
            
        Returns:
            bool: 是否更新成功
        """
        try:
            # 检查数据库文件是否存在
            if not os.path.exists(self.db_path):
                print(f"警告: 数据库文件不存在，无法更新记录: {self.db_path}")
                return False
            
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            cursor.execute('''
                UPDATE complaints SET
                    来文单位 = ?, 来文字号 = ?, 收文日期 = ?, 来文类型 = ?, 收文途径 = ?,
                    文件标题 = ?, 领导批示 = ?, 承办意见 = ?, 办理情况 = ?, 文件名 = ?,
                    更新时间 = ?
                WHERE 收文编号 = ?
            ''', (
                data.get('来文单位', ''),
                data.get('来文字号', ''),
                data.get('收文日期', ''),
                data.get('来文类型', ''),
                data.get('收文途径', ''),
                data.get('文件标题', ''),
                data.get('领导批示', ''),
                data.get('承办意见', ''),
                data.get('办理情况', ''),
                data.get('文件名', ''),
                now,
                receipt_number
            ))
            
            # 检查是否成功更新了记录
            rows_affected = cursor.rowcount
            conn.commit()
            conn.close()
            
            if rows_affected > 0:
                print(f"成功更新数据库记录: {receipt_number} (影响行数: {rows_affected})")
                return True
            else:
                print(f"警告: 未找到收文编号为 {receipt_number} 的记录，更新失败")
                return False
            
        except Exception as e:
            print(f"更新批办单记录失败: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def search_complaints(
        self,
        keyword: str = '',
        field: str = 'all',
        start_date: str = '',
        end_date: str = '',
        limit: int = 9999
    ) -> List[Dict[str, str]]:
        """
        搜索批办单记录
        
        Args:
            keyword: 搜索关键词
            field: 搜索字段 ('all', '来文单位', '文件标题', '收文编号' 等)
            start_date: 开始日期
            end_date: 结束日期
            limit: 返回记录数限制（默认1000，支持大数据量）
            
        Returns:
            List[Dict]: 批办单记录列表
        """
        try:
            # 检查数据库文件是否存在
            if not os.path.exists(self.db_path):
                print(f"警告: 数据库文件不存在，返回空结果: {self.db_path}")
                return []
            
            conn = sqlite3.connect(self.db_path)
            conn.row_factory = sqlite3.Row  # 使结果可以按列名访问
            cursor = conn.cursor()
            
            # 构建查询条件
            conditions = []
            params = []
            
            if keyword:
                if field == 'all':
                    # 全文搜索
                    conditions.append('''(
                        来文单位 LIKE ? OR 文件标题 LIKE ? OR 收文编号 LIKE ? OR
                        来文字号 LIKE ? OR 收文途径 LIKE ? OR 领导批示 LIKE ? OR
                        承办意见 LIKE ? OR 办理情况 LIKE ?
                    )''')
                    params.extend([f'%{keyword}%'] * 8)
                else:
                    # 指定字段搜索
                    conditions.append(f'{field} LIKE ?')
                    params.append(f'%{keyword}%')
            
            if start_date:
                conditions.append('收文日期 >= ?')
                params.append(start_date)
            
            if end_date:
                conditions.append('收文日期 <= ?')
                params.append(end_date)
            
            # 构建SQL查询
            sql = 'SELECT * FROM complaints'
            if conditions:
                sql += ' WHERE ' + ' AND '.join(conditions)
            # 排序：兼容新格式 C-YY-N 和旧格式 C-N
            # 先按年份降序（新格式有年份，旧格式视为0），再按末尾编号降序
            sql += '''
                ORDER BY
                    CASE
                        WHEN 收文编号 LIKE 'C-__-%' THEN CAST(SUBSTR(收文编号, 4, 2) AS INTEGER)
                        ELSE 0
                    END DESC,
                    CAST(
                        CASE
                            WHEN 收文编号 LIKE 'C-__-%' THEN SUBSTR(收文编号, 7)
                            ELSE SUBSTR(收文编号, 3)
                        END
                    AS INTEGER) DESC
                LIMIT ?'''
            params.append(limit)
            
            cursor.execute(sql, params)
            rows = cursor.fetchall()
            
            # 转换为字典列表
            results = []
            for row in rows:
                results.append(dict(row))
            
            conn.close()
            return results
            
        except Exception as e:
            print(f"搜索批办单记录失败: {e}")
            return []
    
    def get_complaint_by_receipt_number(self, receipt_number: str) -> Optional[Dict[str, str]]:
        """
        根据收文编号获取批办单记录
        
        Args:
            receipt_number: 收文编号
            
        Returns:
            Dict: 批办单记录，如果未找到返回None
        """
        try:
            conn = sqlite3.connect(self.db_path)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            
            cursor.execute('SELECT * FROM complaints WHERE 收文编号 = ?', (receipt_number,))
            row = cursor.fetchone()
            
            conn.close()
            
            if row:
                return dict(row)
            return None
            
        except Exception as e:
            print(f"获取批办单记录失败: {e}")
            return None
    
    def delete_complaint(self, receipt_number: str) -> bool:
        """
        删除批办单记录
        
        Args:
            receipt_number: 收文编号
            
        Returns:
            bool: 是否删除成功
        """
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('DELETE FROM complaints WHERE 收文编号 = ?', (receipt_number,))
            
            conn.commit()
            conn.close()
            return True
            
        except Exception as e:
            print(f"删除批办单记录失败: {e}")
            return False
    
    def get_statistics(self) -> Dict[str, int]:
        """
        获取统计信息
        
        Returns:
            Dict: 统计信息字典
        """
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            stats = {}
            
            # 总记录数
            cursor.execute('SELECT COUNT(*) FROM complaints')
            stats['总记录数'] = cursor.fetchone()[0]
            
            # 今年记录数
            current_year = datetime.now().strftime('%Y')
            cursor.execute('SELECT COUNT(*) FROM complaints WHERE 收文日期 LIKE ?', (f'{current_year}%',))
            stats['今年记录数'] = cursor.fetchone()[0]
            
            # 本月记录数
            current_month = datetime.now().strftime('%Y-%m')
            cursor.execute('SELECT COUNT(*) FROM complaints WHERE 收文日期 LIKE ?', (f'{current_month}%',))
            stats['本月记录数'] = cursor.fetchone()[0]
            
            conn.close()
            return stats
            
        except Exception as e:
            print(f"获取统计信息失败: {e}")
            return {}
    
    def optimize_database(self):
        """
        优化数据库性能
        设置SQLite性能参数，支持大数据量操作
        """
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 使用WAL模式提高并发性能
            cursor.execute('PRAGMA journal_mode=WAL')
            
            # 增加缓存大小（10MB）
            cursor.execute('PRAGMA cache_size=-10000')
            
            # 使用内存临时存储
            cursor.execute('PRAGMA temp_store=MEMORY')
            
            # 启用外键约束
            cursor.execute('PRAGMA foreign_keys=ON')
            
            # 优化分析数据库
            cursor.execute('ANALYZE')
            
            conn.close()
            
        except Exception as e:
            print(f"数据库优化失败: {e}")
