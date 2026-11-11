#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
批量处理示例文件并生成批办单
"""

import os
import shutil
from pathlib import Path

def main():
    print("=" * 70)
    print("批量处理示例文件")
    print("=" * 70)

    # 源文件夹和目标文件夹
    sample_dir = "示例"
    target_dir = "待处理"

    # 检查示例文件夹
    if not os.path.exists(sample_dir):
        print(f"[错误] 找不到示例文件夹: {sample_dir}")
        return

    # 确保目标文件夹存在
    os.makedirs(target_dir, exist_ok=True)

    # 获取示例文件
    files = [f for f in os.listdir(sample_dir) if f.endswith(('.docx', '.doc', '.pdf'))]

    if not files:
        print(f"[错误] 示例文件夹中没有文件")
        return

    print(f"\n找到 {len(files)} 个文件")
    print(f"将复制到: {target_dir}/")
    print()

    # 复制文件到待处理文件夹
    copied_count = 0
    for filename in sorted(files):
        source = os.path.join(sample_dir, filename)
        dest = os.path.join(target_dir, filename)

        try:
            shutil.copy2(source, dest)
            print(f"[OK] 已复制: {filename}")
            copied_count += 1
        except Exception as e:
            print(f"[ERROR] 复制失败 {filename}: {e}")

    print("\n" + "=" * 70)
    print(f"完成！共复制 {copied_count} 个文件到 {target_dir}/ 文件夹")
    print("=" * 70)
    print("\n接下来的操作：")
    print("1. 打开批办单管理系统程序")
    print("2. 点击【批量处理】按钮")
    print("3. 程序会自动处理'待处理'文件夹中的所有文件")
    print("4. 生成的批办单将保存在'待打印'文件夹中")
    print("\n或者你可以手动拖拽文件到程序界面进行处理")

if __name__ == "__main__":
    main()
