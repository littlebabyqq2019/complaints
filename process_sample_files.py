#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
处理示例文件并生成批办单
"""

import os
import sys
from main import DocumentProcessor, get_appdata_path

def main():
    print("=" * 70)
    print("处理示例文件夹中的文件")
    print("=" * 70)

    # 初始化处理器
    processor = DocumentProcessor(
        input_dir="示例",
        output_dir="输出测试",
        year=2026
    )

    # 确保输出目录存在
    os.makedirs("输出测试", exist_ok=True)

    # 获取示例文件
    sample_files = []
    if os.path.exists("示例"):
        files = [f for f in os.listdir("示例") if f.endswith(('.docx', '.doc', '.pdf'))]
        sample_files = [os.path.join("示例", f) for f in sorted(files)]

    if not sample_files:
        print("[错误] 示例文件夹中没有找到文件")
        return

    print(f"\n找到 {len(sample_files)} 个文件\n")

    # 处理每个文件
    for i, filepath in enumerate(sample_files, 1):
        filename = os.path.basename(filepath)
        print(f"\n[{i}/{len(sample_files)}] 处理文件: {filename}")
        print("-" * 70)

        try:
            # 处理文件
            result = processor.process_file(filepath)

            if result['success']:
                output_file = result.get('output_file', '')
                print(f"✓ 处理成功")
                print(f"  输出文件: {output_file}")

                # 显示填充的内容
                if result.get('data'):
                    data = result['data']
                    print(f"\n  文件名: {data.get('文件名', '')}")
                    print(f"  文件标题: {data.get('文件标题', '')}")
                    print(f"  拟办意见: {data.get('承办意见', '')[:60]}...")
                    print(f"  办理结果: {data.get('办理情况', '')[:60]}...")
            else:
                print(f"✗ 处理失败: {result.get('error', '未知错误')}")

        except Exception as e:
            print(f"✗ 异常: {e}")
            import traceback
            traceback.print_exc()

    print("\n" + "=" * 70)
    print("处理完成")
    print("=" * 70)
    print(f"\n生成的批办单保存在: 输出测试/ 文件夹")
    print("\n请查看生成的Word文档，确认拟办意见是否正确。")

if __name__ == "__main__":
    main()
