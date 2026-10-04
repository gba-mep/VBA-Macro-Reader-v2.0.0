#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试 VBA Reader
"""

import sys
import os
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from scripts.vba_reader import VBAReader


def test_oletools_mode():
    """测试 oletools 只读模式"""
    print("测试 oletools 模式...")

    # 使用测试文件（如果存在）
    test_file = r"./工作目录\你的宏目录\你的工作簿.xlsm"

    if not os.path.exists(test_file):
        print(f"跳过测试：文件不存在 {test_file}")
        return

    try:
        with VBAReader(test_file, use_win32com=False) as reader:
            # 测试列出模块
            modules = reader.list_modules()
            print(f"✓ 找到 {len(modules)} 个模块")

            # 测试读取代码
            for module_name in modules:
                code = reader.get_module(module_name)
                assert code, f"无法读取 {module_name}"
                print(f"✓ 读取成功: {module_name} ({len(code)} 字符)")

            # 测试列出过程
            procedures = reader.list_procedures()
            print(f"✓ 找到 {len(procedures)} 个过程")

            # 测试分析代码
            for module_name in modules:
                analysis = reader.analyze_code(module_name)
                assert 'line_count' in analysis
                print(f"✓ 分析成功: {module_name}")

        print("\n✓ oletools 模式测试通过\n")

    except Exception as e:
        print(f"✗ 测试失败: {e}")
        raise


def test_win32com_mode():
    """测试 win32com 模式"""
    print("测试 win32com 模式...")

    # 使用测试文件
    test_file = r"./工作目录\你的宏目录\你的工作簿.xlsm"

    if not os.path.exists(test_file):
        print(f"跳过测试：文件不存在 {test_file}")
        return

    try:
        with VBAReader(test_file, use_win32com=True) as reader:
            # 测试列出模块
            modules = reader.list_modules()
            print(f"✓ 找到 {len(modules)} 个模块")

            # 测试读取代码
            for module_name in modules:
                code = reader.get_module(module_name)
                assert code, f"无法读取 {module_name}"
                print(f"✓ 读取成功: {module_name}")

        print("\n✓ win32com 模式测试通过\n")

    except ImportError as e:
        print(f"跳过测试：{e}")
    except Exception as e:
        print(f"✗ 测试失败: {e}")
        raise


if __name__ == "__main__":
    print("=" * 60)
    print("VBA Reader 测试")
    print("=" * 60 + "\n")

    test_oletools_mode()
    test_win32com_mode()

    print("=" * 60)
    print("所有测试完成！")
    print("=" * 60)
