#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
示例：读取你的宏目录系统的 VBA 代码

演示如何使用 VBAReader 读取 .xlsm 文件中的宏代码
"""

import sys
import os

# 添加 scripts 目录到路径
sys.path.insert(0, os.path.join(
    os.path.dirname(os.path.abspath(__file__)), '..', 'scripts'
))

from vba_reader import VBAReader


def example_read_all_modules():
    """示例 1：读取所有模块"""
    print("=" * 60)
    print("示例 1：读取所有模块")
    print("=" * 60)

    filepath = r"./工作目录\你的宏目录\你的工作簿.xlsm"

    try:
        with VBAReader(filepath, use_win32com=False) as reader:
            modules = reader.get_all_modules()

            print(f"\n找到 {len(modules)} 个模块：")
            for name in modules.keys():
                print(f"  - {name}")

            # 保存到文件
            output_dir = r"./工作目录\你的宏目录\vba_backup"
            os.makedirs(output_dir, exist_ok=True)

            for name, code in modules.items():
                filename = f"{name}.bas"
                filepath = os.path.join(output_dir, filename)

                with open(filepath, "w", encoding="utf-8") as f:
                    f.write(code)

                print(f"已保存: {filename}")

    except FileNotFoundError:
        print(f"文件不存在: {filepath}")
        print("请修改 filepath 为实际路径")


def example_read_specific_module():
    """示例 2：读取特定模块"""
    print("\n" + "=" * 60)
    print("示例 2：读取特定模块（mod_Main）")
    print("=" * 60)

    filepath = r"./工作目录\你的宏目录\你的工作簿.xlsm"

    try:
        with VBAReader(filepath, use_win32com=False) as reader:
            code = reader.get_module("mod_Main")

            if code:
                print(f"\n模块代码（前 500 字符）：")
                print(code[:500])

                print(f"\n完整代码长度: {len(code)} 字符")
            else:
                print("模块不存在")

    except FileNotFoundError:
        print(f"文件不存在: {filepath}")


def example_list_procedures():
    """示例 3：列出所有过程"""
    print("\n" + "=" * 60)
    print("示例 3：列出所有 Sub/Function")
    print("=" * 60)

    filepath = r"./工作目录\你的宏目录\你的工作簿.xlsm"

    try:
        with VBAReader(filepath, use_win32com=False) as reader:
            procedures = reader.list_procedures()

            print(f"\n找到 {len(procedures)} 个过程：\n")

            for proc in procedures[:10]:  # 只显示前 10 个
                params = ", ".join(proc['params']) if proc['params'] else "无参数"
                print(f"{proc['type']:8} {proc['module']}.{proc['name']}({params})")

            if len(procedures) > 10:
                print(f"\n... 还有 {len(procedures) - 10} 个过程")

    except FileNotFoundError:
        print(f"文件不存在: {filepath}")


def example_analyze_code():
    """示例 4：分析代码质量"""
    print("\n" + "=" * 60)
    print("示例 4：分析代码质量")
    print("=" * 60)

    filepath = r"./工作目录\你的宏目录\你的工作簿.xlsm"

    try:
        with VBAReader(filepath, use_win32com=False) as reader:
            modules = reader.list_modules()

            print(f"\n分析 {len(modules)} 个模块：\n")

            for module_name in modules:
                analysis = reader.analyze_code(module_name)

                print(f"模块: {module_name}")
                print(f"  代码行数: {analysis.get('line_count', 0)}")
                print(f"  复杂度: {analysis.get('complexity', 0)}")
                print(f"  圈复杂度: {analysis.get('cyclomatic', 0)}")

                if analysis.get('suggestions'):
                    print(f"  建议:")
                    for suggestion in analysis['suggestions']:
                        print(f"    - {suggestion}")

                print()

    except FileNotFoundError:
        print(f"文件不存在: {filepath}")


def example_extract_parameters():
    """示例 5：提取过程参数"""
    print("\n" + "=" * 60)
    print("示例 5：提取过程参数")
    print("=" * 60)

    filepath = r"./工作目录\你的宏目录\你的工作簿.xlsm"

    try:
        with VBAReader(filepath, use_win32com=False) as reader:
            # 提取 FillTemplate 过程的参数
            params = reader.extract_parameters("mod_Main", "FillTemplate")

            print(f"\nFillTemplate 过程参数：")
            for i, param in enumerate(params, 1):
                print(f"  {i}. {param}")

    except FileNotFoundError:
        print(f"文件不存在: {filepath}")


if __name__ == "__main__":
    # 运行所有示例
    example_read_all_modules()
    example_read_specific_module()
    example_list_procedures()
    example_analyze_code()
    example_extract_parameters()

    print("\n" + "=" * 60)
    print("所有示例完成！")
    print("=" * 60)
