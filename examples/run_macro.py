#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
示例：执行 VBA 宏并生成文件

演示如何使用 VBAReader 执行宏代码，生成文件和处理结果
"""

import sys
import os

# 添加 scripts 目录到路径
sys.path.insert(0, os.path.join(
    os.path.dirname(os.path.abspath(__file__)), '..', 'scripts'
))

from vba_reader import VBAReader


def example_run_macro():
    """示例 1：直接执行宏"""
    print("=" * 60)
    print("示例 1：直接执行宏")
    print("=" * 60)

    filepath = r"./工作目录\数据批量套用表格.xltm"

    try:
        # 初始化（必须使用 win32com 模式）
        reader = VBAReader(filepath, use_win32com=True)

        print(f"\n文件: {filepath}")
        print(f"模式: win32com")

        # 列出可用的宏
        procedures = reader.list_procedures()
        print(f"\n可用的宏:")
        for proc in procedures:
            if proc['module'] not in ['ThisWorkbook', 'Sheet1']:
                print(f"  - {proc['name']}")

        # 执行宏
        print(f"\n正在执行宏: FillExcelTemplate")
        print("注意：会弹出用户界面，请选择行范围和模板文件")

        result = reader.run_macro("FillExcelTemplate")

        print(f"\n执行结果:")
        print(f"  状态: {result['status']}")
        print(f"  耗时: {result['duration']} 秒")

        if result['status'] == 'error':
            print(f"  错误: {result['error']}")
        else:
            print(f"  消息: {result['message']}")

            # 获取生成的文件
            generated_files = reader.get_generated_files()
            print(f"\n生成的文件:")
            for f in generated_files:
                print(f"  - {f}")

        # 关闭
        reader.close()

    except FileNotFoundError:
        print(f"文件不存在: {filepath}")
    except Exception as e:
        print(f"错误: {e}")


def example_smart_run():
    """示例 2：智能执行宏"""
    print("\n" + "=" * 60)
    print("示例 2：智能执行宏（自动备份、超时、重试）")
    print("=" * 60)

    filepath = r"./工作目录\数据批量套用表格.xltm"

    try:
        reader = VBAReader(filepath, use_win32com=True)

        # 智能执行
        result = reader.smart_run(
            macro_name="FillExcelTemplate",
            auto_save=True,
            auto_backup=True,
            timeout_seconds=300,  # 5 分钟超时
            retry_on_error=True,
            max_retries=3,
            log_to_file=True
        )

        print(f"\n执行结果:")
        print(f"  状态: {result['status']}")
        print(f"  耗时: {result['duration']} 秒")
        print(f"  尝试次数: {result['attempts']}")

        if result.get('backup_path'):
            print(f"  备份文件: {result['backup_path']}")

        if result['status'] == 'error':
            print(f"  错误: {result.get('error')}")

        reader.close()

    except Exception as e:
        print(f"错误: {e}")


def example_batch_run():
    """示例 3：批量执行多个宏"""
    print("\n" + "=" * 60)
    print("示例 3：批量执行多个宏")
    print("=" * 60)

    filepath = r"./工作目录\数据批量套用表格.xltm"

    try:
        reader = VBAReader(filepath, use_win32com=True)

        # 定义要执行的宏列表
        macros_to_run = [
            ("FillExcelTemplate", []),  # 无参数
            # 可以添加更多宏
        ]

        print(f"\n将执行 {len(macros_to_run)} 个宏:")

        results = reader.run_macros_batch(macros_to_run)

        print(f"\n执行结果:")
        for i, result in enumerate(results, 1):
            print(f"\n  宏 {i}: {result['macro']}")
            print(f"    状态: {result['result']['status']}")
            print(f"    耗时: {result['result']['duration']} 秒")

        reader.close()

    except Exception as e:
        print(f"错误: {e}")


def example_run_with_params():
    """示例 4：执行带参数的宏"""
    print("\n" + "=" * 60)
    print("示例 4：执行带参数的宏")
    print("=" * 60)

    filepath = r"./工作目录\数据批量套用表格.xltm"

    try:
        reader = VBAReader(filepath, use_win32com=True)

        # 获取工作表对象
        ws = reader.workbook.Sheets(1)

        # 执行带参数的宏
        result = reader.run_macro_with_params(
            "ExportWorksheetToPDF",
            [ws, "测试报告_20260603"]
        )

        print(f"\n执行结果:")
        print(f"  状态: {result['status']}")
        print(f"  耗时: {result['duration']} 秒")

        if result['status'] == 'error':
            print(f"  错误: {result['error']}")

        reader.close()

    except Exception as e:
        print(f"错误: {e}")


if __name__ == "__main__":
    print("""
============================================================
    VBA Macro Runner - 执行宏示例
============================================================

本示例演示如何使用 VBAReader 执行 VBA 宏代码

注意事项：
1. 执行宏需要使用 win32com 模式（use_win32com=True）
2. 执行宏时 Excel 会被打开（Visible=True）
3. 某些宏可能需要用户交互（如选（如选择文件、输入参数）
4. 执行完成后记得调用 reader.close() 关闭 Excel

============================================================
""")

    # 运行示例
    example_run_macro()

    # 其他示例（可选）
    # example_smart_run()
    # example_batch_run()
    # example_run_with_params()

    print("\n" + "=" * 60)
    print("示例完成！")
    print("=" * 60)
