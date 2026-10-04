#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
VBA Macro Reader - 读取 .xlsm/.xlam 文件中的 VBA 宏代码

支持两种模式：
1. win32com 模式（可读可写，需要 Excel）
2. oletools 模式（只读，不需要 Excel）

作者：QClaw Agent
日期：2026-06-03
"""

import os
import re
import json
from typing import Dict, List, Optional, Tuple
from pathlib import Path


class VBAReader:
    """VBA 宏代码读取器"""

    def __init__(self, filepath: str, use_win32com: bool = True):
        """
        初始化 VBA 读取器

        Args:
            filepath: .xlsm 文件路径
            use_win32com: True 使用 win32com（可读可写）
                         False 使用 oletools（只读）
        """
        self.filepath = Path(filepath)
        self.use_win32com = use_win32com
        self.excel = None
        self.workbook = None
        self.vbaparser = None
        self._file_snapshot = set()  # 用于 get_generated_files 的快照对比

        if not self.filepath.exists():
            raise FileNotFoundError(f"文件不存在: {filepath}")

        if use_win32com:
            self._init_win32com()
        else:
            self._init_oletools()

    def _take_snapshot(self) -> set:
        """拍摄文件所在目录的快照，返回文件路径集合"""
        import os
        dir_path = os.path.dirname(str(self.filepath))
        if not dir_path:
            dir_path = "."
        return set(
            os.path.join(dir_path, f)
            for f in os.listdir(dir_path)
            if os.path.isfile(os.path.join(dir_path, f))
        )

    def _init_win32com(self):
        """初始化 win32com"""
        try:
            import win32com.client

            self.excel = win32com.client.Dispatch("Excel.Application")
            self.excel.Visible = False
            self.excel.DisplayAlerts = False

            self.workbook = self.excel.Workbooks.Open(str(self.filepath.absolute()))

        except ImportError:
            raise ImportError("未安装 pywin32，请运行: pip install pywin32")
        except Exception as e:
            raise RuntimeError(f"无法打开文件: {e}")

    def _init_oletools(self):
        """初始化 oletools"""
        try:
            from oletools.olevba import VBA_Parser

            self.vbaparser = VBA_Parser(str(self.filepath.absolute()))

        except ImportError:
            raise ImportError("未安装 oletools，请运行: pip install oletools")
        except Exception as e:
            raise RuntimeError(f"无法解析文件: {e}")

    def get_all_modules(self) -> Dict[str, str]:
        """
        返回所有 VBA 模块代码

        Returns:
            Dict[模块名, 代码内容]
        """
        if self.use_win32com:
            return self._get_modules_win32com()
        else:
            return self._get_modules_oletools()

    def _get_modules_win32com(self) -> Dict[str, str]:
        """使用 win32com 读取所有模块"""
        modules = {}

        try:
            vb_project = self.workbook.VBProject

            for comp in vb_project.VBComponents:
                comp_type = comp.Type

                # 1=标准模块, 2=类模块, 3=窗体
                if comp_type in [1, 2, 3]:
                    code_module = comp.CodeModule
                    line_count = code_module.CountOfLines

                    if line_count > 0:
                        code = code_module.Lines(1, line_count)
                        modules[comp.Name] = code

        except Exception as e:
            raise RuntimeError(f"读取 VBA 模块失败: {e}")

        return modules

    def _get_modules_oletools(self) -> Dict[str, str]:
        """使用 oletools 读取所有模块"""
        modules = {}

        try:
            for (filename, stream_path, vba_filename, vba_code) in self.vbaparser.extract_macros():
                if vba_filename and vba_code:
                    # 清理模块名
                    module_name = vba_filename.replace(".bas", "").replace(".cls", "")
                    modules[module_name] = vba_code

        except Exception as e:
            raise RuntimeError(f"解析 VBA 失败: {e}")

        return modules

    def get_module(self, module_name: str) -> Optional[str]:
        """
        返回指定模块代码

        Args:
            module_name: 模块名称

        Returns:
            模块代码，如果不存在返回 None
        """
        modules = self.get_all_modules()
        return modules.get(module_name)

    def list_modules(self) -> List[str]:
        """
        列出所有模块名称

        Returns:
            模块名称列表
        """
        return list(self.get_all_modules().keys())

    def list_procedures(self) -> List[Dict]:
        """
        列出所有 Sub/Function 及参数

        Returns:
            过程列表: [{"type": "Sub|Function", "name": "...", "params": [...]}]
        """
        procedures = []
        modules = self.get_all_modules()

        for module_name, code in modules.items():
            # 提取 Sub
            sub_pattern = r'Sub\s+(\w+)\s*\(([^)]*)\)'
            for match in re.finditer(sub_pattern, code, re.IGNORECASE):
                params_str = match.group(2).strip()
                params = [p.strip() for p in params_str.split(',') if p.strip()]

                procedures.append({
                    "module": module_name,
                    "type": "Sub",
                    "name": match.group(1),
                    "params": params
                })

            # 提取 Function
            func_pattern = r'Function\s+(\w+)\s*\(([^)]*)\)'
            for match in re.finditer(func_pattern, code, re.IGNORECASE):
                params_str = match.group(2).strip()
                params = [p.strip() for p in params_str.split(',') if p.strip()]

                procedures.append({
                    "module": module_name,
                    "type": "Function",
                    "name": match.group(1),
                    "params": params
                })

        return procedures

    def extract_parameters(self, module_name: str, procedure_name: str) -> List[str]:
        """
        提取指定过程的参数列表

        Args:
            module_name: 模块名称
            procedure_name: 过程名称

        Returns:
            参数列表
        """
        procedures = self.list_procedures()

        for proc in procedures:
            if proc["module"] == module_name and proc["name"] == procedure_name:
                return proc["params"]

        return []

    def update_module(self, module_name: str, code: str) -> bool:
        """
        更新模块代码（仅 win32com）

        Args:
            module_name: 模块名称
            code: 新代码

        Returns:
            是否成功
        """
        if not self.use_win32com:
            raise RuntimeError("oletools 模式不支持写入操作")

        try:
            vb_project = self.workbook.VBProject

            for comp in vb_project.VBComponents:
                if comp.Name == module_name:
                    code_module = comp.CodeModule
                    line_count = code_module.CountOfLines

                    # 删除旧代码
                    if line_count > 0:
                        code_module.DeleteLines(1, line_count)

                    # 写入新代码
                    code_module.AddFromString(code)
                    return True

            # 如果模块不存在，创建新模块
            return self.add_module(module_name, code)

        except Exception as e:
            raise RuntimeError(f"更新模块失败: {e}")

    def add_module(self, module_name: str, code: str) -> bool:
        """
        新增模增模块（仅 win32com）

        Args:
            module_name: 模块名称
            code: 代码

        Returns:
            是否成功
        """
        if not self.use_win32com:
            raise RuntimeError("oletools 模式不支持写入操作")

        try:
            import win32com.client

            vb_project = self.workbook.VBProject

            # 创建新模块
            new_comp = vb_project.VBComponents.Add(1)  # 1=标准模块
            new_comp.Name = module_name
            new_comp.CodeModule.AddFromString(code)

            return True

        except Exception as e:
            raise RuntimeError(f"新增模块失败: {e}")

    def delete_module(self, module_name: str) -> bool:
        """
        删除模块（仅 win32com）

        Args:
            module_name: 模块名称

        Returns:
            是否成功
        """
        if not self.use_win32com:
            raise RuntimeError("oletools 模式不支持写入操作")

        try:
            vb_project = self.workbook.VBProject

            for comp in vb_project.VBComponents:
                if comp.Name == module_name:
                    vb_project.VBComponents.Remove(comp)
                    return True

            return False

        except Exception as e:
            raise RuntimeError(f"删除模块失败: {e}")

    def analyze_code(self, module_name: str) -> Dict:
        """
        分析代码质量

        Args:
            module_name: 模块名称

        Returns:
            分析结果
        """
        code = self.get_module(module_name)
        if not code:
            return {}

        lines = code.split('\n')
        code_lines = [l.strip() for l in lines if l.strip() and not l.strip().startswith("'")]

        # 圈复杂度：每个决策点 +1（排除注释行中的关键字）
        decision_keywords = [
            'If ', 'ElseIf ', 'For ', 'For Each ',
            'While ', 'Do While', 'Do Until',
            'Select Case',
        ]
        cyclomatic = 1  # 基础复杂度
        for line in code_lines:
            for kw in decision_keywords:
                if kw in line:
                    cyclomatic += 1
            # And/Or 也算分支（排除 Dim 宣告行）
            if not line.startswith('Dim ') and not line.startswith('Public '):
                if ' And ' in line:
                    cyclomatic += 1
                if ' Or ' in line:
                    cyclomatic += 1

        # 嵌套深度
        max_depth = 0
        current_depth = 0
        depth_openers = ['If ', 'For ', 'While ', 'Do ', 'Select Case', 'With ']
        depth_closers = ['End If', 'Next', 'Wend', 'Loop', 'End Select', 'End With']
        for line in code_lines:
            for opener in depth_openers:
                if opener in line:
                    current_depth += 1
                    break
            for closer in depth_closers:
                if closer in line:
                    current_depth = max(0, current_depth - 1)
                    break
            max_depth = max(max_depth, current_depth)

        # 建议
        suggestions = []

        if cyclomatic > 15:
            suggestions.append(f"圈复杂度 {cyclomatic} 过高，建议拆分成多个小函数")
        elif cyclomatic > 10:
            suggestions.append(f"圈复杂度 {cyclomatic} 偏高，考虑简化条件逻辑")

        if max_depth > 4:
            suggestions.append(f"嵌套深度 {max_depth} 过深，建议提取子过程或使用提前返回")

        if 'On Error Resume Next' in code:
            suggestions.append("使用 On Error Resume Next 可能隐藏错误，建议使用明确的错误处理")

        if 'Select *' in code or 'select *' in code.lower():
            suggestions.append("避免使用 SELECT *，建议明确指定字段")

        # 检查是否有注释
        comment_lines = [l for l in lines if l.strip().startswith("'")]
        if len(comment_lines) < len(code_lines) * 0.1:
            suggestions.append("注释不足，建议增加代码注释")

        return {
            "line_count": len(lines),
            "code_lines": len(code_lines),
            "cyclomatic_complexity": cyclomatic,
            "max_nesting_depth": max_depth,
            "has_error_handling": 'On Error' in code,
            "suggestions": suggestions
        }

    def optimize_module(self, module_name: str) -> str:
        """
        返回优化后的代码（建议性优化）

        Args:
            module_name: 模块名称

        Returns:
            优化后的代码
        """
        code = self.get_module(module_name)
        if not code:
            return ""

        lines = code.split('\n')
        optimized_lines = []

        has_option_explicit = any(
            'Option Explicit' in line for line in lines
        )

        # 若缺少 Option Explicit，在文件顶部插入（VBA 要求在所有过程之前）
        if not has_option_explicit:
            optimized_lines.append("Option Explicit")
            optimized_lines.append("")  # 空行分隔

        for line in lines:
            stripped = line.strip()
            # 对高密度属性访问行添加建议（排除注释、声明行）
            if (not stripped.startswith("'")
                    and not stripped.startswith("Dim ")
                    and not stripped.startswith("Public ")
                    and not stripped.startswith("Private ")
                    and line.count('.') > 3
                    and 'With' not in line):
                optimized_lines.append(
                    f"' OPTIMIZE: 考虑使用 With 语句 — {stripped}"
                )
            optimized_lines.append(line.rstrip())

        return '\n'.join(optimized_lines)

    def save(self) -> bool:
        """
        保存修改（仅 win32com）

        Returns:
            是否成功
        """
        if not self.use_win32com:
            raise RuntimeError("oletools 模式不支持保存操作")

        try:
            self.workbook.Save()
            return True
        except Exception as e:
            raise RuntimeError(f"保存失败: {e}")

    # ==================== ★执行宏功能 ====================

    def run_macro(self, macro_name: str) -> Dict:
        """
        执行宏代码（无参数）

        Args:
            macro_name: 宏名称（如 "FillExcelTemplate"）

        Returns:
            执行结果: {"status": "success/error", "message": "...", "duration": 秒}
        """
        if not self.use_win32com:
            raise RuntimeError("oletools 模式不支持执行宏")

        import time
        start_time = time.time()

        # 拍摄快照以追踪生成的文件
        self._file_snapshot = self._take_snapshot()

        try:
            # 确保 Excel 可见（可选）
            self.excel.Visible = True

            # 执行宏
            self.excel.Run(f"'{self.workbook.Name}'!{macro_name}")

            duration = time.time() - start_time

            return {
                "status": "success",
                "message": f"宏 {macro_name} 执行成功",
                "duration": round(duration, 2)
            }

        except Exception as e:
            duration = time.time() - start_time
            return {
                "status": "error",
                "error": str(e),
                "duration": round(duration, 2)
            }

    def run_macro_with_params(self, macro_name: str, params: List) -> Dict:
        """
        执行宏代码并传递参数

        Args:
            macro_name: 宏名称
            params: 参数列表

        Returns:
            执行结果
        """
        if not self.use_win32com:
            raise RuntimeError("oletools 模式不支持执行宏")

        import time
        start_time = time.time()

        # 拍摄快照以追踪生成的文件
        self._file_snapshot = self._take_snapshot()

        try:
            self.excel.Visible = True

            # 构建参数字符串
            param_str = ", ".join([repr(p) if isinstance(p, str) else str(p) for p in params])

            # 执行带参数的宏
            result = self.excel.Run(f"'{self.workbook.Name}'!{macro_name}", *params)

            duration = time.time() - start_time

            return {
                "status": "success",
                "message": f"宏 {macro_name} 执行成功",
                "result": result,
                "duration": round(duration, 2)
            }

        except Exception as e:
            duration = time.time() - start_time
            return {
                "status": "error",
                "error": str(e),
                "duration": round(duration, 2)
            }

    def run_macro_monitored(self, macro_name: str, timeout_seconds: int = 300, log_file: str = None) -> Dict:
        """
        执行宏并监控执行过程

        Args:
            macro_name: 宏名称
            timeout_seconds: 超时秒数
            log_file: 日志文件路径

        Returns:
            执行结果
        """
        import time
        import threading

        # 拍摄快照以追踪生成的文件
        self._file_snapshot = self._take_snapshot()

        result = {"status": "running", "start_time": time.time()}

        def run_macro_thread():
            import pythoncom
            pythoncom.CoInitialize()  # COM 线程安全初始化
            try:
                self.excel.Visible = True
                self.excel.Run(f"'{self.workbook.Name}'!{macro_name}")
                result["status"] = "success"
                result["end_time"] = time.time()
            except Exception as e:
                result["status"] = "error"
                result["error"] = str(e)
                result["end_time"] = time.time()
            finally:
                pythoncom.CoUninitialize()

        # 启动线程
        thread = threading.Thread(target=run_macro_thread)
        thread.start()

        # 等待完成或超时
        thread.join(timeout=timeout_seconds)

        if thread.is_alive():
            result["status"] = "timeout"
            result["message"] = f"执行超时（>{timeout_seconds}秒）"

        result["duration"] = time.time() - result["start_time"]

        # 写入日志
        if log_file:
            with open(log_file, "a", encoding="utf-8") as f:
                f.write(f"\n[{time.strftime('%Y-%m-%d %H:%M:%S')}] {macro_name}: {result}\n")

        return result

    def run_macros_batch(self, macros: List[Tuple[str, List]]) -> List[Dict]:
        """
        批量执行多个宏

        Args:
            macros: 宏列表 [("宏名", [参数列表]), ...]

        Returns:
            执行结果列表
        """
        results = []

        for macro_name, params in macros:
            if params:
                result = self.run_macro_with_params(macro_name, params)
            else:
                result = self.run_macro(macro_name)

            results.append({
                "macro": macro_name,
                "result": result
            })

        return results

    def smart_run(self, macro_name: str, auto_save: bool = True, auto_backup: bool = True,
                  timeout_seconds: int = 300, retry_on_error: bool = False,
                  max_retries: int = 3, log_to_file: bool = False) -> Dict:
        """
        智能执行宏（自动处理参数、监控、错误处理）

        Args:
            macro_name: 宏名称
            auto_save: 执行后自动保存
            auto_backup: 执行前自动备份
            timeout_seconds: 超时秒数
            retry_on_error: 错误时自动重试
            max_retries: 最大重试次数
            log_to_file: 记录到日志文件

        Returns:
            执行结果
        """
        import time
        import shutil
        from datetime import datetime

        # 执行前备份
        backup_path = None
        if auto_backup:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_path = str(self.filepath).replace(".xlsm", f"_backup_{timestamp}.xlsm")
            try:
                shutil.copy2(self.filepath, backup_path)
            except:
                pass

        # 执行宏
        attempts = 0
        result = None
        log_path = None
        if log_to_file:
            import os
            from datetime import datetime
            log_dir = os.path.dirname(str(self.filepath))
            if not log_dir:
                log_dir = "."
            log_path = os.path.join(
                log_dir,
                f"macro_run_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
            )

        while attempts <= (max_retries if retry_on_error else 0):
            result = self.run_macro_monitored(
                macro_name, timeout_seconds, log_file=log_path
            )

            if result["status"] == "success":
                break

            attempts += 1
            if attempts <= max_retries:
                time.sleep(2)  # 等待 2 秒后重试

        # 执行后保存
        if auto_save and result["status"] == "success":
            try:
                self.save()
            except:
                pass

        # 添加备份信息
        result["backup_path"] = backup_path
        result["attempts"] = attempts

        return result

    def get_generated_files(self) -> List[str]:
        """
        获取宏执行后生成的文件列表（基于快照对比）

        Returns:
            生成的文件路径列表
        """
        current = self._take_snapshot()
        return sorted(current - self._file_snapshot)

    def close(self):
        """关闭文件"""
        try:
            if self.use_win32com:
                if self.workbook:
                    self.workbook.Close(SaveChanges=False)
                if self.excel:
                    self.excel.Quit()
            else:
                if self.vbaparser:
                    self.vbaparser.close()
        except:
            pass

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def __del__(self):
        self.close()


# 命令行接口
if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="VBA 宏代码读取器")
    parser.add_argument("file", help=".xlsm 文件路径")
    parser.add_argument("--module", "-m", help="指定模块名称")
    parser.add_argument("--list", "-l", action="store_true", help="列出所有模块")
    parser.add_argument("--procs", "-p", action="store_true", help="列出所有过程")
    parser.add_argument("--analyze", "-a", action="store_true", help="分析代码质量")
    parser.add_argument("--read-only", action="store_true", help="使用只读模式（不需要 Excel）")
    parser.add_argument("--run", "-r", help="执行宏名称")
    parser.add_argument("--params", nargs="*", help="宏参数列表")

    args = parser.parse_args()

    use_win32com = not args.read_only

    try:
        with VBAReader(args.file, use_win32com=use_win32com) as reader:
            if args.run:
                # 执行宏
                print(f"=== 执行宏: {args.run} ===")
                
                if args.params:
                    result = reader.run_macro_with_params(args.run, args.params)
                else:
                    result = reader.run_macro(args.run)
                
                print(f"\n执行结果:")
                print(f"  状态: {result['status']}")
                print(f"  耗时: {result['duration']} 秒")
                
                if result['status'] == 'error':
                    print(f"  错误: {result.get('error', 'Unknown error')}")
                else:
                    print(f"  消息: {result.get('message', 'Success')}")
                    
                    # 获取生成的文件
                    generated = reader.get_generated_files()
                    if generated:
                        print(f"\n生成的文件:")
                        for f in generated:
                            print(f"  - {f}")
            
            elif args.list:
                print("=== 模块列表 ===")
                for module in reader.list_modules():
                    print(f"  - {module}")

            elif args.procs:
                print("=== 过程列表 ===")
                for proc in reader.list_procedures():
                    print(f"  {proc['type']}: {proc['module']}.{proc['name']}({', '.join(proc['params'])})")

            elif args.analyze:
                module_name = args.module or reader.list_modules()[0]
                print(f"=== 分析: {module_name} ===")
                analysis = reader.analyze_code(module_name)
                print(json.dumps(analysis, indent=2, ensure_ascii=False))

            elif args.module:
                code = reader.get_module(args.module)
                if code:
                    print(f"=== {args.module} ===")
                    print(code)
                else:
                    print(f"模块 {args.module} 不存在")

            else:
                print("=== 所有模块 ===")
                for name, code in reader.get_all_modules().items():
                    print(f"\n{'='*50}")
                    print(f"模块: {name}")
                    print(f"{'='*50}")
                    print(code)

    except Exception as e:
        print(f"错误: {e}")
        exit(1)
