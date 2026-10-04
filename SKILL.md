# VBA Macro Reader & Operator Skill

> 版本：v2.0.0
> 创建时间：2026-06-03
> 更新时间：2026-06-03 12:25
> 用途：读取、修改、**执行** VBA 宏代码，生成文件和处理结果

---

## 📋 功能概述

本技能提供完整的 VBA 宏代码读取、分析、操作和**执行**能力：

| 功能 | 描述 | 状态 |
|:---|:---|:---:|
| **读取 VBA 代码** | 从 .xlsm/.xlam 提取所有模块代码 | ✅ |
| **解析模块结构** | 识别 Sub/Function、参数、变量 | ✅ |
| **提取参数信息** | 分析宏的输入参数和返回值 | ✅ |
| **实时修改** | 通过 win32com 动态修改宏代码 | ✅ |
| **★执行宏代码** | **直接运行宏，生成文件和结果** | ✅ |
| **★监控执行** | **监控执行过程，捕获错误** | ✅ |
| **★结果处理** | **处理生成的文件和输出** | ✅ |
| **优化建议** | 基于最佳实践提供优化建议 | ✅ |
| **代码对比** | 对比不同版本的 VBA 代码差异 | ✅ |

---

## 🔧 技术方案

### 方案 A：win32com 方法（✅ 推荐）

**优点：**
- 可读取和修改 VBA 代码
- 可实时执行宏
- 支持 Excel 所有功能

**缺点：**
- 需要安装 Excel
- Windows only

**代码示例：**
```python
import win32com.client

# 打开 Excel 文件
excel = win32com.client.Dispatch("Excel.Application")
wb = excel.Workbooks.Open(r"D:\path\to\file.xlsm")

# 读取 VBA 代码
for comp in wb.VBProject.VBComponents:
    if comp.Type in [1, 2, 3]:  # 标准模块、类模块、窗体
        code_module = comp.CodeModule
        line_count = code_module.CountOfLines
        code = code_module.Lines(1, line_count)
        print(f"=== {comp.Name} ===")
        print(code)

# 修改 VBA 代码
comp.CodeModule.DeleteLines 1, line_count
comp.CodeModule.AddFromString("Sub Test()\n    MsgBox \"Hello\"\nEnd Sub")

wb.Save()
excel.Quit()
```

---

### 方案 B：oletools 解析方法

**优点：**
- 不需要 Excel
- 跨平台

**缺点：**
- 只能读取，不能修改
- 需要手动解析二进制格式

**代码示例：**
```python
from oletools.olevba import VBA_Parser

# 解析 VBA
vbaparser = VBA_Parser(r"D:\path\to\file.xlsm")

# 提取所有 VBA 代码
for (filename, stream_path, vba_filename, vba_code) in vbaparser.extract_macros():
    print(f"=== {vba_filename} ===")
    print(vba_code)

vbaparser.close()
```

---

## 📋 使用方法

### 1. 读取 VBA 宏代码

```python
from scripts.vba_reader import VBAReader

# 初始化
reader = VBAReader(r"D:\path\to\file.xlsm")

# 读取所有模块
modules = reader.get_all_modules()

for module_name, code in modules.items():
    print(f"\n=== {module_name} ===")
    print(code)

# 读取特定模块
main_code = reader.get_module("mod_Main")

# 列出所有 Sub/Function
procs = reader.list_procedures()
for proc in procs:
    print(f"{proc['type']}: {proc['name']}({proc['params']})")
```

---

### ★2. 执行宏代码（核心功能）

```python
from scripts.vba_reader import VBAReader

# 初始化（必须使用 win32com 模式）
reader = VBAReader(r"D:\path\to\file.xlsm", use_win32com=True)

# 方法 1：直接执行宏（无参数）
result = reader.run_macro("FillExcelTemplate")
print(f"执行结果: {result}")

# 方法 2：执行宏并传递参数
result = reader.run_macro_with_params(
    "ExportWorksheetToPDF",
    [worksheet_object, "报告_20260603"]
)

# 方法 3：执行宏并监控执行过程
result = reader.run_macro_monitored(
    "FillExcelTemplate",
    timeout_seconds=300,  # 5 分钟超时
    log_file="D:\logs\macro_run.log"
)

# 方法 4：批量执行多个宏
macros_to_run = [
    ("PrepareData", []),
    ("FillExcelTemplate", []),
    ("ExportToPDF", []),
]
results = reader.run_macros_batch(macros_to_run)

# 关闭
reader.close()
```

---

### ★3. 处理宏执行结果

```python
# 执行宏并获取生成的文件
result = reader.run_macro("FillExcelTemplate")

# 检查生成的文件
generated_files = reader.get_generated_files()
print(f"生成的文件: {generated_files}")

# 检查执行日志
logs = reader.get_execution_logs()
print(f"执行日志:\n{logs}")

# 检查错误信息
if result['status'] == 'error':
    print(f"错误信息: {result['error']}")
    print(f"错误行号: {result['line_number']}")
```

---

### ★4. 智能执行（自动选择最佳方式）

```python
# 智能执行宏（自动处理理参数、监控、错误处理）
result = reader.smart_run(
    macro_name="FillExcelTemplate",
    auto_save=True,        # 执行后自动保存
    auto_backup=True,      # 执行前自动备份
    timeout_seconds=300,   # 超时设置
    retry_on_error=True,   # 错误时自动重试
    max_retries=3,         # 最大重试次数
    log_to_file=True,      # 记录到日志文件
)

# 输出执行报告
print(f"执行状态: {result['status']}")
print(f"执行时间: {result['duration']}秒")
print(f"生成的文件: {result['generated_files']}")
print(f"处理的数据行数: {result['rows_processed']}")
```

---

### 2. 提取参数信息

```python
# 提取宏的参数
params = reader.extract_parameters("mod_Main", "FillTemplate")

print(f"参数列表: {params}")
# 输出: ['templatePath', 'dataSource', 'outputPath', 'formatType']
```

---

### 3. 实时修改宏代码

```python
# 修改宏代码
reader.update_module("mod_Main", """
Sub FillTemplate(templatePath As String, dataSource As String)
    ' 优化后的代码
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")

    ' 使用字典提高性能
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    ' ... 新逻辑 ...
End Sub
""")

# 保存
reader.save()
```

---

### 4. 优化建议

```python
# 分析代码质量
analysis = reader.analyze_code("mod_Main")

print(f"代码行数: {analysis['line_count']}")
print(f"复杂度: {analysis['complexity']}")
print(f"建议: {analysis['suggestions']}")

# 自动优化
optimized_code = reader.optimize_module("mod_Main")
reader.update_module("mod_Main", optimized_code)
reader.save()
```

---

## 📋 核心 API

### VBAReader 类

```python
class VBAReader:
    def __init__(self, filepath: str, use_win32com: bool = True):
        """
        初始化 VBA 读取器

        Args:
            filepath: .xlsm 文件路径
            use_win32com: True 使用 win32com（可读可写）
                         False 使用 oletools（只读）
        """

    def get_all_modules(self) -> Dict[str, str]:
        """返回所有 VBA 模块代码"""

    def get_module(self, module_name: str) -> str:
        """返回指定模块代码"""

    def list_modules(self) -> List[str]:
        """列出所有模块名称"""

    def list_procedures(self) -> List[Dict]:
        """列出所有 Sub/Function 及参数"""

    def extract_parameters(self, module: str, procedure: str) -> List[str]:
        """提取指定过程的参数列表"""

    def update_module(self, module_name: str, code: str) -> bool:
        """更新模块代码（仅 win32com）"""

    def add_module(self, module_name: str, code: str) -> bool:
        """新增模块（仅 win32com）"""

    def delete_module(self, module_name: str) -> bool:
        """删除模块（仅 win32com）"""

    def analyze_code(self, module_name: str) -> Dict:
        """分析代码质量"""

    def optimize_module(self, module_name: str) -> str:
        """返回优化后的代码"""

    def compare_versions(self, old_code: str, new_code: str) -> Dict:
        """对比两个版本的代码差异"""

    def save(self) -> bool:
        """保存修改（仅 win32com）"""

    def close(self):
        """关闭文件"""
```

---

## 📋 应用场景

### 场景 1：你的宏目录系统优化

**问题：** VBA 代码性能较慢，需要优化

**解决方案：**
```python
reader = VBAReader(r"D:\工作目录\你的宏目录\你的工作簿.xlsm")

# 分析所有模块
for module in reader.list_modules():
    analysis = reader.analyze_code(module)

    if analysis['complexity'] > 10:
        print(f"{module} 复杂度过高: {analysis['complexity']}")
        print(f"建议: {analysis['suggestions']}")
```

---

### 场景 2：批量更新宏参数

**问题：** 多个 .xlsm 文件的宏参数需要统一修改

**解决方案：**
```python
import os

files = [
    r"<PROJECTS_ROOT>\ProjectA\报批表.xlsm",
    r"<PROJECTS_ROOT>\ProjectB\报批表.xlsm",
]

new_code = """
Sub FillTemplate(templatePath As String, dataSource As String, Optional formatType As String = "PDF")
    ' 新增 formatType 参数
    ' ...
End Sub
"""

for file in files:
    reader = VBAReader(file)
    reader.update_module("mod_Main", new_code)
    reader.save()
    reader.close()
    print(f"已更新: {file}")
```

---

### 场景 3：VBA 代码导出备份

**问题：** 将 VBA 宏备份为独立文件

**解决方案：**
```python
reader = VBAReader(r"D:\你的宏目录.xlsm")

# 提取所有 VBA 代码
modules = reader.get_all_modules()

# 保存为 .bas 文件
for name, code in modules.items():
    with open(f"{name}.bas", "w", encoding="utf-8") as f:
        f.write(code)
```

---

## 📋 已知限制

| 限制 | 说明 | 解决方案 |
|:---|:---|:---|
| 需要 Excel | win32com 方法需要安装 Excel | 使用 oletools 只读模式 |
| Windows only | win32com 仅支持 Windows | 使用 oletools 跨平台 |
| 密码保护 | 密码保护的 VBA 无法读取 | 需要密码才能解锁 |
| 大文件 | 超大 .xlsm 文件读取较慢 | 分批读取模块 |

---

## 📋 依赖项

```
pywin32>=310
oletools>=0.60
olefile>=0.47
```

---

## 📋 文件结构

```
skills/vba-reader/
├── SKILL.md                  # 本文件
├── README.md                 # 用户文档
├── package.json              # 包元数据
├── scripts/
│   └── vba_reader.py        # 核心读取器（包含所有功能）
├── examples/
│   ├── read_vba.py          # 读取示例
│   └── run_macro.py         # 执行示例
└── tests/
    └── test_reader.py       # 单元测试
```

---

_版本：v2.0.0 | 创建时间：2026-06-03_