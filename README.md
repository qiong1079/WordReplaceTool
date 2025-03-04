# DocReplacePro - 专业文档批量替换工具

![Python Version](https://img.shields.io/badge/python-3.7%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)
![Platform](https://img.shields.io/badge/platform-windows-lightgrey)

## 📌 核心功能

- **多格式支持**：原生处理`.docx`文件，兼容旧版`.doc`格式自动转换
- **格式保留**：完美保持原始文档的字体、段落、表格等格式
- **智能替换**：
  - 大小写智能匹配（自动识别全大写/首字母大写/普通格式）
  - 正则表达式支持
  - 跨Run内容替换
- **专业处理**：
  - 文档属性（标题/主题/关键词）
  - 页眉页脚（标准/首页/奇偶页）
  - 超链接和形状文本
- **详细报告**：
  - 实时替换统计面板
  - 完整错误日志记录
  - 可视化处理结果输出

## 📥 安装指南

### 环境要求
- Windows系统（需已安装Microsoft Word）
- Python 3.7+

### 依赖安装
```bash
pip install python-docx pywin32
🚀 快速开始
三步基础操作
创建配置文件

在项目根目录新建config.py：

Python
REPLACE_RULES = {
    "旧文本": "新文本",
    r"\b\d{11}\b": "[手机号]"  # 正则表达式示例
}

SOURCE_FOLDER = r"E:\原始文档"
OUTPUT_FOLDER = r"E:\处理结果"
执行处理程序

Bash
python main.py
查看处理结果

输出目录生成替换后的文档
根目录生成processing.log日志文件
首次运行示例
Bash
# 项目初始结构
DocReplacePro/
├── main.py
└── config.py

# 执行后输出
┌──────────────────────────────┬────────┬───────┐
│          文件名             │ 替换数 │ 状态  │
├──────────────────────────────┼────────┼───────┤
│ 合同_最终版.docx          │   5    │ 成功 │
│ 技术白皮书_v3.doc        │   12   │ 成功 │
└──────────────────────────────┴────────┴───────┘
⚙ 高级配置
文件结构说明
 DocReplacePro/
├── main.py         # 主程序入口
├── config.py       # 配置文件（必须）
├── processing.log  # 自动生成的日志文件
└── /output         # 处理结果输出目录
正则表达式配置示例
Python
# config.py
REPLACE_RULES = {
    # 日期脱敏
    r"\d{4}-\d{2}-\d{2}": "<DATE>",
    
    # 保留大小写的品牌替换
    "OldBrand": "NewBrand",
    
    # 多层级替换
    "[机密]": "[已审核]",
    "草案": "正式版"
}
📊 输出说明
控制台输出
Plaintext
┌──────────────────────────────┬────────┬───────┐
│          文件名             │ 替换数 │ 状态  │
├──────────────────────────────┼────────┼───────┤
│ 合同_最终版.docx          │   5    │ 成功 │
│ 失效文档.doc             │   --   │ 失败 │
└──────────────────────────────┴────────┴───────┘
处理完成：共处理 2 个文件 | 成功：1 | 失败：1
日志文件示例
Log
[2023-10-01 09:00:00] INFO: 开始处理 合同_最终版.docx
[2023-10-01 09:00:02] SUCCESS: 替换5处
[2023-10-01 09:00:05] ERROR: 失效文档.doc - 文件格式损坏
⚠️ 重要提示
系统要求：

必须启用Word宏权限（文件 > 选项 > 信任中心 > 宏设置）
推荐使用Windows 10/11系统
路径规范：

Python
# 正确格式（原始字符串）
SOURCE_FOLDER = r"C:\User\Documents"

# 错误格式（转义字符问题）
SOURCE_FOLDER = "C:\User\Documents"  # 会引发路径错误
安全建议：

处理前备份原始文档
敏感操作建议在虚拟机执行
🤝 参与贡献
欢迎通过以下方式参与：

提交Issue反馈问题
Fork仓库并提交PR
完善单元测试
更新多语言文档



📌 Core Features
Multi-Format Support: Native handling of .docx files with auto-conversion for legacy .doc format
Format Preservation: Maintain original document formatting including fonts, paragraphs, tables, etc.
Smart Replacement:
Case-sensitive matching (auto-detects ALL CAPS/Title Case/Regular case)
Regular expression support
Cross-Run content replacement
Professional Processing:
Document properties (Title/Subject/Keywords)
Headers & Footers (Standard/First Page/Odd-Even Pages)
Hyperlinks and shape texts
Detailed Reporting:
Real-time replacement statistics panel
Comprehensive error logging
Visualized processing results
📥 Installation Guide
Requirements
Windows OS (with Microsoft Word installed)
Python 3.7+
Install Dependencies
Bash
pip install python-docx pywin32
🚀 Quick Start
Three-Step Workflow
Create Configuration File
Create config.py in project root:

Python
REPLACE_RULES = {
    "Old Text": "New Text",
    r"\b\d{11}\b": "[PHONE]"  # Regex example
}

SOURCE_FOLDER = r"E:\SourceDocs"
OUTPUT_FOLDER = r"E:\ProcessedDocs"
Execute Processor
Bash
python main.py
View Results
Processed documents in output folder
processing.log generated in root directory
First-Run Example
Bash
# Initial structure
DocReplacePro/
├── main.py
└── config.py

# After execution
┌──────────────────────────────┬────────┬────────┐
│         Filename            │ Count  │ Status │
├──────────────────────────────┼────────┼────────┤
│ Contract_Final.docx        │   5    │ Success│
│ Whitepaper_v3.doc          │   12   │ Success│
└──────────────────────────────┴────────┴────────┘
⚙ Advanced Configuration
Project Structure
 DocReplacePro/
├── main.py         # Main entry
├── config.py       # Configuration (required)
├── processing.log  # Auto-generated log
└── /output         # Processed documents
Regex Configuration Example
Python
# config.py
REPLACE_RULES = {
    # Date masking
    r"\d{4}-\d{2}-\d{2}": "<DATE>",
    
    # Case-preserved brand replacement
    "OldBrand": "NewBrand",
    
    # Multi-level replacement
    "[Confidential]": "[Approved]",
    "Draft": "Final"
}
📊 Output Details
Console Display
 ┌──────────────────────────────┬────────┬────────┐
│         Filename            │ Count  │ Status │
├──────────────────────────────┼────────┼────────┤
│ Contract_Final.docx        │   5    │ Success│
│ CorruptedFile.doc          │   --   │ Failed │
└──────────────────────────────┴────────┴────────┘
Processing Complete: 2 files | Success: 1 | Failed: 1
Log File Example
 [2023-10-01 09:00:00] INFO: Processing Contract_Final.docx
[2023-10-01 09:00:02] SUCCESS: 5 replacements
[2023-10-01 09:00:05] ERROR: CorruptedFile.doc - File format corrupted
⚠ Important Notes
System Requirements:
Enable Word Macro permissions (File > Options > Trust Center > Macro Settings)
Recommended: Windows 10/11
Path Specifications:
Python
# Correct format (raw strings)
SOURCE_FOLDER = r"C:\User\Documents"

# Incorrect format (escape characters)
SOURCE_FOLDER = "C:\User\Documents"  # Causes path errors
Security Recommendations:
Backup original documents before processing
Perform sensitive operations in virtual machines
🤝 Contribution
We welcome contributions through:

Submit issues for bug reports
Fork repository and submit PRs
Improve unit tests
Enhance multilingual documentation
