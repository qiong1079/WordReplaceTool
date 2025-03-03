# WordReplaceTool
这是一个批量替换word文档字段的工具
# DocReplacePro - 专业文档批量替换工具

![Python Version](https://img.shields.io/badge/python-3.7%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)

一款专业的Word文档批量处理工具，支持智能内容替换、格式保留和跨文档类型处理，适用于法律文书、技术文档等专业场景的批量修改需求。

## 📌 核心功能

- **多格式支持**：原生处理`.docx`文件，兼容旧版`.doc`格式自动转换
- **格式保留**：完美保持原始文档的字体、段落、表格等格式
- **智能替换**：
  - 大小写智能匹配（自动识别全大写/首字母大写/普通格式）
  - 正则表达式支持
  - 跨Run内容替换
- **专业场景支持**：
  - 文档核心属性修改（标题、主题、关键词）
  - 页眉页脚处理（标准/首页/奇偶页）
  - 超链接和形状文本处理
- **详细报告**：
  - 实时替换统计面板
  - 完整错误日志记录
  - 处理结果可视化输出

## 📥 安装指南

### 环境要求
- Windows系统（需已安装Microsoft Word）
- Python 3.7+

### 快速安装
```bash
pip install python-docx pywin32
🚀 快速入门
配置规则

修改config.py中的设置：

Python
REPLACE_RULES = {
    "旧文本": "新文本",
    r"\b敏感词\b": "***",  # 正则表达式示例
    # 添加更多规则...
}

SOURCE_FOLDER = "输入目录路径"
OUTPUT_FOLDER = "输出目录路径"
运行程序

Bash
python main.py
查看结果

替换后的文档保存于输出目录
详细日志查看processing.log
🛠 专业配置
高级替换规则示例
Python
REPLACE_RULES = {
    # 大小写敏感替换
    "Confidential": "[REDACTED]",
    
    # 正则表达式匹配日期格式
    r"\d{4}-\d{2}-\d{2}": "YYYY-MM-DD",
    
    # 保留原始大小写的智能替换
    "CompanyName": "NewBrand",
}
命令行参数
Bash
python main.py [选项]
选项:
  --strict    严格模式（区分大小写）
  --verbose   显示详细调试信息
  --log PATH  指定日志文件路径
📊 输出示例
终端输出示例

⚠️ 重要提示
DOC文件处理：

需安装Microsoft Word 2007+
首次运行时授予Word宏权限
路径规范：

使用原始字符串表示路径：r"C:\My Documents"
安全建议：

处理前备份原始文档
避免在路径中包含中文或特殊字符
🤝 参与贡献
欢迎通过以下方式参与项目：

提交Issue报告问题
Fork项目并提交Pull Request
完善文档或翻译
📜 许可证
本项目采用 MIT License

专业用户推荐：建议配合文档版本控制系统使用，确保修改可追溯

