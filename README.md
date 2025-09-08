# COSMIC 数据处理工具

这是一个专门用于处理COSMIC工作量评估相关Excel文件的自动化工具。该工具可以批量处理附件文件，包括文件重命名、单元格数据更新、需求内容概述生成以及WBS工作量分解。

## 项目结构

```
cosmic_data/
├── code/                          # 源代码目录
│   ├── process_attachments.py     # 主程序文件
│   ├── config_template.py         # 配置模板文件
│   ├── requirements.txt           # Python依赖包
│   ├── README.md                 # 代码说明文档
│   └── 一二三级功能点.xlsx        # 功能点码值参考表
├── data_file/                     # 数据文件目录
│   ├── 附件1-xxx@需求规格说明书.docx
│   ├── 附件2-xxx@WBS工作量分解表.xlsx
│   ├── 附件3-xxx@COSMIC工作量评估基础表.xlsx
│   ├── 附件4-xxx@工作量送审表.xlsx
│   └── 附件5-xxx@ams工作量.xls
└── README.md                     # 项目主说明文档
```

## 功能特性

### 🔄 自动化处理流程
1. **文件批量重命名** - 统一附件文件的需求名称
2. **需求信息更新** - 自动填充需求名称到相关单元格
3. **工作量计算** - 自动计算并分配工作量数据
4. **时间戳填充** - 自动填充送审日期
5. **功能点统计** - 自动统计功能点数量
6. **智能概述生成** - 使用AI生成需求内容概述
7. **WBS工作量分解** - 基于AI匹配更新WBS工作量文档

### 🤖 AI增强功能
- 使用DeepSeek API智能分析需求内容
- 自动匹配功能点码值
- 智能分配工作量估算
- 生成结构化需求概述

## 快速开始

### 1. 环境配置

```bash
# 克隆项目
git clone https://github.com/JN2020527/cosmic_data.git
cd cosmic_data

# 安装依赖
cd code
pip install -r requirements.txt
```

### 2. 配置设置

```bash
# 复制配置模板
cp config_template.py config.py

# 编辑配置文件，填入你的API密钥和数据路径
# config.py 文件内容示例：
# DEEPSEEK_API_KEY = "your-api-key"
# DATA_DIR = "/path/to/your/data_file"
```

### 3. 运行程序

```bash
python process_attachments.py
```

程序会提示输入需求名称，然后自动执行所有处理步骤。

## 系统要求

- Python 3.6+
- openpyxl
- xlrd (用于处理.xls文件)
- requests

## 文件格式支持

- ✅ `.xlsx` - 使用openpyxl处理
- ✅ `.xls` - 使用xlrd处理  
- ✅ `.docx` - 用于文件重命名

## 注意事项

- 确保 `data_file` 目录下包含所需的附件文件
- 附件5必须是.xls格式，其他附件建议使用.xlsx格式
- 需要有效的DeepSeek API密钥才能使用AI功能
- 程序会自动处理跨工作表的公式计算

## 贡献指南

欢迎提交Issue和Pull Request来改进这个项目。

## 许可证

本项目采用MIT许可证。

## 作者

- GitHub: [@JN2020527](https://github.com/JN2020527) 