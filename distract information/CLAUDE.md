# 新三板失信人信息提取与财务数据处理项目

## 项目概述
本项目用于从NEEQ（全国中小企业股份转让系统，新三板）爬取失信被执行人公告PDF文件，并从中提取信息（姓名、公司、案号等），最终汇总整理为结构化财务数据。

## 数据流水线
1. **下载阶段** (`1.download files from NEEQ/`) — Selenium + Edge WebDriver 从 neeq.com.cn 爬取PDF
2. **PDF提取** (`2.distract information from pdf/`) — pdfplumber + jieba/LAC 从PDF提取失信人信息
3. **PDF转TXT** (`3.pdf transform txt/`) — PDF格式转换（中间步骤，当前流水线不依赖此步骤）
4. **TXT提取** (`4.distract information from txt/`) — 从TXT文件提取失信人及案号信息（独立提取路径）
5. **财务排序** (`5.financial date sorting/`) — 新三板公司财务数据按模板整理
6. **补充数据** (`6.supplement/`) — 失信人员名单爬虫 + 补充处理
7. **数据整理** (`7.collate date/`) — 最终汇总、财务数据合并及统计

## 关键文件
- `7.collate date/2.财务数据汇总.py` — 核心汇总脚本：将4张财报 + 基本信息 + 人员构成 + 研发数据按时间双向匹配合并
- `7.collate date/1.整理数据.py` — 数据清洗与模糊匹配（thefuzz），生成面板格式数据
- `样表-结构分析/` — 7张Excel模板文件，定义了数据字段结构和列映射关系

## 技术栈
- **爬虫**: selenium, requests, Edge WebDriver
- **PDF解析**: pdfplumber
- **中文NLP**: jieba (分词), LAC (百度词法分析，命名实体识别)
- **数据处理**: pandas, openpyxl, python-docx
- **模糊匹配**: thefuzz

## 代码约定
- 脚本按编号和功能命名（如 `1.3 提取pdf中的信息（lac）.py`）
- 文件路径硬编码在脚本顶部（注意运行时修改路径）
- 姓氏表路径: `D:\This computer\desktop\姓氏表\姓氏表.docx`
- 使用中文注释和日志输出
- 日志文件生成在脚本同目录下（如 `extract_*.log`）

## 关键依赖关系
- `7.collate date/2.财务数据汇总.py` → 需要 `样表-结构分析/` 中的Excel模板确定字段结构
- `7.collate date/1.整理数据.py` → 输入为各批次提取结果Excel文件
- PDF提取脚本 → 需要姓氏表 docx 文件用于姓名验证
- 下载脚本 → 需要 Edge WebDriver 及手动登录NEEQ网站

## 输出文件
- `失信人信息提取结果_完整版.csv` — PDF/TXT提取的原始结果
- `失信数据汇总结果_面板格式.xlsx` — 整理后的面板数据
- `需人工识别.xlsx` — 无法自动匹配的记录
- `失信数据汇总_含财务数据*.xlsx` — 最终含财务数据的汇总结果
