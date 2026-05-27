import os
import re
import pandas as pd
import pdfplumber  # 用于读取PDF最后一行
from pathlib import Path

# ---------------------------- 配置 ----------------------------
FOLDER_PATH = Path(r'E:\Research\DATA SOURCES\新三板失信人\2017年失信人文件')
EXCEL_FILE = FOLDER_PATH / '失信人数据_合并.xlsx'

EXEC_KEYWORDS = ['董事长', '董事', '股东', '实际控制人', '股东', '高级管理人员', '监视']
COMPANY_TRIGGER = '公司被'
COMPANY_EXCLUDE = '及公司被'

def extract_filename_prefix(filename):
    match = re.match(r'^(\d+)', filename)
    return match.group(1) if match else filename

def build_prefix_to_filename_map(folder):
    """构建 数字前缀 -> 完整Path对象 的映射"""
    pdf_files = list(folder.glob('*.pdf')) + list(folder.glob('*.PDF'))
    mapping = {}
    for pdf_path in pdf_files:
        prefix = extract_filename_prefix(pdf_path.name)
        mapping[prefix] = pdf_path  # 存储Path对象
    return mapping

def extract_case_number_from_filename(filename):
    """
    从文件名中尝试匹配标准案号格式，返回匹配到的第一个完整案号字符串。
    格式示例：(2016)川 0108 执2176 号
    若未匹配则返回空字符串。
    """
    provinces = r'京津沪渝冀豫云辽黑湘皖鲁苏浙赣鄂桂甘晋蒙陕吉闽贵粤川琼青宁新藏'
    pattern = rf'\((\d{{4}})\)\s*([{provinces}])\s*(\d{{4}})\s*执\s*(\d{{3,4,5}})\s*号'
    match = re.search(pattern, filename)
    if match:
        return match.group(0)
    return ''

def main():
    excel_path = EXCEL_FILE
    if not excel_path.exists():
        print(f"错误：文件不存在 {excel_path}")
        return

    # 1. 读取Excel（指定案号列为字符串，避免自动转换）
    try:
        df = pd.read_excel(excel_path, dtype={'来源': str, '案号': str})
    except PermissionError:
        print(f"无法读取文件 {excel_path}，请确认该文件未被其他程序（如Excel）占用，然后重试。")
        return
    print(f"原始数据行数：{len(df)}")

    # 2. 格式化来源为4位数字
    df['来源'] = df['来源'].apply(lambda x: str(x).strip().zfill(4) if pd.notna(x) else '')

    # 3. 构建文件名映射（来源 -> 完整路径）
    file_map = build_prefix_to_filename_map(FOLDER_PATH)
    print(f"共找到 {len(file_map)} 个PDF文件")

    # 初始化计数器
    exec_filled_count = 0
    case_filled_count = 0
    date_filled_count = 0

    # 4. 填补被执行人
    for idx, row in df.iterrows():
        if row['已撤销'] == '不属于':
            continue
        if pd.notna(row['被执行人']) and str(row['被执行人']).strip() != '':
            continue

        source = row['来源']
        if source not in file_map:
            print(f"警告：来源 {source} 未找到对应PDF文件")
            continue

        filename = file_map[source].name  # 只取文件名用于关键词检查
        has_exec_keyword = any(kw in filename for kw in EXEC_KEYWORDS)
        has_shixin = '失信被执行' in filename
        if has_exec_keyword and has_shixin:
            df.at[idx, '被执行人'] = '董高监控股人'
            exec_filled_count += 1
            print(f"来源 {source}：填充为“董高监控股人”（文件名：{filename}）")
            continue

        if COMPANY_TRIGGER in filename and COMPANY_EXCLUDE not in filename:
            df.at[idx, '被执行人'] = '公司'
            exec_filled_count += 1
            print(f"来源 {source}：填充为“公司”（文件名：{filename}）")
    print(f"被执行人填补完成，共补充 {exec_filled_count} 条记录")

    # 5. 填补案号
    if '案号' in df.columns:
        for idx, row in df.iterrows():
            if pd.isna(row['案号']) or str(row['案号']).strip() == '':
                source = row['来源']
                if source in file_map:
                    filename = file_map[source].name
                    case_num = extract_case_number_from_filename(filename)
                    if case_num:
                        df.at[idx, '案号'] = case_num
                        case_filled_count += 1
                        print(f"来源 {source}：从文件名提取案号“{case_num}”")
        print(f"案号填补完成，共补充 {case_filled_count} 条记录")
    else:
        print("警告：Excel文件中没有“案号”列，跳过案号提取。")

    # 6. 填补信息发布日期（从PDF最后一行提取）
    if '信息发布日期' in df.columns:
        for idx, row in df.iterrows():
            if pd.isna(row['信息发布日期']) or str(row['信息发布日期']).strip() == '':
                source = row['来源']
                if source in file_map:
                    pdf_path = file_map[source]
                    try:
                        with pdfplumber.open(pdf_path) as pdf:
                            last_line = None
                            # 从最后一页向前找最后一非空行
                            for page in reversed(pdf.pages):
                                text = page.extract_text()
                                if text:
                                    lines = text.splitlines()
                                    for line in reversed(lines):
                                        if line.strip():
                                            last_line = line.strip()
                                            break
                                if last_line:
                                    break
                            if last_line:
                                df.at[idx, '信息发布日期'] = last_line
                                date_filled_count += 1
                                print(f"来源 {source}：从PDF最后一行提取发布日期“{last_line}”")
                    except Exception as e:
                        print(f"读取PDF失败 {pdf_path}: {e}")
                        continue
        print(f"信息发布日期填补完成，共补充 {date_filled_count} 条记录")
    else:
        print("警告：Excel文件中没有“信息发布日期”列，跳过日期填补。")

    # 7. 保存更新后的Excel
    try:
        df.to_excel(excel_path, index=False)
        print(f"处理完成，已更新 {excel_path}")
    except PermissionError:
        print(f"无法写入文件 {excel_path}，请确认该文件未被其他程序（如Excel）占用，然后重试。")
        return

if __name__ == "__main__":
    main()