import os
import re
import pandas as pd
import pdfplumber
from pathlib import Path
from itertools import zip_longest

# ---------------------------- 配置区 ----------------------------
FOLDER_PATH = r'E:\Research\DATA SOURCES\新三板失信人\2017年失信人文件'
# 字段中文名及对应的可能出现在PDF中的多种表头/标签
FIELD_CONFIG = {
    '案号': ['案号', '案号'],
    '被执行人': ['被执行人', '被执行人姓名/名称', '姓名/名称', '被执行人姓名','被执行人（姓名/名称）','失信被执行人名称'],
    '执行依据文号': ['执行依据文号', '执行依据文号'],
    '立案时间': ['立案时间', '立案时间'],
    '出执行依据单位': ['出执行依据单位', '执行依据单位'],
    '失信被执行人行为情形': ['失信被执行人行为具体情形', '失信被执行人行为情形', '行为情形'],
    '信息发布日期': ['信息发布日期', '发布日期'],
    '证券代码': ['证券代码', '证券代码'],
    '证券简称': ['证券简称', '证券简称'],
    '被执行人履行情况': ['被执行人履行情况', '履行情况'],
    '生效法律文书确定的义务': ['生效法律文书确定的义务', '法律生效文书确定的义务'],
    '处理办法': ['处理办法', '处理方式'],                     # 新增字段：处理办法
    '已撤销': [],                                            # 新增字段：从文件名提取，不用于文本/表格匹配
}

# 用于匹配字段名后面内容的通用正则（非贪婪，遇到换行停止）
FIELD_PATTERN = r'{field}[：:]\s*(.*?)(?:\n|$)'

# 履行情况的关键词
STATUS_KEYWORDS = ['全部未履行', '部分未履行']

# 上下文模式配置（用于没有明确标签的情况）
CONTEXT_PATTERNS = {
    '被执行人': [
        r'(?P<position>董事|董事长)\s*(?P<name>.*?)(?:女士|先生)?\s*被纳入失信被执行人'
    ]
}

# 撤销关键词（用于修改义务字段）
REVOKE_KEYWORDS = ['撤销被纳入失信', '撤销失信', '失信被撤销']

# ---------------------------- 工具函数 ----------------------------
def extract_filename_prefix(filename):
    """提取文件名开头的数字编号作为来源"""
    match = re.match(r'^(\d+)', filename)
    return match.group(1) if match else filename

def get_filename_status(filename):
    """根据文件名判断状态：已撤销 / 被纳入 / 空"""
    revoke_keywords = ['撤销', '被撤销', '已撤销', '已被撤销']
    include_keywords = ['纳入', '被纳入', '列入', '被列入','被纳','被列为','属于失信']
    exclude_keywords =['不属于失信','不为失信','不是失信','从失信被执行人名单库中删除']
    for kw in exclude_keywords:
        if kw in filename:
            return '不属于'
    for kw in include_keywords:
        if kw in filename:
            return '被纳入'
    for kw in revoke_keywords:
        if kw in filename:
            return '已撤销'
    return ''

def find_table_by_headers(tables, headers_map):
    """
    在表格列表中查找包含指定表头的表格
    headers_map: {标准字段名: [可能的表头列表]}
    返回 (表格数据, 列索引映射) 或 (None, None)
    """
    for table in tables:
        if not table or len(table) < 2:
            continue
        header_row = table[0]
        col_mapping = {}
        for std_field, variants in headers_map.items():
            for idx, cell in enumerate(header_row):
                if cell and any(var in cell for var in variants):
                    col_mapping[std_field] = idx
                    break
        if col_mapping:
            return table, col_mapping
    return None, None

def extract_from_table(table, col_mapping):
    """从表格中按行提取记录，返回记录列表（每行为一个字典）"""
    records = []
    for row in table[1:]:
        record = {}
        for std_field, col_idx in col_mapping.items():
            if col_idx < len(row):
                val = row[col_idx]
                if val is not None:
                    val = str(val).strip()
                else:
                    val = ''
                record[std_field] = val
            else:
                record[std_field] = ''
        records.append(record)
    return records

def extract_handling_method(text):
    """从文本中提取处理办法：改选、另聘、改选或另聘（优先级：改选或另聘 > 改选 > 另聘）"""
    if re.search(r'改选\s*[和或]\s*另聘', text):
        return ['改选或另聘']
    if re.search(r'改选', text):
        return ['改选']
    if re.search(r'另聘', text):
        return ['另聘']
    return ['']

def extract_obligation(text):
    """
    提取生效法律文书确定的义务（跨行段落，直到下一个字段或空行）
    """
    # 收集所有字段标签变体（用于识别结束位置）
    all_variants = []
    for variants in FIELD_CONFIG.values():
        all_variants.extend(variants)
    # 过滤掉空字符串，并构建正则
    next_field_pattern = '|'.join(re.escape(v) for v in all_variants if v)
    # 义务字段自身的变体
    ob_variants = FIELD_CONFIG.get('生效法律文书确定的义务', ['生效法律文书确定的义务'])
    for variant in ob_variants:
        # 匹配字段标签后的内容，直到遇到下一个字段标签、连续两个换行或文本结尾
        pattern = rf'{re.escape(variant)}[：:]\s*(.*?)(?=\n\s*(?:{next_field_pattern})[：:]|\n\s*\n|$)'
        match = re.search(pattern, text, re.MULTILINE | re.DOTALL)
        if match:
            return [match.group(1).strip()]
    return ['']

def extract_status(text):
    """提取被执行人履行情况，优先查找关键词"""
    for kw in STATUS_KEYWORDS:
        if kw in text:
            return [kw]
    pattern = r'被执行人履行情况[：:]\s*(.*?)(?:\n|$)'
    matches = re.findall(pattern, text, re.MULTILINE)
    if matches:
        return [matches[0].strip()]
    return ['']

def extract_positions_names_by_context(text, patterns):
    """从文本中提取职位和人名对"""
    names = []
    positions = []
    for pattern in patterns:
        for match in re.finditer(pattern, text):
            pos = match.group('position')
            name_part = match.group('name').strip()
            individual_names = re.split(r'[、，,]\s*', name_part)
            for n in individual_names:
                n = n.strip()
                if n:
                    names.append(n)
                    positions.append(pos)
    return {'被执行人': names, '职位': positions}

def extract_date_from_end(text):
    """从文本末尾提取日期（用于信息发布日期）"""
    lines = text.splitlines()
    date_patterns = [
        r'(\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)',
        r'(\d{4}\s*-\s*\d{1,2}\s*-\s*\d{1,2})',
        r'(\d{4}\s*/\s*\d{1,2}\s*/\s*\d{1,2})',
        r'(\d{4}\s*年\s*\d{1,2}\s*月)',
    ]
    for line in reversed(lines):
        line = line.strip()
        if not line:
            continue
        for pat in date_patterns:
            match = re.search(pat, line)
            if match:
                raw_date = match.group(1).strip()
                return [raw_date]

    return ['']

def expand_record(record_dict, source):
    """
    将可能包含多值列表的字典展开为多条记录
    record_dict: {字段名: [值列表]}，其中来源是单值
    返回 [记录1, 记录2, ...]
    """
    max_len = max((len(v) for v in record_dict.values() if isinstance(v, list)), default=1)
    expanded = []
    for i in range(max_len):
        row = {'来源': source}
        for field, values in record_dict.items():
            if isinstance(values, list):
                if i < len(values):
                    val = values[i]
                elif values:
                    val = values[-1]
                else:
                    val = ''
                row[field] = val
            else:
                row[field] = values
        expanded.append(row)
    return expanded

def clean_stock_code(val):
    """从字符串中提取6位数字作为证券代码"""
    if pd.isna(val) or val == '':
        return ''
    match = re.search(r'(\d{6})', str(val))
    return match.group(1) if match else val

def extract_from_text(text):
    """
    从纯文本中提取字段值（可能多值）
    返回字典 {字段名: [值列表]}
    """
    result = {}
    for std_field, variants in FIELD_CONFIG.items():
        # 处理办法使用关键词提取
        if std_field == '处理办法':
            result[std_field] = extract_handling_method(text)
            continue

        # 证券简称特殊处理：提取直到“主办券商”或换行
        if std_field == '证券简称':
            for variant in variants:
                pattern = rf'{re.escape(variant)}[：:]\s*(.*?)(?=\s*主办券商|\n|$)'
                match = re.search(pattern, text, re.DOTALL | re.MULTILINE)
                if match:
                    result[std_field] = [match.group(1).strip()]
                    break
            else:
                result[std_field] = ['']
            continue

        values = []
        if not variants:
            continue
        # 尝试第一个变体
        label = variants[0]
        pattern = FIELD_PATTERN.format(field=re.escape(label))
        matches = re.findall(pattern, text, re.MULTILINE)
        for m in matches:
            cleaned = m.strip()
            if cleaned:
                values.append(cleaned)
        # 若无匹配，尝试其他变体
        if not values and len(variants) > 1:
            for alt_label in variants[1:]:
                pattern = FIELD_PATTERN.format(field=re.escape(alt_label))
                matches = re.findall(pattern, text, re.MULTILINE)
                for m in matches:
                    cleaned = m.strip()
                    if cleaned:
                        values.append(cleaned)
                if values:
                    break
        result[std_field] = values if values else ['']

    return result

def process_pdf(pdf_path, filename):
    """
    处理单个PDF，返回记录列表（每个记录是一个字典，字段为单值）
    """
    source = extract_filename_prefix(filename)
    filename_status = get_filename_status(filename)

    try:
        with pdfplumber.open(pdf_path) as pdf:
            all_text = ''
            all_tables = []
            for page in pdf.pages:
                page_text = page.extract_text() or ''
                all_text += page_text + '\n'
                tables = page.extract_tables()
                all_tables.extend(tables)
    except Exception as e:
        print(f"读取PDF失败 {pdf_path}: {e}")
        return []

    # 1. 优先从表格提取结构化数据
    table, col_mapping = find_table_by_headers(all_tables, FIELD_CONFIG)
    if table:
        records = extract_from_table(table, col_mapping)
        for rec in records:
            rec['来源'] = source
            rec['已撤销'] = filename_status
            if '证券代码' in rec:
                rec['证券代码'] = clean_stock_code(rec['证券代码'])
        return records

    # 2. 无合适表格，从文本提取
    obligation = extract_obligation(all_text)
    status = extract_status(all_text)

    extracted = extract_from_text(all_text)
    extracted['生效法律文书确定的义务'] = obligation
    extracted['被执行人履行情况'] = status
    extracted['已撤销'] = [filename_status]

    # 3. 如果被执行人字段为空，尝试上下文提取
    if not extracted.get('被执行人') or extracted['被执行人'] == ['']:
        context_data = extract_positions_names_by_context(all_text, CONTEXT_PATTERNS['被执行人'])
        if context_data['被执行人']:
            extracted['被执行人'] = context_data['被执行人']
            extracted['职位'] = context_data['职位']

       # 4. 如果信息发布日期为空，尝试从文本末尾提取日期
    # 4. 如果信息发布日期为空，尝试从文本末尾提取日期
    if not extracted.get('信息发布日期') or extracted['信息发布日期'] == ['']:
        date_from_end = extract_date_from_end(all_text)
        if date_from_end and date_from_end != ['']:
            extracted['信息发布日期'] = date_from_end
        else:
            # 保护性检查：确保 all_text 非空且至少有一行
            lines = all_text.strip().splitlines()
            if lines:
                last_line = lines[-1].strip()
                if re.search(r'\d', last_line):
                    extracted['信息发布日期'] = [last_line]
                    print(f"   [调试] 使用最后一行作为日期: {last_line}")

    # 5. 展开多值记录
    records = expand_record(extracted, source)

    # 6. 清洗每条记录的证券代码
    for rec in records:
        if '证券代码' in rec:
            rec['证券代码'] = clean_stock_code(rec['证券代码'])

    return records

# ---------------------------- 主流程 ----------------------------
def main():
    folder = Path(FOLDER_PATH)
    if not folder.exists():
        print(f"文件夹不存在: {FOLDER_PATH}")
        return

    pdf_files = list(folder.glob('*.pdf')) + list(folder.glob('*.PDF'))
    pdf_files.sort()
    total = len(pdf_files)
    if total == 0:
        print("未找到PDF文件")
        return

    print(f"共找到 {total} 个PDF文件")

    temp_dir = folder / 'temp'
    temp_dir.mkdir(exist_ok=True)

    batch_size = 5
    batch_records = []
    file_count = 0
    batch_index = 1

    for i, pdf_path in enumerate(pdf_files, 1):
        print(f"处理第 {i}/{total} 个文件: {pdf_path.name}")
        records = process_pdf(pdf_path, pdf_path.name)
        if not records:
            print(f"  未提取到数据，跳过")
            continue

        batch_records.extend(records)
        file_count += 1

        # 保存单个文件的临时Excel
        temp_file = temp_dir / f"{pdf_path.stem}_temp.xlsx"
        df_single = pd.DataFrame(records)
        df_single.to_excel(temp_file, index=False)

        if file_count == batch_size or i == total:
            print(f"  合并第 {batch_index} 批数据，共 {len(batch_records)} 条记录")
            df_batch = pd.DataFrame(batch_records)
            # 调整列顺序：来源 -> 已撤销 -> 职位 -> 被执行人 -> 其他字段
            base_cols = ['来源', '已撤销', '职位', '被执行人']
            other_cols = [c for c in FIELD_CONFIG.keys() if c not in base_cols]
            cols = [c for c in base_cols + other_cols if c in df_batch.columns]
            df_batch = df_batch[cols]
            output_path = folder / f"失信人数据_合并{batch_index}.xlsx"
            df_batch.to_excel(output_path, index=False)
            print(f"  已保存至 {output_path}")

            # 删除该批次对应的临时文件
            for j in range(batch_index * batch_size - batch_size, min(batch_index * batch_size, total)):
                if j < len(pdf_files):
                    temp_name = pdf_files[j].stem + "_temp.xlsx"
                    temp_file_path = temp_dir / temp_name
                    if temp_file_path.exists():
                        temp_file_path.unlink()
                        print(f"  删除临时文件 {temp_name}")

            # 重置批次
            batch_records = []
            file_count = 0
            batch_index += 1

    # 清理临时目录
    try:
        temp_dir.rmdir()
    except OSError:
        pass

    # 合并所有批次文件
    print("开始合并所有批次文件...")
    merged_files = list(folder.glob('失信人数据_合并*.xlsx'))
    if merged_files:
        all_dfs = []
        for f in merged_files:
            df = pd.read_excel(f)
            all_dfs.append(df)
        total_df = pd.concat(all_dfs, ignore_index=True)
        base_cols = ['来源', '已撤销', '职位', '被执行人']
        other_cols = [c for c in FIELD_CONFIG.keys() if c not in base_cols]
        cols = [c for c in base_cols + other_cols if c in total_df.columns]
        total_df = total_df[cols]
        # 修改最终合并文件名为“失信人数据_合并.xlsx”
        total_output = folder / '失信人数据_合并.xlsx'
        total_df.to_excel(total_output, index=False)
        print(f"所有批次合并完成，总记录数：{len(total_df)}，已保存至 {total_output}")
    else:
        print("未找到任何合并文件，跳过总合并。")

    print("处理完成！")

if __name__ == "__main__":
    main()