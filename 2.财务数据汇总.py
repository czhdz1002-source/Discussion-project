import pandas as pd
import os
import gc


# ======================== 路径配置 ========================
FINANCIAL_DATA_DIR = r"D:\This computer\desktop\三表数据"
DISHONEST_DATA_PATH = r"D:\This computer\desktop\失信数据汇总结果_面板格式2.xlsx"
OUTPUT_PATH_PHASE1 = r"D:\This computer\desktop\失信数据汇总_含财务数据5_阶段1.xlsx"
OUTPUT_PATH_FINAL = r"D:\This computer\desktop\失信数据汇总_含财务数据5.xlsx"


# ======================== 工具函数 ========================
def clean_code(x):
    """清理证券代码：取6位数字，左侧补零"""
    if pd.isna(x) or str(x).lower() == 'nan' or str(x).strip() == '':
        return None
    return str(x).split('.')[0].strip().zfill(6)


def parse_report_date(date_col):
    """解析统计截止日期：处理 Excel 数字日期和字符串日期"""
    clean_str = date_col.astype(str).str.split('.').str[0]
    return pd.to_datetime(clean_str, errors='coerce')


def normalize_code_column(df):
    """统一代码列名：股票代码 → 证券代码"""
    if '股票代码' in df.columns:
        df = df.rename(columns={'股票代码': '证券代码'})
    return df


# ======================== 各表加载函数 ========================
def load_dishonesty_data(path):
    """加载失信汇总数据并预处理"""
    if not os.path.exists(path):
        raise FileNotFoundError(f"找不到文件: {path}")

    df = pd.read_excel(path, sheet_name='多次公告(面板)')
    df['证券代码'] = df['证券代码'].apply(clean_code)
    df['公告时间_DT'] = pd.to_datetime(df['公告时间'], errors='coerce')
    df = df.dropna(subset=['公告时间_DT', '证券代码']).sort_values('公告时间_DT')
    df = df.reset_index(drop=True)
    return df


def prepare_financial_table(filepath, prefix, target_codes):
    """
    加载单张财务报表（利润表/资产负债表/现金流量表）。
    仅保留 报表类型='A'，给数据列加前缀，返回精简表。
    """
    df = pd.read_excel(filepath)
    total_rows = len(df)

    df = normalize_code_column(df)
    df['证券代码'] = df['证券代码'].apply(clean_code)
    df = df[df['证券代码'].isin(target_codes)].copy()

    if '报表类型' in df.columns:
        df = df[df['报表类型'] == 'A'].copy()

    df['报告期_DT'] = parse_report_date(df['统计截止日期'])
    df = df.dropna(subset=['报告期_DT', '证券代码'])

    # 给数据列加前缀
    key_cols = {'证券代码', '统计截止日期', '报表类型', '公告来源', '上市公司ID', '报告期_DT'}
    data_cols = [c for c in df.columns if c not in key_cols]
    rename_map = {c: f"{prefix}{c}" for c in data_cols}
    df = df.rename(columns=rename_map)

    keep = ['证券代码', '报告期_DT'] + list(rename_map.values())
    df = df[[c for c in keep if c in df.columns]]

    # 去重
    df = df.groupby(['证券代码', '报告期_DT'], as_index=False).first()

    print(f"   [{prefix}] 原始 {total_rows} → 保留 {len(df)} 行")
    return df


def merge_financial_tables(filepath_map, target_codes):
    """
    将4张财务报表合并为统一的财务宽表。
    每张表独立加载后，按 (证券代码, 报告期_DT) 做 outer join。
    返回：一行包含同一报告期所有财务报表数据的宽表。
    """
    tables = []
    for filename, prefix in filepath_map.items():
        filepath = os.path.join(FINANCIAL_DATA_DIR, filename)
        if not os.path.exists(filepath):
            print(f"   [警告] 文件不存在: {filepath}，跳过")
            continue
        df = prepare_financial_table(filepath, prefix, target_codes)
        tables.append((prefix.strip('_'), df))

    if not tables:
        return pd.DataFrame(columns=['证券代码', '报告期_DT'])

    # 以第一张表为底，逐一 outer join
    unified = tables[0][1]
    for _, tdf in tables[1:]:
        unified = unified.merge(tdf, on=['证券代码', '报告期_DT'], how='outer')
        gc.collect()

    unified = unified.sort_values('报告期_DT')
    print(f"   [财务合并] 统一宽表共 {len(unified)} 行 x {len(unified.columns)} 列")
    return unified


def prepare_basic_info_table(filepath, target_codes):
    """加载上市公司年度基本信息表"""
    df = pd.read_excel(filepath)
    total_rows = len(df)

    df = normalize_code_column(df)
    df['证券代码'] = df['证券代码'].apply(clean_code)
    df = df[df['证券代码'].isin(target_codes)].copy()

    df['报告期_DT'] = parse_report_date(df['统计截止日期'])
    df = df.dropna(subset=['报告期_DT', '证券代码'])

    key_cols = {'证券代码', '统计截止日期', '上市公司ID', '报告期_DT'}
    data_cols = [c for c in df.columns if c not in key_cols]
    rename_map = {c: f"基本信息_{c}" for c in data_cols}
    df = df.rename(columns=rename_map)

    keep = ['证券代码', '报告期_DT'] + list(rename_map.values())
    df = df[[c for c in keep if c in df.columns]]

    df = df.groupby(['证券代码', '报告期_DT'], as_index=False).first()
    df = df.sort_values('报告期_DT')

    print(f"   [基本信息] 原始 {total_rows} → 保留 {len(df)} 行")
    return df


def prepare_personnel_table(filepath, target_codes):
    """
    加载人员构成表，仅保留"员工总计"行。
    同 key 多行：期初/期末人数一致则合并，不一致则填 "default"。
    """
    df = pd.read_excel(filepath)
    total_rows = len(df)

    df = normalize_code_column(df)
    df['证券代码'] = df['证券代码'].apply(clean_code)
    df = df[df['证券代码'].isin(target_codes)].copy()

    mask = df['人员明细'].astype(str).str.contains('员工总计', na=False)
    df = df[mask].copy()

    df['报告期_DT'] = parse_report_date(df['统计截止日期'])
    df = df.dropna(subset=['报告期_DT', '证券代码'])

    agg_rows = []
    for (code, dt), group in df.groupby(['证券代码', '报告期_DT']):
        qc_vals = group['期初人数'].dropna().unique()
        qm_vals = group['期末人数'].dropna().unique()

        agg_rows.append({
            '证券代码': code,
            '报告期_DT': dt,
            '人员构成_期初人数': qc_vals[0] if len(qc_vals) == 1 else (
                'default' if len(qc_vals) > 1 else None),
            '人员构成_期末人数': qm_vals[0] if len(qm_vals) == 1 else (
                'default' if len(qm_vals) > 1 else None),
        })

    result = pd.DataFrame(agg_rows)
    if not result.empty:
        result = result.sort_values('报告期_DT')

    print(f"   [人员构成] 原始 {total_rows} → 员工总计 {len(df)} → 聚合 {len(result)} 行")
    return result


def prepare_rd_table(filepath, target_codes):
    """
    加载研发人员情况表，同(证券代码, 日期)下聚合教育程度分布。
    """
    df = pd.read_excel(filepath)
    total_rows = len(df)

    df = normalize_code_column(df)
    df['证券代码'] = df['证券代码'].apply(clean_code)
    df = df[df['证券代码'].isin(target_codes)].copy()

    df['报告期_DT'] = parse_report_date(df['统计截止日期'])
    df = df.dropna(subset=['报告期_DT', '证券代码'])

    agg_rows = []
    for (code, dt), group in df.groupby(['证券代码', '报告期_DT']):
        parts = []
        for _, row in group.iterrows():
            edu = row.get('教育程度', '')
            edu = str(edu) if pd.notna(edu) else ''
            val = row.get('期末人数/比例', '')
            if pd.notna(val) and edu:
                parts.append(f"{edu}:{val}")
        agg_rows.append({
            '证券代码': code,
            '报告期_DT': dt,
            '研发人员_教育程度分布': ', '.join(parts) if parts else None,
        })

    result = pd.DataFrame(agg_rows)
    if not result.empty:
        result = result.sort_values('报告期_DT')

    print(f"   [研发人员] 原始 {total_rows} → 聚合 {len(result)} 行")
    return result


# ======================== merge_asof 封装 ========================
def merge_table(result_df, table_df, label, direction='backward', add_prefix=None):
    """
    将一张预处理好的表 merge_asof 到结果中。
    - direction: 'backward'（公告前最近一期）或 'forward'（公告后最近一期）
    - add_prefix: 可选，给右表所有非 key 列加上方向前缀（如 '公告前_'）
    """
    result_df = result_df.sort_values('公告时间_DT')
    table_df = table_df.sort_values('报告期_DT')

    if add_prefix:
        key_cols = {'证券代码', '报告期_DT'}
        rename = {c: f"{add_prefix}{c}" for c in table_df.columns if c not in key_cols}
        table_df = table_df.rename(columns=rename)

    n_cols_before = len(result_df.columns)

    result_df = pd.merge_asof(
        result_df,
        table_df,
        left_on='公告时间_DT',
        right_on='报告期_DT',
        by='证券代码',
        direction=direction
    )

    # 保留匹配到的报告期日期，以便后续核对
    if '报告期_DT' in result_df.columns:
        date_label = f"{add_prefix}" if add_prefix else f"{label}_"
        result_df = result_df.rename(columns={'报告期_DT': f'{date_label}报告期_DT'})

    n_new_cols = len(result_df.columns) - n_cols_before
    direction_name = {'backward': '公告前', 'forward': '公告后'}.get(direction, direction)
    print(f"   merge [{label}]({direction_name}) → +{n_new_cols} 列，当前 {len(result_df)} 行 x {len(result_df.columns)} 列")
    return result_df


# ======================== 主流程 ========================
def main():
    print("=" * 70)
    print("  失信数据 + 财务数据 汇总脚本 v3（双向匹配 + 财务合并）")
    print("=" * 70)

    # ---- [1/9] 加载失信汇总数据 ----
    print("\n[1/9] 加载失信汇总数据...")
    dishonest_df = load_dishonesty_data(DISHONEST_DATA_PATH)
    target_codes = set(dishonest_df['证券代码'].dropna().unique())
    print(f"   -> {len(target_codes)} 家有失信记录的公司，{len(dishonest_df)} 条公告")

    result_df = dishonest_df.copy()

    # ---- [2/9] 合并4张财务报表 ----
    print("\n[2/9] 合并4张财务报表（利润表+资产负债表+现金流量表直接法+间接法）...")
    financial_files = {
        '1.利润表.xlsx':                '利润表_',
        '2.资产负债表.xlsx':             '资产负债表_',
        '3.现金流量表（直接法）.xlsx':    '现金流量表直接法_',
        '4.现金流量表（间接法）.xlsx':    '现金流量表间接法_',
    }
    unified_financial = merge_financial_tables(financial_files, target_codes)

    # ---- [3/9] 双向匹配：公告前 ----
    print("\n[3/9] 匹配【公告前】最近一期财务数据 (backward)...")
    result_df = merge_table(result_df, unified_financial, '统一财务',
                            direction='backward', add_prefix='公告前_')
    del unified_financial
    gc.collect()

    # ---- [4/9] 双向匹配：公告后 ----
    # 重新加载统一财务表（因为上一轮merge已消费）
    print("\n[4/9] 匹配【公告后】最近一期财务数据 (forward)...")
    unified_financial2 = merge_financial_tables(financial_files, target_codes)
    result_df = merge_table(result_df, unified_financial2, '统一财务',
                            direction='forward', add_prefix='公告后_')
    del unified_financial2
    gc.collect()

    # ---- [5/9] 基本信息表（公告前最近一期）----
    print("\n[5/9] 匹配 基本信息（公告前最近一期）...")
    filepath = os.path.join(FINANCIAL_DATA_DIR, '5.上市公司年度基本信息表.xlsx')
    if os.path.exists(filepath):
        info_df = prepare_basic_info_table(filepath, target_codes)
        result_df = merge_table(result_df, info_df, '基本信息', direction='backward')
        del info_df
        gc.collect()
    else:
        print(f"   [警告] 文件不存在，跳过")

    # ---- 阶段1输出 ----
    result_phase1 = result_df.drop(columns=['公告时间_DT'], errors='ignore')

    original_cols = list(dishonest_df.columns)
    original_cols = [c for c in original_cols if c != '公告时间_DT']
    new_cols = [c for c in result_phase1.columns if c not in original_cols]
    result_phase1 = result_phase1[original_cols + new_cols]

    print(f"\n--- 阶段1 输出（4财报 + 基本信息）---")
    print(f"  保存到: {OUTPUT_PATH_PHASE1}")
    result_phase1.to_excel(OUTPUT_PATH_PHASE1, index=False, float_format=None)
    print(f"  行数: {len(result_phase1)}, 列数: {len(result_phase1.columns)}")
    print(f"  新增列: {len(new_cols)}")

    # ---- [6/9] 人员构成表（公告前最近一期）----
    print(f"\n[6/9] 匹配 人员构成-员工总计（公告前最近一期）...")
    filepath = os.path.join(FINANCIAL_DATA_DIR, '6.上市公司人员构成表.xlsx')
    if os.path.exists(filepath):
        personnel_df = prepare_personnel_table(filepath, target_codes)
        if not personnel_df.empty:
            result_df = merge_table(result_df, personnel_df, '人员构成', direction='backward')
        else:
            print(f"   [警告] 无员工总计数据")
        del personnel_df
        gc.collect()
    else:
        print(f"   [警告] 文件不存在，跳过")

    # ---- [7/9] 研发人员情况（公告前最近一期）----
    print(f"\n[7/9] 匹配 研发人员情况（公告前最近一期）...")
    filepath = os.path.join(FINANCIAL_DATA_DIR, '7.研发人员情况.xlsx')
    if os.path.exists(filepath):
        rd_df = prepare_rd_table(filepath, target_codes)
        if not rd_df.empty:
            result_df = merge_table(result_df, rd_df, '研发人员', direction='backward')
        del rd_df
        gc.collect()
    else:
        print(f"   [警告] 文件不存在，跳过")

    # ---- 最终输出 ----
    print(f"\n[8/9] 保存最终结果...")
    result_final = result_df.drop(columns=['公告时间_DT'], errors='ignore')

    original_cols = [c for c in dishonest_df.columns if c != '公告时间_DT']
    final_new_cols = [c for c in result_final.columns if c not in original_cols]
    result_final = result_final[original_cols + final_new_cols]

    print(f"\n--- 最终输出 ---")
    print(f"  保存到: {OUTPUT_PATH_FINAL}")
    result_final.to_excel(OUTPUT_PATH_FINAL, index=False, float_format=None)
    print(f"  行数: {len(result_final)}, 列数: {len(result_final.columns)}")
    print(f"  总新增列: {len(final_new_cols)}")

    # ---- 匹配率诊断 ----
    print(f"\n--- 匹配率诊断 ---")
    for prefix in ['公告前_利润表_营业收入', '公告后_利润表_营业收入']:
        if prefix in result_final.columns:
            matched = result_final[prefix].notna().sum()
            print(f"  {prefix}: {matched}/{len(result_final)} = {matched/len(result_final)*100:.1f}%")
    for col in ['基本信息_证券简称', '人员构成_期末人数', '研发人员_教育程度分布']:
        if col in result_final.columns:
            matched = result_final[col].notna().sum()
            print(f"  {col}: {matched}/{len(result_final)} = {matched/len(result_final)*100:.1f}%")

    print(f"\n{'=' * 70}")
    print("  任务圆满完成！")
    print(f"{'=' * 70}")


if __name__ == "__main__":
    main()
