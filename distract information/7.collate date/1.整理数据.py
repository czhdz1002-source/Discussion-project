import pandas as pd
import os
import glob
from thefuzz import process, fuzz

# --- 固定工作路径 ---
TARGET_DIR = r"D:\This computer\desktop\整理数据"


def clean_code(x):
    """标准化证券代码格式"""
    if pd.isna(x) or str(x).lower() == 'nan' or str(x).strip() == '':
        return None
    code_str = str(x).split('.')[0].strip()
    return code_str.zfill(6) if code_str.isdigit() else code_str


def parse_excel_date(x):
    """智能时间解析引擎：解决 1970-01-01 错误"""
    if pd.isna(x) or str(x).lower() == 'nan' or str(x).strip() == '':
        return pd.NaT

    # 1. 如果本身已经被 read_excel 解析为正确的日期时间格式
    if hasattr(x, 'strftime'):
        return pd.to_datetime(x)

    # 2. 核心修复：如果是 Excel 的纯数字天数（如 43332）
    # Excel 的日期起始点是 1899-12-30，通常现代日期的数字大于 30000
    if isinstance(x, (int, float)):
        if x > 10000:
            return pd.to_datetime(x, unit='D', origin='1899-12-30')

    # 3. 如果是字符串格式
    if isinstance(x, str):
        x_str = x.strip()
        # 有些数字会被读成带小数点的字符串，如 "43332.0"
        if x_str.replace('.', '', 1).isdigit():
            num = float(x_str)
            if num > 10000:
                return pd.to_datetime(num, unit='D', origin='1899-12-30')

        # 普通的文本日期，如 "2018-08-20"
        return pd.to_datetime(x_str, errors='coerce')

    # 终极兜底
    return pd.to_datetime(str(x), errors='coerce')


def main():
    if not os.path.exists(TARGET_DIR):
        print(f"错误：找不到指定的路径 -> {TARGET_DIR}")
        return

    file_list = glob.glob(os.path.join(TARGET_DIR, "*.xlsx"))
    file_list = [f for f in file_list if "汇总结果" not in f and "需人工识别" not in f]

    if not file_list:
        print(f"提示：在文件夹中未发现可处理的 Excel 文件。")
        return

    print(f"共发现 {len(file_list)} 个文件，开始读取并合并...")

    all_dfs = []
    for path in file_list:
        file_name = os.path.basename(path)
        try:
            df = pd.read_excel(path)
            df = df.loc[:, ~df.columns.duplicated()]
            df['数据来源'] = file_name
            all_dfs.append(df)
            print(f"成功载入: {file_name}")
        except Exception as e:
            print(f"载入失败 {file_name}: {e}")

    if not all_dfs: return
    full_df = pd.concat(all_dfs, ignore_index=True)

    # ====== 修正后的时间转换逻辑 ======
    if '公告时间' in full_df.columns:
        # 生成专门的隐形列用于精准排序
        full_df['时间_SORT'] = full_df['公告时间'].apply(parse_excel_date)
        # 将我们能看到的列，变成干净的 YYYY-MM-DD 字符串，过滤掉时分秒
        full_df['公告时间'] = full_df['时间_SORT'].dt.strftime('%Y-%m-%d').fillna(full_df['公告时间'])
    # =================================

    full_df['证券代码'] = full_df['证券代码'].apply(clean_code)
    full_df['证券简称'] = full_df['证券简称'].astype(str).replace(['nan', 'None', ''], [None, None, None])

    mask_manual = full_df['证券代码'].isna() & full_df['证券简称'].isna()
    df_manual = full_df[mask_manual].copy()
    df_working = full_df[~mask_manual].copy()

    print("正在进行跨年份公司匹配...")
    known_map = df_working.dropna(subset=['证券代码', '证券简称']).drop_duplicates('证券简称')
    name_to_code = dict(zip(known_map['证券简称'].astype(str), known_map['证券代码']))
    unique_names = [str(n) for n in name_to_code.keys()]

    def get_final_uid(row):
        code, name = row['证券代码'], row['证券简称']
        if pd.notna(code): return str(code)
        if pd.isna(name): return None

        name_str = str(name)
        if name_str in name_to_code: return name_to_code[name_str]

        if unique_names:
            try:
                best_match = process.extractOne(name_str, unique_names, scorer=fuzz.token_set_ratio)
                if best_match and best_match[1] >= 90:
                    return name_to_code[best_match[0]]
            except:
                pass
        return name_str

    df_working['UID'] = df_working.apply(get_final_uid, axis=1)

    counts = df_working['UID'].value_counts()
    multi_uids = counts[counts > 1].index

    df_panel = df_working[df_working['UID'].isin(multi_uids)].copy()
    df_single = df_working[~df_working['UID'].isin(multi_uids)].copy()

    # 使用隐藏的时间列进行精确排序
    if '时间_SORT' in df_panel.columns:
        df_panel = df_panel.sort_values(by=['UID', '时间_SORT'], ascending=[True, True])
        df_panel = df_panel.drop(columns=['时间_SORT'])
        df_single = df_single.drop(columns=['时间_SORT'])
    else:
        df_panel = df_panel.sort_values(by=['UID'], ascending=[True])

    output_main = os.path.join(TARGET_DIR, '失信数据汇总结果_面板格式.xlsx')
    output_manual = os.path.join(TARGET_DIR, '需人工识别.xlsx')

    try:
        with pd.ExcelWriter(output_main) as writer:
            df_panel.to_excel(writer, sheet_name='多次公告(面板)', index=False)
            df_single.to_excel(writer, sheet_name='单个公告', index=False)

        if not df_manual.empty:
            df_manual.to_excel(output_manual, index=False)
            print(f"⚠️ 发现 {len(df_manual)} 条信息不全数据，已保存至: 需人工识别.xlsx")

        print(f"\n🎉 处理成功！时间格式已完美修复。")
        print(f"总览数据：面板({len(df_panel)}条), 单个({len(df_single)}条)")
        print(f"结果存放在：{TARGET_DIR}")
    except Exception as e:
        print(f"保存失败！请确认结果文件没有被 Excel 打开：{e}")


if __name__ == "__main__":
    main()