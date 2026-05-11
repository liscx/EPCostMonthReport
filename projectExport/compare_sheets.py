# -*- coding: utf-8 -*-
import sys
import pandas as pd
import warnings

sys.stdout.reconfigure(encoding='utf-8')
warnings.filterwarnings('ignore', message='Workbook contains no default style')

FILE1 = '中间数据/合同汇总20260101-20260331.xlsx'
FILE2 = '中间数据/1-3月数据.xlsx'
SHEET1 = None  # None则读取第一个sheet
SHEET2 = None  # None则读取第一个sheet
OUTPUT = '结果数据/对照结果.xlsx'


def main():
    df1 = pd.read_excel(FILE1, sheet_name=SHEET1 if SHEET1 else 0)
    df2 = pd.read_excel(FILE2, sheet_name=SHEET2 if SHEET2 else 0)

    # 统一列名，忽略 sheet2 的序号列
    key = '合同编号'
    cols_compare = ['专区名称', '预算使用']

    df1[key] = df1[key].astype(str).str.strip()
    df2[key] = df2[key].astype(str).str.strip()

    set1 = set(df1[key])
    set2 = set(df2[key])

    # 1. 文件1 有，文件2 没有
    only_in_1 = df1[df1[key].isin(set1 - set2)].copy()
    only_in_1['来源'] = '仅文件1'

    # 2. 文件2 有，文件1 没有
    only_in_2 = df2[df2[key].isin(set2 - set1)].copy()
    only_in_2 = only_in_2[[key] + cols_compare]
    only_in_2['来源'] = '仅文件2'

    # 3. 两边都有，但数据不一致
    common = set1 & set2
    df1_common = df1[df1[key].isin(common)].drop_duplicates(subset=[key]).set_index(key)
    df2_common = df2[df2[key].isin(common)].drop_duplicates(subset=[key]).set_index(key)

    diff_rows = []
    for cid in common:
        row1 = df1_common.loc[cid]
        row2 = df2_common.loc[cid]
        diffs = []
        diff_types = []
        for col in cols_compare:
            v1 = str(row1[col]) if not pd.isna(row1[col]) else ''
            v2 = str(row2[col]) if not pd.isna(row2[col]) else ''
            if v1 != v2:
                diffs.append(f"{col}: [{v1}] vs [{v2}]")
                diff_types.append(f'{col}不一致')
        if diffs:
            diff_rows.append({
                key: cid,
                '差异类型': '；'.join(diff_types),
                '差异': '；'.join(diffs),
                '文件1_专区名称': row1.get('专区名称', ''),
                '文件1_预算使用': row1.get('预算使用', ''),
                '文件2_专区名称': row2.get('专区名称', ''),
                '文件2_预算使用': row2.get('预算使用', ''),
            })

    diff_df = pd.DataFrame(diff_rows)

    # 输出
    with pd.ExcelWriter(OUTPUT, engine='openpyxl') as writer:
        only_in_1[[key] + cols_compare + ['来源']].to_excel(writer, sheet_name='仅文件1有', index=False)
        only_in_2[[key] + cols_compare + ['来源']].to_excel(writer, sheet_name='仅文件2有', index=False)
        if not diff_df.empty:
            diff_df.to_excel(writer, sheet_name='数据不一致', index=False)

    print(f"文件1 记录数: {len(df1)}")
    print(f"文件2 记录数: {len(df2)}")
    print(f"共同合同编号: {len(common)} 个")
    print(f"\n仅文件1有: {len(only_in_1)} 个")
    print(f"仅文件2有: {len(only_in_2)} 个")
    print(f"数据不一致: {len(diff_df)} 个")
    print(f"\n导出完成: {OUTPUT}")


if __name__ == '__main__':
    main()
