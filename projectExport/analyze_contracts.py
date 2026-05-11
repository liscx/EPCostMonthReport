import sys
import pandas as pd
import warnings

sys.stdout.reconfigure(encoding='utf-8')
warnings.filterwarnings('ignore', message='Workbook contains no default style')

INPUT_FILE = '合同汇总.xlsx'
OUTPUT_FILE = '合同汇总20260101-20260430.xlsx'


def main():
    df = pd.read_excel(INPUT_FILE)
    print(f"读取: {INPUT_FILE}, 共 {len(df)} 条记录")

    is_new = df['合同编号'].astype(str).str.startswith('C2026')
    df['专属属性'] = is_new.map({True: '新开', False: '历史'})
    df['核定总额'] = is_new.map({True: 50000, False: None})

    new_count = is_new.sum()
    print(f"新开: {new_count} 个, 历史: {len(df) - new_count} 个")

    df.to_excel(OUTPUT_FILE, index=False, sheet_name='合同汇总')
    print(f"导出完成: {OUTPUT_FILE}")


if __name__ == '__main__':
    main()
