import sys
import pandas as pd
import warnings
import yaml
import os
# 2025年数据匹配
sys.stdout.reconfigure(encoding='utf-8')
warnings.filterwarnings('ignore', message='Workbook contains no default style')

# 读取配置
with open('config.yaml', 'r', encoding='utf-8') as f:
    config = yaml.safe_load(f)

start_date = config['date_range']['start_date']
end_date = config['date_range']['end_date']
start_str = start_date.replace('-', '')
end_str = end_date.replace('-', '')

# 输入输出文件配置
INPUT_FILE = f'中间数据/合同汇总{start_str}-{end_str}.xlsx'
OUTPUT_FILE = f'中间数据/合同汇总{start_str}-{end_str}.xlsx'


def main():
    df = pd.read_excel(INPUT_FILE)
    print(f"读取: {INPUT_FILE}, 共 {len(df)} 条记录")

    is_new = df['合同编号'].astype(str).str.startswith('C2026')
    df['专属属性'] = is_new.map({True: '新开', False: '历史'})

    new_count = is_new.sum()
    print(f"新开: {new_count} 个, 历史: {len(df) - new_count} 个")

    df.to_excel(OUTPUT_FILE, index=False, sheet_name='合同汇总')
    print(f"导出完成: {OUTPUT_FILE}")


if __name__ == '__main__':
    main()
