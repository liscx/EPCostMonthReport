import sys
import pandas as pd
import warnings
import yaml
import os
import json
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
INPUT_FILE = f'中间数据/合同汇总{start_str}-{end_str}_fixx.xlsx'
OUTPUT_FILE = f'中间数据/合同汇总{start_str}-{end_str}_fixx.xlsx'


def main():
    df = pd.read_excel(INPUT_FILE)
    print(f"读取: {INPUT_FILE}, 共 {len(df)} 条记录")

    is_new = df['合同编号'].astype(str).str.startswith('C2026')
    df['专属属性'] = is_new.map({True: '新开', False: '历史'})

    new_count = is_new.sum()
    print(f"新开: {new_count} 个, 历史: {len(df) - new_count} 个")

    # 预算使用率已经是百分比格式，直接使用原始数据
    # 确保输出时带有百分号
    if df['预算使用率'].dtype == 'float64':
        # 如果是小数形式，转换为百分比格式（乘以100再加%）
        df['预算使用率'] = df['预算使用率'].apply(lambda x: f"{round(x * 100, 2)}%" if pd.notna(x) else "0%")
    else:
        # 如果已经是字符串形式，确保有百分号
        df['预算使用率'] = df['预算使用率'].apply(
            lambda x: str(x) if pd.notna(x) and str(x).endswith('%') else (f"{x}%" if pd.notna(x) else "0%")
        )

    # 计算收入成本比 = 实际运营收益 ÷ 已使用预算
    # 实际运营收益 = 1-4月收益列（J列）, 已使用预算 = 预算使用列
    # J列为空时按0计算
    df['1-4月收益'] = df['1-4月收益'].fillna(0)
    df['收入成本比'] = df.apply(
        lambda row: round(row['1-4月收益'] / row['预算使用'], 2) if row['预算使用'] > 0 else 0,
        axis=1
    )

    print(f"预算使用率计算完成")
    print(f"收入成本比计算完成")

    # 分类函数：判断每行属于哪个颜色
    def classify_row(row):
        # 将百分比字符串转换为数值
        budget_rate_str = row['预算使用率']
        budget_rate = float(budget_rate_str.replace('%', ''))
        cost_ratio = row['收入成本比']
        revenue_1_4 = row['1-4月收益'] if pd.notna(row['1-4月收益']) else 0
        revenue_25 = row['25年收益'] if pd.notna(row['25年收益']) else 0

        # 绿色（良性运营）：收入成本比≥1.72
        if cost_ratio >= 1.72:
            return '绿色'

        # 红色（严重问题）
        # 预算使用率≥100% 且 实际收益 = 0
        if budget_rate >= 100 and revenue_1_4 == 0:
            return '红色'

        # 橙色（需改进）
        # 1. 50%≤预算使用率＜100% 且 收入成本比＜1
        if 50 <= budget_rate < 100 and cost_ratio < 1:
            return '橙色'
        # 2. 预算使用率≥100%、实际收益＞0且收入成本比＜1
        if budget_rate >= 100 and revenue_1_4 > 0 and cost_ratio < 1:
            return '橙色'

        # 黄色（持续关注）：30%≤预算使用率＜50% 且 实际收益 = 0
        if 30 <= budget_rate < 50 and revenue_1_4 == 0:
            return '黄色'

        # 不符合条件的返回空
        return ''

    # 添加颜色分类列
    df['颜色分类'] = df.apply(classify_row, axis=1)

    # 统计各颜色数量
    color_counts = df['颜色分类'].value_counts()
    print(f"\n颜色分类统计:")
    for color in ['红色', '橙色', '黄色', '绿色']:
        count = color_counts.get(color, 0)
        print(f"  {color}: {count} 个")
    print(f"  不符合条件: {len(df[df['颜色分类'] == ''])} 个")

    # 输出到Excel，每个颜色一个sheet
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        # 总表
        df.to_excel(writer, index=False, sheet_name='合同汇总')

        # 各颜色分类sheet
        for color in ['红色', '橙色', '黄色', '绿色']:
            color_df = df[df['颜色分类'] == color]
            if not color_df.empty:
                color_df.to_excel(writer, index=False, sheet_name=f'{color}专区')

    print(f"\n导出完成: {OUTPUT_FILE}")
    print(f"包含sheet: 合同汇总, 红色专区, 橙色专区, 黄色专区, 绿色专区")

    # 计算统计数据并保存到JSON
    stats = {
        "预算使用总额": float(round(df['预算使用'].sum(), 2)),
        "实际运营收益": float(round(df['1-4月收益'].sum(), 2)),
        "有成本投入专区总数": int(df['预算使用'].count()),
        "红色专区个数": int(color_counts.get('红色', 0)),
        "橙色专区个数": int(color_counts.get('橙色', 0)),
        "黄色专区个数": int(color_counts.get('黄色', 0)),
        "绿色专区个数": int(color_counts.get('绿色', 0))
    }

    json_file = f'中间数据/统计数据{start_str}-{end_str}.json'
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(stats, f, ensure_ascii=False, indent=2)

    print(f"\n统计数据已保存到: {json_file}")
    for key, value in stats.items():
        print(f"  {key}: {value}")


if __name__ == '__main__':
    main()
