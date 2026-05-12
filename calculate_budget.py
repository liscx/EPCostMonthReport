import pandas as pd
import sys
import yaml
import os

sys.stdout.reconfigure(encoding='utf-8')

# 读取配置
with open('config.yaml', 'r', encoding='utf-8') as f:
    config = yaml.safe_load(f)

start_date = config['date_range']['start_date']
end_date = config['date_range']['end_date']
start_str = start_date.replace('-', '')
end_str = end_date.replace('-', '')
# 读取数据
input_file = f'中间数据/合同汇总{start_str}-{end_str}_fix.xlsx'
output_file =f'中间数据/合同汇总{start_str}-{end_str}_核定总额计算结果.xlsx'
df = pd.read_excel(input_file)

def calc_quota(row):
    contract_id = str(row['合同编号'])
    revenue = row['25年收益']

    year = int(contract_id[1:5])

    # 1. 2026年：新开专区，固定50000
    if year == 2026:
        return 50000, '2026年新开专区，核定总额=50000'

    # 2. 2025年10月之后，收益为空或为0，按50000算
    if year == 2025:
        month = int(contract_id[5:7])
        if month >= 10 and (pd.isna(revenue) or revenue == 0):
            return 50000, f'2025年{month}月开设，收益为空或0，新开专区，核定总额=50000'

    # 3. 2025年10月之前收益为空或为0，保底32000
    if year == 2025 and month < 10 and (pd.isna(revenue) or revenue == 0):
        return 32000, f'2025年{month}月开设，收益为空或0，保底32000'

    # 4. 收益为空或为0（2025年之前），保底32000
    if pd.isna(revenue) or revenue == 0:
        return 32000, '收益为空或0，保底32000'

    # 5. 2025年合同（10月之前，收益不为0）：年化收益×0.3，最低32000
    if year == 2025:
        months = 12 - month
        if months <= 0:
            quota = revenue * 0.3
            return max(quota, 32000), f'2025年{month}月开设，收益×0.3={quota:.0f}'
        annual_revenue = revenue / months * 12
        quota = annual_revenue * 0.3
        if quota >= 32000:
            return quota, f'2025年{month}月开设，{months}个月收益，年化={annual_revenue:.0f}，×0.3={quota:.0f}'
        else:
            return 32000, f'2025年{month}月开设，{months}个月收益，年化={annual_revenue:.0f}，×0.3={quota:.0f}<32000，保底32000'

    # 6. 2025年之前合同：收益×0.3，最低32000
    quota = revenue * 0.3
    if quota >= 32000:
        return quota, f'2025年前合同，收益×0.3={quota:.0f}'
    else:
        return 32000, f'2025年前合同，收益×0.3={quota:.0f}<32000，保底32000'



# 只计算核定总额为空的数据
mask = df['核定总额'].isna() | (df['核定总额'] == '')

# 记录原始核定总额有值的数量
existing_count = len(df) - mask.sum()

# 计算核定总额和计算规则（仅对空值计算）
if mask.any():
    df.loc[mask, ['核定总额', '核定总额计算规则']] = df[mask].apply(calc_quota, axis=1, result_type='expand')

# 保存结果
df.to_excel(output_file, index=False)

print(f'计算完成，结果已保存到: {output_file}')
print(f'\n统计信息:')
print(f'总记录数: {len(df)}')
print(f'已有核定总额（跳过）: {existing_count}条')
print(f'本次计算核定总额: {mask.sum()}条')
print(f'核定总额为50000的(2026年新开): {len(df[df["核定总额"] == 50000])}条')
print(f'核定总额为32000的(保底): {len(df[df["核定总额"] == 32000])}条')
print(f'其他核定总额: {len(df[(df["核定总额"] != 32000) & (df["核定总额"] != 50000)])}条')

