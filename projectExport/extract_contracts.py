import sys
import pandas as pd
import warnings

sys.stdout.reconfigure(encoding='utf-8')

warnings.filterwarnings('ignore', message='Workbook contains no default style')

SOURCE_FILE = 'ejyExport0101_0331.xlsx'
REF_FILE = '01-新点电子交易专区&项目跟进表（重要）.xlsx'
OUTPUT_FILE = '合同汇总0331.xlsx'

# 按列索引定位 ejyExport.xlsx
COL_CONTRACT = 20   # 合同编号
COL_COST = 12       # 实际人工成本
COL_BUDGET = 23     # 任务预算使用


def load_contracts_from_ejy():
    """从 ejyExport.xlsx 提取C开头的合同编号及汇总数据，返回 {合同编号: {实际人工成本, 任务预算使用}} 字典"""
    xls = pd.ExcelFile(SOURCE_FILE)
    print(f"源文件: {SOURCE_FILE}, 共 {len(xls.sheet_names)} 个Sheet")

    all_data = []
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=0)
        if df.empty or len(df.columns) <= max(COL_CONTRACT, COL_COST, COL_BUDGET):
            print(f"  跳过空Sheet: {sheet_name}")
            continue

        contract = df.iloc[:, COL_CONTRACT]
        cost = pd.to_numeric(df.iloc[:, COL_COST], errors='coerce').fillna(0)
        budget = pd.to_numeric(df.iloc[:, COL_BUDGET], errors='coerce').fillna(0)

        temp = pd.DataFrame({
            '合同编号': contract,
            '实际人工成本': cost,
            '任务预算使用': budget,
        })
        temp = temp[temp['合同编号'].notna() & (temp['合同编号'] != '')]
        temp = temp[temp['合同编号'].astype(str).str.startswith('C')]
        if not temp.empty:
            all_data.append(temp)
            print(f"  {sheet_name}: {len(temp)} 条记录")

    if not all_data:
        print("未找到任何数据")
        return None

    merged = pd.concat(all_data, ignore_index=True)
    agg = merged.groupby('合同编号', as_index=False).agg(
        实际人工成本=('实际人工成本', 'sum'),
        任务预算使用=('任务预算使用', 'sum'),
    )
    agg['实际人工成本'] = agg['实际人工成本'].round(2)
    agg['任务预算使用'] = agg['任务预算使用'].round(2)

    cost_dict = {}
    for _, r in agg.iterrows():
        cost_dict[str(r['合同编号'])] = {
            '实际人工成本': r['实际人工成本'],
            '任务预算使用': r['任务预算使用'],
        }
    print(f"\n去重后合同数: {len(cost_dict)}")
    return cost_dict


def _process_sheet(df, value_cols, cost_dict):
    """处理一个参考sheet：按行拆分合同编号，同单元格的合同合计为一行"""
    grouped_rows = []
    grouped_cids = set()

    for _, row in df.iterrows():
        raw = str(row['ref_contract'])
        cids = [c.strip() for c in raw.split('\n') if c.strip()]
        if not cids:
            continue

        # 在 ejyExport 中查找：待匹配数据被数组元素包含
        matched_cids = []
        total_cost = 0
        total_budget = 0
        for dict_cid in cost_dict:
            for ref_cid in cids:
                if dict_cid in ref_cid:
                    matched_cids.append(dict_cid)
                    total_cost += cost_dict[dict_cid]['实际人工成本']
                    total_budget += cost_dict[dict_cid]['任务预算使用']
                    break

        if matched_cids:
            grouped_cids.update(matched_cids)
            info = {col: str(row[col]) if pd.notna(row[col]) else '' for col in value_cols}
            grouped_rows.append({
                '合同编号': '\n'.join(matched_cids),
                '实际人工成本': round(total_cost, 2),
                '任务预算使用': round(total_budget, 2),
                **info,
            })

    return grouped_rows, grouped_cids


def match_reference(cost_dict):
    """从参考文件匹配：同单元格多合同合计为一行，拼接合同编号"""
    print(f"\n读取参考文件: {REF_FILE}")

    # 专区管控表
    df_zone = pd.read_excel(REF_FILE, sheet_name='专区管控表', header=0, usecols=[0, 2, 5, 11])
    df_zone.columns = ['ref_contract', '专区名称', '省份（已按新组织调整）', '商务']
    df_zone = df_zone.dropna(subset=['ref_contract'])
    df_zone['ref_contract'] = df_zone['ref_contract'].astype(str)
    zone_cols = ['省份（已按新组织调整）', '专区名称', '商务']
    zone_rows, zone_cids = _process_sheet(df_zone, zone_cols, cost_dict)
    print(f"  专区管控表: {len(zone_rows)} 组, 覆盖 {len(zone_cids)} 个合同编号")

    # 落地项目
    df_proj = pd.read_excel(REF_FILE, sheet_name='落地项目', header=0, usecols=[0, 1, 2, 3, 4])
    df_proj.columns = ['ref_contract', '项目名称', '分公司', '地区', '负责人']
    df_proj = df_proj.dropna(subset=['ref_contract'])
    df_proj['ref_contract'] = df_proj['ref_contract'].astype(str)
    proj_cols = ['项目名称', '分公司', '地区', '负责人']
    proj_rows, proj_cids = _process_sheet(df_proj, proj_cols, cost_dict)
    print(f"  落地项目: {len(proj_rows)} 组, 覆盖 {len(proj_cids)} 个合同编号")

    return zone_rows, zone_cids, proj_rows, proj_cids


def main():
    cost_dict = load_contracts_from_ejy()
    if cost_dict is None:
        return

    zone_rows, zone_cids, proj_rows, proj_cids = match_reference(cost_dict)

    # 收集已被参考文件匹配走的合同编号
    covered_cids = zone_cids | proj_cids

    # 未匹配的合同：单独一行
    unmatched_rows = []
    for cid, costs in cost_dict.items():
        if cid not in covered_cids:
            unmatched_rows.append({
                '合同编号': cid,
                '实际人工成本': costs['实际人工成本'],
                '任务预算使用': costs['任务预算使用'],
            })
    if unmatched_rows:
        print(f"\n未匹配参考文件的合同: {len(unmatched_rows)} 个")

    # 合并所有行
    all_rows = zone_rows + proj_rows + unmatched_rows
    result = pd.DataFrame(all_rows)

    result.to_excel(OUTPUT_FILE, index=False, sheet_name='合同汇总')
    print(f"\n导出完成: {OUTPUT_FILE}")
    print(f"总行数: {len(result)}")


if __name__ == '__main__':
    main()
