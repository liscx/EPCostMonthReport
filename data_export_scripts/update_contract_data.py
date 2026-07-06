"""
合同汇总数据更新脚本

功能：
1. 根据专区码匹配"新点电子交易专区&项目跟进表.xlsx"的"专区管控表"sheet，更新"商务"列
2. 根据专区码匹配"70DQ.xlsx"和"80DQ.xlsx"，更新"专区名称"列
3. 生成txt日志文件记录更新的数据
"""
import os
import sys
import glob
import pandas as pd
from datetime import datetime

sys.stdout.reconfigure(encoding='utf-8')

# 路径配置
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SOURCE_DIR = os.path.join(BASE_DIR, '源数据')
TEMP_DIR = os.path.join(BASE_DIR, '中间数据')


def get_latest_contract_summary():
    """获取合同汇总源数据文件"""
    contract_file = os.path.join(SOURCE_DIR, '合同汇总表.xlsx')
    if os.path.exists(contract_file):
        return contract_file
    return None


def update_contract_summary():
    """更新合同汇总源数据"""
    today = datetime.now().strftime("%Y%m%d")
    log_lines = []
    log_lines.append(f"更新时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log_lines.append("=" * 60)

    # 1. 找到最新的合同汇总源数据
    contract_file = get_latest_contract_summary()
    if not contract_file:
        log_lines.append("[ERROR] 未找到合同汇总源数据文件")
        return log_lines

    log_lines.append(f"合同汇总文件: {contract_file}")
    df_contract = pd.read_excel(contract_file)
    log_lines.append(f"原始记录数: {len(df_contract)}")
    log_lines.append("")

    # 2. 读取"新点电子交易专区&项目跟进表.xlsx"的"专区管控表"sheet
    ref_file = os.path.join(SOURCE_DIR, '新点电子交易专区&项目跟进表.xlsx')
    if os.path.exists(ref_file):
        log_lines.append(f"读取参考文件: {ref_file}")
        try:
            df_ref = pd.read_excel(ref_file, sheet_name='专区管控表', header=0)
            ref_cols = df_ref.columns.tolist()
            log_lines.append(f"  专区管控表列名: {ref_cols}")

            # 查找专区码和商务列
            zone_code_col = None
            business_col = None
            for col in ref_cols:
                col_str = str(col)
                if '专区码' in col_str or '专区编号' in col_str:
                    zone_code_col = col
                if '商务' in col_str:
                    business_col = col

            if zone_code_col and business_col:
                # 构建专区码 -> 商务映射
                ref_map = {}
                for _, row in df_ref.iterrows():
                    code = str(row[zone_code_col]).strip() if pd.notna(row[zone_code_col]) else ''
                    business = str(row[business_col]).strip() if pd.notna(row[business_col]) else ''
                    if code and code != 'nan':
                        ref_map[code] = business

                log_lines.append(f"  专区码映射数量: {len(ref_map)}")

                # 更新合同汇总的商务列
                if '专区码' in df_contract.columns and '商务' in df_contract.columns:
                    updated_count = 0
                    for idx, row in df_contract.iterrows():
                        zone_code = str(row['专区码']).strip() if pd.notna(row['专区码']) else ''
                        if zone_code in ref_map and ref_map[zone_code]:
                            old_val = str(row['商务']) if pd.notna(row['商务']) else ''
                            new_val = ref_map[zone_code]
                            if old_val != new_val:
                                df_contract.at[idx, '商务'] = new_val
                                updated_count += 1
                                log_lines.append(f"  [更新商务] 行{idx+2} 专区码={zone_code}: {old_val} -> {new_val}")

                    log_lines.append(f"  商务列更新数量: {updated_count}")
                else:
                    log_lines.append(f"  [WARN] 合同汇总缺少'专区码'或'商务'列")
            else:
                log_lines.append(f"  [WARN] 专区管控表未找到专区码或商务列")

        except Exception as e:
            log_lines.append(f"  [ERROR] 读取参考文件失败: {e}")
    else:
        log_lines.append(f"[WARN] 参考文件不存在: {ref_file}")

    log_lines.append("")

    # 3. 读取70DQ文件，更新专区名称
    dq70_file = os.path.join(TEMP_DIR, '70DQ.xlsx')
    if os.path.exists(dq70_file):
        log_lines.append(f"读取70DQ文件: {dq70_file}")
        try:
            df_70 = pd.read_excel(dq70_file)
            dq70_map = {}
            for _, row in df_70.iterrows():
                code = str(row['专区编号']).strip() if pd.notna(row['专区编号']) else ''
                name = str(row['专区名称']).strip() if pd.notna(row['专区名称']) else ''
                if code and code != 'nan':
                    dq70_map[code] = name

            log_lines.append(f"  70DQ专区数量: {len(dq70_map)}")

            if '专区码' in df_contract.columns and '专区名称' in df_contract.columns:
                updated_count = 0
                for idx, row in df_contract.iterrows():
                    zone_code = str(row['专区码']).strip() if pd.notna(row['专区码']) else ''
                    if zone_code in dq70_map and dq70_map[zone_code]:
                        old_val = str(row['专区名称']) if pd.notna(row['专区名称']) else ''
                        new_val = dq70_map[zone_code]
                        if old_val != new_val:
                            df_contract.at[idx, '专区名称'] = new_val
                            updated_count += 1
                            log_lines.append(f"  [更新专区名称-70] 行{idx+2} 专区码={zone_code}: {old_val} -> {new_val}")

                log_lines.append(f"  70DQ专区名称更新数量: {updated_count}")
            else:
                log_lines.append(f"  [WARN] 合同汇总缺少'专区码'或'专区名称'列")

        except Exception as e:
            log_lines.append(f"  [ERROR] 读取70DQ文件失败: {e}")
    else:
        log_lines.append(f"[WARN] 70DQ文件不存在: {dq70_file}")

    log_lines.append("")

    # 4. 读取80DQ文件，更新专区名称
    dq80_file = os.path.join(TEMP_DIR, '80DQ.xlsx')
    if os.path.exists(dq80_file):
        log_lines.append(f"读取80DQ文件: {dq80_file}")
        try:
            df_80 = pd.read_excel(dq80_file)
            dq80_map = {}
            for _, row in df_80.iterrows():
                code = str(row['专区标识']).strip() if pd.notna(row['专区标识']) else ''
                name = str(row['专区名称']).strip() if pd.notna(row['专区名称']) else ''
                if code and code != 'nan':
                    dq80_map[code] = name

            log_lines.append(f"  80DQ专区数量: {len(dq80_map)}")

            if '专区码' in df_contract.columns and '专区名称' in df_contract.columns:
                updated_count = 0
                for idx, row in df_contract.iterrows():
                    zone_code = str(row['专区码']).strip() if pd.notna(row['专区码']) else ''
                    if zone_code in dq80_map and dq80_map[zone_code]:
                        old_val = str(row['专区名称']) if pd.notna(row['专区名称']) else ''
                        new_val = dq80_map[zone_code]
                        if old_val != new_val:
                            df_contract.at[idx, '专区名称'] = new_val
                            updated_count += 1
                            log_lines.append(f"  [更新专区名称-80] 行{idx+2} 专区码={zone_code}: {old_val} -> {new_val}")

                log_lines.append(f"  80DQ专区名称更新数量: {updated_count}")
            else:
                log_lines.append(f"  [WARN] 合同汇总缺少'专区码'或'专区名称'列")

        except Exception as e:
            log_lines.append(f"  [ERROR] 读取80DQ文件失败: {e}")
    else:
        log_lines.append(f"[WARN] 80DQ文件不存在: {dq80_file}")

    log_lines.append("")

    # 5. 保存更新后的合同汇总文件（覆盖固定名文件）
    output_file = os.path.join(SOURCE_DIR, '合同汇总表.xlsx')
    df_contract.to_excel(output_file, index=False, sheet_name='合同汇总')
    log_lines.append(f"[OK] 更新后的合同汇总已保存: {output_file}")
    log_lines.append(f"最终记录数: {len(df_contract)}")

    return log_lines


def save_log(log_lines):
    """保存日志到txt文件"""
    today = datetime.now().strftime("%Y%m%d")
    log_dir = os.path.join(BASE_DIR, '日志')
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, f'合同汇总更新日志_{today}.txt')

    with open(log_file, 'w', encoding='utf-8') as f:
        f.write('\n'.join(log_lines))

    print(f"\n日志已保存: {log_file}")
    return log_file


def main():
    """主函数"""
    print("=" * 60)
    print("更新合同汇总源数据")
    print("=" * 60)

    log_lines = update_contract_summary()

    print("\n" + "=" * 60)
    print("保存更新日志")
    print("=" * 60)
    save_log(log_lines)

    # 打印日志内容
    print("\n" + "-" * 60)
    print("更新日志:")
    print("-" * 60)
    for line in log_lines:
        print(line)


if __name__ == "__main__":
    main()
