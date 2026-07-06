"""
打包月报脚本

将依赖数据、中间数据、结果数据打包为zip文件
"""
import os
import sys
import yaml
import glob
import zipfile
from datetime import datetime

sys.stdout.reconfigure(encoding='utf-8')

# 路径配置
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def main():
    # 读取配置
    with open(os.path.join(BASE_DIR, 'config.yaml'), 'r', encoding='utf-8') as f:
        config = yaml.safe_load(f)

    start_date = config['date_range']['start_date']
    end_date = config['date_range']['end_date']
    start_str = start_date.replace('-', '')
    end_str = end_date.replace('-', '')

    # 月报月：取结束日期的月份
    report_month = int(end_date.split('-')[1])
    # 当前时间戳
    timestamp = datetime.now().strftime('%Y%m%d')
    # 报告年份
    report_year = int(end_date.split('-')[0])

    # workflow时间戳
    ts = os.environ.get('WORKFLOW_TIMESTAMP', '')
    suffix = f'_{ts}' if ts else ''

    # zip文件名
    zip_name = f'{report_year}年{report_month}月报_{timestamp}.zip'
    zip_path = os.path.join(BASE_DIR, '结果数据', zip_name)

    # 要打包的文件列表（全部放在zip根目录，不分文件夹）
    all_files = []

    # 依赖数据
    all_files.append(os.path.join(BASE_DIR, '源数据', '合同汇总表.xlsx'))
    all_files.extend(glob.glob(os.path.join(BASE_DIR, '源数据', 'export', f'projExport{start_str}-{end_str}*.xlsx')))
    all_files.extend(glob.glob(os.path.join(BASE_DIR, '源数据', 'export', f'shouyi{start_str}-{end_str}*.xlsx')))

    # 中间数据
    all_files.append(os.path.join(BASE_DIR, '中间数据', '70DQ.xlsx'))
    all_files.append(os.path.join(BASE_DIR, '中间数据', '80DQ.xlsx'))
    all_files.extend(glob.glob(os.path.join(BASE_DIR, '中间数据', f'合同汇总{start_str}-{end_str}{suffix}*.xlsx')))
    all_files.extend(glob.glob(os.path.join(BASE_DIR, '中间数据', f'统计数据{start_str}-{end_str}{suffix}*.json')))

    # 结果数据
    all_files.extend(glob.glob(os.path.join(BASE_DIR, '结果数据', f'SaaS月报{start_str}-{end_str}{suffix}*.docx')))

    # 打包（不保留目录结构，所有文件放根目录）
    os.makedirs(os.path.dirname(zip_path), exist_ok=True)

    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for file_path in all_files:
            if os.path.exists(file_path):
                arcname = os.path.basename(file_path)
                zf.write(file_path, arcname)
                print(f'  添加: {arcname}')
            else:
                print(f'  跳过（不存在）: {os.path.basename(file_path)}')

    print(f'\n打包完成: {zip_path}')
    print(f'文件大小: {os.path.getsize(zip_path) / 1024:.1f} KB')


if __name__ == '__main__':
    main()
