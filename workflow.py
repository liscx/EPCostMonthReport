import sys
import os
import importlib
import traceback
import yaml
from datetime import datetime, date
import calendar

sys.stdout.reconfigure(encoding='utf-8')
os.chdir(os.path.dirname(os.path.abspath(__file__)))

WORKFLOW = [
    ('data_export_scripts.workflow_data_export', '数据导出'),
    ('data_handle_scripts.workflow_data_handle', '数据处理'),
]


def main():
    ts = datetime.now().strftime('%m%d')
    os.environ['WORKFLOW_TIMESTAMP'] = ts

    # auto_range 检测：自动写回日期到 config.yaml
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.yaml')
    with open(config_path, 'r', encoding='utf-8') as f:
        config = yaml.safe_load(f)
    if config.get('auto_range'):
        today = date.today()
        auto_start = f"{today.year}-01-01"
        prev_month = today.month - 1 if today.month > 1 else 12
        prev_year = today.year if today.month > 1 else today.year - 1
        last_day = calendar.monthrange(prev_year, prev_month)[1]
        auto_end = f"{prev_year}-{prev_month:02d}-{last_day:02d}"
        config['date_range'] = {'start_date': auto_start, 'end_date': auto_end}
        with open(config_path, 'w', encoding='utf-8') as f:
            yaml.dump(config, f, allow_unicode=True, default_flow_style=False)
        print(f"[auto_range] 日期已自动更新: {auto_start} ~ {auto_end}")

    print(f"\n{'#'*60}")
    print(f"# 完整月报工作流")
    print(f"# 时间戳: {ts}")
    print(f"{'#'*60}")

    for idx, (module_name, step_name) in enumerate(WORKFLOW, 1):
        print(f"\n{'='*60}")
        print(f"[{idx}/{len(WORKFLOW)}] {step_name} ({module_name}.py)")
        print(f"{'='*60}")

        try:
            mod = importlib.import_module(module_name)
            mod.main()
            print(f"\n[OK] {step_name} 完成")
        except Exception as e:
            print(f"\n[FAIL] {step_name} 失败: {e}")
            traceback.print_exc()
            sys.exit(1)

    # 输出最终文件路径
    with open('config.yaml', 'r', encoding='utf-8') as f:
        config = yaml.safe_load(f)
    start_str = config['date_range']['start_date'].replace('-', '')
    end_str = config['date_range']['end_date'].replace('-', '')

    ts = os.environ.get('WORKFLOW_TIMESTAMP', '')
    suffix = f'_{ts}' if ts else ''
    report_file = f'结果数据/SaaS月报{start_str}-{end_str}{suffix}.docx'

    print(f"\n{'#'*60}")
    print(f"# 完整月报工作流完成")
    print(f"{'#'*60}")
    print(f"REPORT_FILE:{os.path.abspath(report_file)}")


if __name__ == '__main__':
    main()
