import sys
import os
import importlib
import traceback
from datetime import datetime

sys.stdout.reconfigure(encoding='utf-8')
os.chdir(os.path.dirname(os.path.abspath(__file__)))

WORKFLOW = [
    ('data_export', '数据导出'),
    ('extract_contracts', '合同匹配'),
    ('calculate_budget', '计算核定总额'),
]


def main():
    ts = datetime.now().strftime('%m%d')
    os.environ['WORKFLOW_TIMESTAMP'] = ts

    print(f"\n{'#'*60}")
    print(f"# 数据处理工作流")
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

    print(f"\n{'#'*60}")
    print(f"# 数据处理工作流完成")
    print(f"{'#'*60}")


if __name__ == '__main__':
    main()
