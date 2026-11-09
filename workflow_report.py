import sys
import os
import importlib
import traceback
from datetime import datetime

sys.stdout.reconfigure(encoding='utf-8')
_PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
os.chdir(_PROJECT_ROOT)

from log_utils import RunLogger

WORKFLOW = [
    # ('data_export_scripts.shouyi_export', '收益导出'),
    ('data_handle_scripts.fill_revenue', '收益回填'),
    ('data_handle_scripts.analyze_contracts', '合同分析'),
    ('data_handle_scripts.generate_report', '生成月报'),
]


def main():
    ts = datetime.now().strftime('%m%d')
    os.environ['WORKFLOW_TIMESTAMP'] = ts

    rl = RunLogger(title='月报生成工作流')
    print(f"\n{'#'*60}")
    print(f"# 月报生成工作流")
    print(f"# 时间戳: {ts}")
    print(f"# 日志目录: {rl.run_dir}")
    print(f"{'#'*60}")

    failed = False
    for idx, (module_name, step_name) in enumerate(WORKFLOW, 1):
        with rl.step(step_name) as st:
            print(f"\n{'='*60}")
            print(f"[{idx}/{len(WORKFLOW)}] {step_name} ({module_name}.py)")
            print(f"{'='*60}")

            try:
                mod = importlib.import_module(module_name)
                mod.main()
                print(f"\n[OK] {step_name} 完成")
                st.ok()
            except Exception as e:
                print(f"\n[FAIL] {step_name} 失败: {e}")
                traceback.print_exc()
                st.fail(e)
                failed = True

        if failed:
            print(f"\n[FAIL] 工作流中断于第 {idx} 步: {step_name}")
            rl.finish('失败')
            sys.exit(1)

    print(f"\n{'#'*60}")
    print(f"# 月报生成工作流完成")
    print(f"# 日志目录: {rl.run_dir}")
    print(f"{'#'*60}")
    rl.finish('完成')


if __name__ == '__main__':
    main()
