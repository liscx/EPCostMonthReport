"""
合同汇总表更新工作流

执行顺序：
1. getDQandName70.py - 爬取70专区列表
2. getDQandName80.py - 导出80专区运营统计
3. export_tencent.py - 从腾讯文档导出项目跟进表
4. update_contract_data.py - 更新合同汇总数据
"""
import sys
import os
import importlib
import traceback
from datetime import datetime

sys.stdout.reconfigure(encoding='utf-8')
os.chdir(os.path.dirname(os.path.abspath(__file__)))

WORKFLOW = [
    ('data_export_scripts.getDQandName70', '爬取70专区列表'),
    ('data_export_scripts.getDQandName80', '导出80专区运营统计'),
    ('data_export_scripts.export_tencent', '导出腾讯文档项目跟进表'),
    ('data_export_scripts.update_contract_data', '更新合同汇总数据'),
]


def main():
    ts = datetime.now().strftime('%m%d')
    os.environ['WORKFLOW_TIMESTAMP'] = ts

    print(f"\n{'#'*60}")
    print(f"# 合同汇总表更新工作流")
    print(f"# 时间戳: {ts}")
    print(f"{'#'*60}")

    for idx, (module_name, step_name) in enumerate(WORKFLOW, 1):
        print(f"\n{'='*60}")
        print(f"[{idx}/{len(WORKFLOW)}] {step_name} ({module_name})")
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
    print(f"# 合同汇总表更新工作流完成")
    print(f"{'#'*60}")


if __name__ == '__main__':
    main()
