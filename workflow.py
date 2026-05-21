import sys
import os
import importlib
import traceback
import yaml
from datetime import datetime

sys.stdout.reconfigure(encoding='utf-8')
os.chdir(os.path.dirname(os.path.abspath(__file__)))

WORKFLOW = [
    ('data_export', '数据导出'),
    ('shouyi_export', '收益导出'),
    ('extract_contracts', '合同匹配'),
    ('calculate_budget', '计算核定总额'),
    ('fill_revenue', '收益回填'),
    ('analyze_contracts', '合同分析'),
    ('generate_report', '生成月报'),
]


def main():
    ts = datetime.now().strftime('%m%d')
    os.environ['WORKFLOW_TIMESTAMP'] = ts

    print(f"\n{'#'*60}")
    print(f"# 完整月报工作流")
    print(f"# 时间戳: {ts}")
    print(f"{'#'*60}")

    driver = None
    for idx, (module_name, step_name) in enumerate(WORKFLOW, 1):
        print(f"\n{'='*60}")
        print(f"[{idx}/{len(WORKFLOW)}] {step_name} ({module_name}.py)")
        print(f"{'='*60}")

        try:
            mod = importlib.import_module(module_name)
            # data_export 返回 driver，传递给 shouyi_export 复用浏览器
            if module_name == 'data_export':
                driver = mod.main()
                print(f"\n[OK] {step_name} 完成")
            elif module_name == 'shouyi_export' and driver:
                # 使用相同的浏览器实例，不需要重新登录
                mod.main(debug_port=9222)
                print(f"\n[OK] {step_name} 完成")
            else:
                mod.main()
                print(f"\n[OK] {step_name} 完成")
        except Exception as e:
            print(f"\n[FAIL] {step_name} 失败: {e}")
            traceback.print_exc()
            # 关闭浏览器后再退出
            if driver:
                try:
                    driver.quit()
                except:
                    pass
            sys.exit(1)

    # 关闭浏览器
    if driver:
        try:
            print("\n关闭浏览器...")
            driver.quit()
        except:
            pass

    # 输出最终文件路径
    with open('config.yaml', 'r', encoding='utf-8') as f:
        config = yaml.safe_load(f)
    start_str = config['date_range']['start_date'].replace('-', '')
    end_str = config['date_range']['end_date'].replace('-', '')

    source_file = f'中间数据/合同汇总{start_str}-{end_str}_{ts}.xlsx'
    report_file = f'结果数据/SaaS月报{start_str}-{end_str}_{ts}.docx'

    print(f"\n{'#'*60}")
    print(f"# 完整月报工作流完成")
    print(f"{'#'*60}")
    print(f"SOURCE_FILE:{os.path.abspath(source_file)}")
    print(f"REPORT_FILE:{os.path.abspath(report_file)}")


if __name__ == '__main__':
    main()
