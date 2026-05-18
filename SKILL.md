# SaaS 月报自动化项目

## 项目概述

自动化生成 SaaS 运营月报，包含数据导出、合同匹配、核定总额计算、收益导出回填、合同分析和报告生成。

## 工作流

本项目包含三种运行方式：

### 完整工作流（推荐）

```bash
python workflow.py
```

执行顺序：
1. `data_export.py` — 从 OA 系统导出 projExport 数据（Selenium，需登录）
2. `extract_contracts.py` — 合同匹配，生成合同汇总表
3. `calculate_budget.py` — 计算核定总额
4. `shouyi_export.py` — 从专区收益统计页面导出收益数据（Playwright，需登录）
5. `fill_revenue.py` — 根据辖区code匹配专区码，回填26年收益
6. `analyze_contracts.py` — 合同分析，计算预算使用率、收入成本比、颜色分类
7. `generate_report.py` — 基于 Word 模板生成月报文档

### 工作流一：数据处理

```bash
python workflow_data.py
```

执行步骤 1-3，仅处理 projExport 数据。

### 工作流二：月报生成

```bash
python workflow_report.py
```

执行步骤 4-7，需要流程一的产出文件。

## 配置文件

`config.yaml` 控制所有参数：

```yaml
date_range:
  start_date: "2026-01-01"   # 数据起始日期
  end_date: "2026-04-30"     # 数据结束日期

# 收益列名，留空则自动根据当前月份生成（如5月→1-4月收益）
dataRange: ""

# 数据年份简称，留空则自动根据当前年份生成（如2026→26）
dataYear: ""

split_export:
  segment_months: 2           # 分段导出的月数
  projects:                   # 需要分段导出的项目
    - "新点电子交易平台"
    - "江苏省限额以下电子交易平台"
```

## 文件结构

```
源数据/
  export/
    projExport{start}-{end}.xlsx   ← data_export 产出
    shouyi{start}-{end}.xlsx       ← shouyi_export 产出
  合同汇总表.xlsx                  ← 手动维护的参考数据
  新点电子交易专区&项目跟进表.xlsx  ← 旧参考数据（专区码映射）

中间数据/
  合同汇总{start}-{end}_{ts}.xlsx  ← 各步骤共享的中间文件
  统计数据{start}-{end}_{ts}.json  ← analyze_contracts 产出

结果数据/
  SaaS月报{start}-{end}_{ts}.docx ← 最终月报

Model/
  SaaSReportModel.docx             ← Word 月报模板

projList.xlsx                      ← 项目列表（名称+URL）
config.yaml                        ← 全局配置
```

## 时间戳机制

workflow 运行时自动生成时间戳（月+日，如 `0513`），通过环境变量 `WORKFLOW_TIMESTAMP` 传递给各脚本。中间文件统一使用 `_时间戳` 后缀，避免覆盖历史数据。

单独运行各脚本时不受影响，使用默认文件名。

## 关键业务规则

- 合同编号格式：`C{年份4位}{月份2位}{序号4位}`，如 `C2026050001`
- 颜色分类基于预算使用率和收入成本比自动判定
- 核定总额仅对空值计算，已有值保留不动
- 预算使用率始终在 analyze_contracts.py 中计算（预算使用 ÷ 核定总额）
- 收益回填通过辖区code匹配专区码，将实得收益写入26年收益列
- shouyi_export 使用 Playwright，与 data_export 的 Selenium 互不影响

## 依赖

- Python 3.12+
- pandas, openpyxl, pyyaml, python-docx, selenium, playwright
