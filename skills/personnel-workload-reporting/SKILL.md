---
name: personnel-workload-reporting
description: Generate and revise personnel workload survey reports for机关/职能部门摸排, including supervisor 20-question recalculation, ERP usage, workload, certificates, feedback summaries, conclusion analysis, and comprehensive reports. Use when producing or modifying these reports from Excel/Markdown/PDF inputs in this repository.
---

# Personnel Workload Reporting

Use this skill when generating, checking, or modifying the机关工作量摸排系列报告. The reports must be generated from the current folder's Excel/Markdown/image inputs and should preserve the established report structure, data口径, and PDF formatting rules.

## Core Workflow

1. Generate `机关员工主管评价20题重算报告` first.
2. Generate the branch reports:
   - `机关员工ERP使用情况分析`
   - `机关员工工作负荷报告`
   - `机关员工证书客观维度分析`
   - `机关工作量摸排报告`
   - `机关工作量摸排结论分析`
   - `有想法人员及妈妈岗意向名单`
   - `重点问题反馈汇总`
   - `重点问题反馈结论分析`
3. Generate `机关工作量摸排综合报告` last.
4. For every generated Markdown report, also generate the corresponding PDF.
5. Do not overwrite old reports; output files keep the run date in the filename.

## Global Rules

- Current scope may be single department, multiple departments, or all departments. All reports must use the same scope wording.
- All `按20题计算等级` / `20题重算评级` / `综合重算评级` values come from the supervisor 20-question recalculation report.
- Do not use `主管评级对应公司等级` for analysis.
- ERP operation data uses `ERP使用情况.xlsx` sheet `2025年全年`, not `近三年`.
- If supervisor evaluation sample count is 0, generate reports without crashing; do not show empty rating matrices or `当前0名...` phrasing.
- If a statistic is 0, write `当前未发现` or `暂不涉及`; do not continue with fixed-language lists such as `其中` or `主要包括`.
- PDF tables use light-blue headers, comfortable row height, and adjusted column widths for long text.

## Reference Files

Load only the reference needed for the report being generated or revised:

| Task | Reference |
| --- | --- |
| Supervisor 20-question recalculation | `references/机关员工主管评价20题重算报告规则.md` |
| ERP usage analysis | `references/机关员工ERP使用情况分析规则.md` |
| Workload analysis | `references/机关员工工作负荷报告规则.md` |
| Certificate objective dimension analysis | `references/机关员工证书客观维度分析规则.md` |
| Workload survey main report | `references/机关工作量摸排报告规则.md` |
| Workload conclusion analysis | `references/机关工作量摸排结论分析规则.md` |
| Idea and mommy-post intention list | `references/有想法人员及妈妈岗意向名单规则.md` |
| Key issue feedback summary | `references/重点问题反馈汇总规则.md` |
| Key issue feedback conclusion | `references/重点问题反馈结论分析规则.md` |
| Comprehensive report | `references/机关工作量摸排综合报告规则.md` |

## Validation Checklist

Before finishing:

- Confirm all required reports for the requested scope were generated.
- Confirm MD and PDF versions both exist.
- Confirm ERP text says `2025年全年`.
- Confirm supervisor recalculation is generated before dependent reports.
- Confirm no contradictory zero-person language appears.
- Confirm current-scope wording is consistent.
- Confirm Git changes only include intended skill/report files when committing.
