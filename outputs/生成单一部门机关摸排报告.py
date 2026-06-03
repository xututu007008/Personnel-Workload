from __future__ import annotations

import html
import math
import os
import re
import sys
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

import pandas as pd

BUNDLED_PYTHON_PACKAGES = Path.home() / ".cache/codex-runtimes/codex-primary-runtime/dependencies/python"
if BUNDLED_PYTHON_PACKAGES.exists() and str(BUNDLED_PYTHON_PACKAGES) not in sys.path:
    sys.path.append(str(BUNDLED_PYTHON_PACKAGES))
BUNDLED_SITE_PACKAGES = BUNDLED_PYTHON_PACKAGES / "lib/python3.12/site-packages"
if BUNDLED_SITE_PACKAGES.exists() and str(BUNDLED_SITE_PACKAGES) not in sys.path:
    sys.path.append(str(BUNDLED_SITE_PACKAGES))

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.platypus import Image as RLImage
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


ROOT = Path(__file__).resolve().parent
os.environ.setdefault("MPLCONFIGDIR", str(ROOT / "outputs" / ".matplotlib"))
DATE = datetime.now().strftime("%Y-%m-%d")
DEFAULT_DEPTS = [
    "HSE部",
    "人力资源部",
    "信息化中心",
    "南京永利重工制造有限公司",
    "合同管理部",
    "审计部",
    "技术中心",
    "技术质量部",
    "海外事业部",
    "物资部",
    "综合办公室",
    "市场开发部",
    "资产财务部",
]
SCOPE_DEPTS = DEFAULT_DEPTS.copy()
STAFF_COUNTS = {
    "HSE部": 4,
    "人力资源部": 4,
    "信息化中心": 6,
    "南京永利重工制造有限公司": 9,
    "合同管理部": 7,
    "审计部": 1,
    "市场开发部": 15,
    "技术中心": 20,
    "技术质量部": 7,
    "海外事业部": 6,
    "物资部": 8,
    "综合办公室": 19,
    "资产财务部": 21,
}
TOTAL_STAFF = sum(STAFF_COUNTS.values())
OUTPUT_PREFIX = ""
GRADE_ORDER = ["S", "A", "B", "C", "D", "E", "F"]
SALARY = {"S": 10000, "A": 9000, "B": 8000, "C": 6500, "D": 5000, "E": 4000, "F": 3000}
LOAD_ORDER = ["超负荷", "满负荷", "均衡", "欠饱和", "非饱和"]
LOAD_VALUE = {"非饱和": 1, "欠饱和": 2, "均衡": 3, "满负荷": 4, "超负荷": 5}

SUPERVISOR_RULE_PREFIX = "机关员工主管评价赋分规则"
GRADE_SCORE_RULES = [
    ("S", 95, float("inf")),
    ("A", 85, 95),
    ("B", 75, 85),
    ("C", 60, 75),
    ("D", 50, 60),
    ("E", 40, 50),
    ("F", float("-inf"), 40),
]
NECESSITY_SCORE = {"A": 0, "B": -1, "C": -3, "D": -5}
MIDDLE_RE = "中层|副总师|副总经理|董事长|党委书记|总经理|总会计师|纪委书记"
NANJING_STAFF = ["董如明", "蔡正荣", "殷玉善", "陈永奇", "王磊", "吴昊", "吴有鹏", "周子豪", "江晨阳"]


def scope_label() -> str:
    return "当前部门" if len(SCOPE_DEPTS) == 1 else "当前部门范围"


def scope_detail() -> str:
    return "、".join(SCOPE_DEPTS)


def scope_output_prefix() -> str:
    if len(SCOPE_DEPTS) == 1:
        return f"{SCOPE_DEPTS[0]}-"
    if set(SCOPE_DEPTS) == set(DEFAULT_DEPTS):
        return "全部部门-"
    return "部分部门-"


def latest(prefix: str, suffix: str) -> Path:
    files = [p for p in ROOT.iterdir() if p.name.startswith(prefix) and p.suffix.lower() == suffix]
    if not files:
        raise FileNotFoundError(prefix)
    return max(files, key=lambda p: (p.stat().st_mtime, p.name))


def unique_path(name: str) -> Path:
    path = ROOT / name
    if not path.exists():
        return path
    stem, suffix = path.stem, path.suffix
    i = 1
    while True:
        candidate = ROOT / f"{stem}-{i}{suffix}"
        if not candidate.exists():
            return candidate
        i += 1


def clean(v) -> str:
    if pd.isna(v):
        return ""
    text = str(v).strip()
    return "" if text.lower() == "nan" else text


def num(v) -> float:
    try:
        if pd.isna(v):
            return 0.0
        return float(v)
    except Exception:
        return 0.0


def fmt(v, digits: int = 2) -> str:
    try:
        f = float(v)
    except Exception:
        return "-"
    if abs(f - round(f)) < 1e-9:
        return str(int(round(f)))
    return f"{f:.{digits}f}".rstrip("0").rstrip(".")


def pct(n, d, digits: int = 1) -> str:
    return f"{(n / d * 100 if d else 0):.{digits}f}%"


def table(headers, rows) -> str:
    out = ["| " + " | ".join(headers) + " |", "| " + " | ".join(["---"] * len(headers)) + " |"]
    for row in rows:
        out.append("| " + " | ".join(str(x).replace("\n", "<br>") for x in row) + " |")
    return "\n".join(out)


def by_dept_names(records, name_col="姓名", dept_col="部门") -> str:
    grouped = defaultdict(list)
    for r in records:
        grouped[clean(r.get(dept_col))].append(clean(r.get(name_col)))
    parts = []
    for dept in SCOPE_DEPTS:
        names = [n for n in grouped.get(dept, []) if n]
        if names:
            parts.append(f"{dept}：{'、'.join(names)}")
    return "；".join(parts) or "-"


def read_sheet1(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="Sheet1", header=1)
    return df.loc[:, [c for c in df.columns if not str(c).startswith("Unnamed:")]]


def infer_department_scope() -> tuple[list[str], dict[str, int]]:
    """Infer one or more target departments from current workload/supervisor files."""
    names_by_dept: dict[str, set[str]] = defaultdict(set)
    try:
        workload = read_sheet1(latest("职能部门工作量摸排表", ".xlsx"))
        for _, row in workload.iterrows():
            dept = clean(row.get("部门"))
            name = clean(row.get("姓名"))
            if dept and name:
                names_by_dept[dept].add(name)
    except FileNotFoundError:
        pass
    try:
        supervisor = read_sheet1(latest("部门成员工作主管评定", ".xlsx"))
        for _, row in supervisor.iterrows():
            dept = clean(row.get("成员部门"))
            name = clean(row.get("成员姓名"))
            if dept and name:
                names_by_dept[dept].add(name)
    except FileNotFoundError:
        pass
    default_order = {dept: i for i, dept in enumerate(DEFAULT_DEPTS)}
    depts = sorted([dept for dept in names_by_dept if dept], key=lambda d: (default_order.get(d, 999), d))
    if not depts:
        return DEFAULT_DEPTS.copy(), STAFF_COUNTS.copy()
    try:
        roster = pd.read_excel(ROOT / "ERP使用情况.xlsx", sheet_name="职能部门人员台账")
        roster["部门"] = roster["部门"].map(clean)
        roster["姓名"] = roster["姓名"].map(clean)
        roster["职位"] = roster["职位"].map(clean)
        counts = {}
        for dept in depts:
            sub = roster[
                roster["部门"].eq(dept)
                & roster["姓名"].ne("")
                & ~roster["职位"].str.contains(MIDDLE_RE, na=False)
            ][["部门", "姓名"]].drop_duplicates(["部门", "姓名"])
            counts[dept] = len(sub) if not sub.empty else len(names_by_dept[dept])
        return depts, counts
    except Exception:
        return depts, {dept: len(names_by_dept[dept]) for dept in depts}


def latest_person_rows(df: pd.DataFrame, name_col: str, dept_col: str, time_col: str) -> pd.DataFrame:
    df = df[df[name_col].notna() & df[dept_col].isin(SCOPE_DEPTS)].copy()
    df[time_col] = df[time_col].astype(str)
    df = df.sort_values(time_col).drop_duplicates([dept_col, name_col], keep="last")
    return df.reset_index(drop=True)


def latest_optional_file(prefix: str, suffix: str) -> Path | None:
    files = [p for p in ROOT.iterdir() if p.name.startswith(prefix) and p.suffix.lower() == suffix]
    if not files:
        return None
    return max(files, key=lambda p: (p.stat().st_mtime, p.name))


def grade_from_score(score) -> str:
    try:
        value = float(score)
    except (TypeError, ValueError):
        return "-"
    for grade, low, high in GRADE_SCORE_RULES:
        if low <= value < high:
            return grade
    return "-"


def normalize_option_text(value) -> str:
    text = str(value or "").strip().replace("`", "")
    text = re.sub(r"\s+", "", text)
    text = text.replace("：", ":").replace("。", "")
    return text


def option_letter(value) -> str:
    text = str(value or "").strip()
    match = re.match(r"^([A-Z])\s*[.．、]", text)
    return match.group(1) if match else ""


def load_supervisor_rule_scores() -> dict[int, dict]:
    rule_path = latest_optional_file(SUPERVISOR_RULE_PREFIX, ".md")
    if not rule_path:
        return {}
    md = rule_path.read_text(encoding="utf-8")
    rules: dict[int, dict] = {}
    for qno in range(1, 21):
        pattern = rf"### 第{qno}题.*?(?=\n### 第{qno + 1}题|\n## |\Z)"
        match = re.search(pattern, md, flags=re.S)
        if not match:
            continue
        section = match.group(0)
        exact: dict[str, float] = {}
        by_letter: dict[str, float] = {}
        for line in section.splitlines():
            if not line.strip().startswith("|"):
                continue
            cells = [cell.strip().strip("`") for cell in line.strip().strip("|").split("|")]
            if len(cells) < 2 or cells[0] in {"选项", "---"}:
                continue
            try:
                score = float(cells[1].strip("`"))
            except ValueError:
                continue
            exact[normalize_option_text(cells[0])] = score
            letter = option_letter(cells[0])
            if letter:
                by_letter[letter] = score
        rules[qno] = {"exact": exact, "by_letter": by_letter}
    return rules


_SUPERVISOR_RULE_SCORES: dict[int, dict] | None = None


def supervisor_rule_scores() -> dict[int, dict]:
    global _SUPERVISOR_RULE_SCORES
    if _SUPERVISOR_RULE_SCORES is None:
        _SUPERVISOR_RULE_SCORES = load_supervisor_rule_scores()
    return _SUPERVISOR_RULE_SCORES


def question_header(rec: dict, qno: int) -> str | None:
    prefix = f"{qno}."
    prefix_alt = f"{qno}．"
    for key in rec:
        text = str(key).strip()
        if text.startswith(prefix) or text.startswith(prefix_alt):
            return key
    return None


def enrich_supervisor_score(rec: dict) -> None:
    rules = supervisor_rule_scores()
    total = 0.0
    exact_count = 0
    letter_count = 0
    missing: list[str] = []
    for qno in range(1, 21):
        header = question_header(rec, qno)
        answer = rec.get(header) if header else None
        q_rules = rules.get(qno, {})
        score = None
        normalized = normalize_option_text(answer)
        if normalized in q_rules.get("exact", {}):
            score = q_rules["exact"][normalized]
            exact_count += 1
        else:
            letter = option_letter(answer)
            if letter and letter in q_rules.get("by_letter", {}):
                score = q_rules["by_letter"][letter]
                letter_count += 1
        if score is None:
            score = 0.0
            missing.append(f"题{qno}")
        total += score
    pure_total = round(total, 1)
    necessity_adjustment = NECESSITY_SCORE.get(option_letter(rec.get("该成员所在岗位是否有存在的必要性")), 0)
    recalculated = round(pure_total + necessity_adjustment, 1)
    if float(pure_total).is_integer():
        pure_total = int(pure_total)
    if float(recalculated).is_integer():
        recalculated = int(recalculated)
    rec["纯20题总分"] = pure_total
    rec["岗位必要性修正分"] = necessity_adjustment
    rec["规则重算总分"] = recalculated
    rec["规则重算评级"] = grade_from_score(recalculated)
    rec["规则匹配情况"] = f"精确{exact_count}题/字母{letter_count}题/未匹配{len(missing)}题"
    rec["规则未匹配题目"] = "、".join(missing) if missing else "无"
    rec["规则重算一致性"] = "一致" if rec.get("自动生成评级") == rec.get("规则重算评级") else "不一致"


def load_supervisor_evaluations(deduplicate: bool = False):
    df = read_sheet1(latest("部门成员工作主管评定", ".xlsx"))
    df = df[df["成员姓名"].notna() & df["成员部门"].isin(SCOPE_DEPTS)].copy()
    df["评价时间(必填)"] = df["评价时间(必填)"].astype(str)
    df = df.sort_values(["成员部门", "成员姓名", "评价时间(必填)"])
    if deduplicate:
        df = df.drop_duplicates(["成员部门", "成员姓名"], keep="last")
    records = []
    for rec in df.to_dict("records"):
        enrich_supervisor_score(rec)
        records.append(rec)
    return records


def load_supervisor():
    return load_supervisor_evaluations(deduplicate=True)


def load_workload():
    df = latest_person_rows(read_sheet1(latest("职能部门工作量摸排表", ".xlsx")), "姓名", "部门", "填写日期")
    base = load_employee_base()
    if not base.empty:
        valid_people = set(map(tuple, base[["部门", "姓名"]].values.tolist()))
        df = df[df.apply(lambda r: (clean(r.get("部门")), clean(r.get("姓名"))) in valid_people, axis=1)].copy()
    return df.reset_index(drop=True)


def norm_load(v) -> str:
    text = clean(v)
    for item in LOAD_ORDER:
        if item in text:
            return item
    return ""


def diff_label(a, b) -> str:
    if not a or not b:
        return "-"
    d = LOAD_VALUE.get(a, 0) - LOAD_VALUE.get(b, 0)
    if d == 0:
        return "一致"
    return ("高" if d > 0 else "低") + f"{abs(d)}档"


def diff_value(a, b) -> int | None:
    if not a or not b:
        return None
    return LOAD_VALUE.get(a, 0) - LOAD_VALUE.get(b, 0)


def distribution(records, field, name_col="姓名", dept_col="部门", order=None):
    order = order or sorted({clean(r.get(field)) for r in records if clean(r.get(field))})
    rows = []
    for key in order:
        subset = [r for r in records if clean(r.get(field)) == key]
        if subset:
            rows.append([f"`{key}`", f"`{len(subset)}`", f"`{by_dept_names(subset, name_col, dept_col)}`"])
    return table(["档位", "人数", "部门及人员"], rows)


def build_supervisor_report(records):
    scores = [num(r.get("规则重算总分")) for r in records]
    rule_counts = Counter(r.get("规则重算评级") for r in records)
    sup_counts = Counter(clean(r.get("主管主观评级")) for r in records)
    auto_counts = Counter(clean(r.get("自动生成评级")) for r in records)
    q_rows = [["综合重算等级（含岗位必要性修正）", *[rule_counts.get(g, 0) for g in GRADE_ORDER]],
              ["主管评价等级", *[sup_counts.get(g, 0) for g in GRADE_ORDER]],
              ["表内自动等级", *[auto_counts.get(g, 0) for g in GRADE_ORDER]]]
    detail = []
    for r in sorted(records, key=lambda x: (SCOPE_DEPTS.index(clean(x.get("成员部门"))) if clean(x.get("成员部门")) in SCOPE_DEPTS else 99, -num(x.get("规则重算总分")))):
        grade = clean(r.get("规则重算评级"))
        sup = clean(r.get("主管主观评级"))
        detail.append([clean(r.get("成员部门")), clean(r.get("成员姓名")), fmt(r.get("规则重算总分")), grade, SALARY.get(grade, ""), sup, SALARY.get(sup, "")])
    md = f"""# 机关员工主管评价20题重算报告-{DATE}

## 一、测算口径

本报告依据 `机关员工主管评价赋分规则`，结合主管对员工 `20` 道问卷题及 `岗位必要性` 独立观察项的评价结果测算形成。本轮先读取20题原始选项文本逐题重算，再按岗位必要性观察项执行扣分修正，不覆盖表内原公式。

## 二、整体结果

- 纳入测算评价记录：`{len(records)}` 条
- 涉及员工：`{len({clean(r.get("成员姓名")) for r in records})}` 人
- 同一员工如存在多名主管评价，按主管分别列示，不合并、不去重。

### 1. 三类等级分布

{table(["口径", *GRADE_ORDER], q_rows)}

## 三、样本校核结果

### 1. 当前总分统计

{table(["指标", "数值"], [["`最低分`", f"`{fmt(min(scores))}`"], ["`25分位`", f"`{fmt(pd.Series(scores).quantile(.25))}`"], ["`中位数`", f"`{fmt(pd.Series(scores).median())}`"], ["`75分位`", f"`{fmt(pd.Series(scores).quantile(.75))}`"], ["`最高分`", f"`{fmt(max(scores))}`"]])}

### 2. 当前样本分布

{distribution(records, "规则重算评级", "成员姓名", "成员部门", GRADE_ORDER).replace("档位", "档位").replace("部门及人员", "部门及人员")}

## 四、人员明细

{table(["部门", "姓名", "主管姓名", "评价时间", "纯20题总分", "岗位必要性修正分", "综合重算总分", "综合重算等级", "综合重算等级对应薪资中位数", "主管评价等级", "主管评价等级对应薪资中位数"], [[clean(r.get("成员部门")), clean(r.get("成员姓名")), clean(r.get("主管姓名(必填)")), clean(r.get("评价时间(必填)")), fmt(r.get("纯20题总分")), fmt(r.get("岗位必要性修正分")), fmt(r.get("规则重算总分")), clean(r.get("规则重算评级")), SALARY.get(clean(r.get("规则重算评级")), ""), clean(r.get("主管主观评级")), SALARY.get(clean(r.get("主管主观评级")), "")] for r in sorted(records, key=lambda x: (SCOPE_DEPTS.index(clean(x.get("成员部门"))) if clean(x.get("成员部门")) in SCOPE_DEPTS else 99, clean(x.get("成员姓名")), clean(x.get("评价时间(必填)"))))])}
"""
    return md


def build_workload_report(workload, supervisor):
    sup_load = {clean(r.get("成员姓名")): norm_load(r.get("13. 以目前的业务量，您认为该员工的工作忙闲程度如何？")) for r in supervisor}
    sup_load_text = {
        clean(r.get("成员姓名")): clean(r.get("13. 以目前的业务量，您认为该员工的工作忙闲程度如何？"))
        for r in supervisor
    }
    records = []
    for r in workload.to_dict("records"):
        rec = dict(r)
        rec["员工自评层级"] = norm_load(rec.get("总体工作负荷"))
        rec["自动层级"] = norm_load(rec.get("自动工作负荷"))
        rec["主管层级"] = sup_load.get(clean(rec.get("姓名")), "")
        rec["主管主观工作负荷原文"] = sup_load_text.get(clean(rec.get("姓名")), "")
        rec["员工-主管差异"] = diff_label(rec["员工自评层级"], rec["主管层级"])
        rec["自动-主管差异"] = diff_label(rec["自动层级"], rec["主管层级"])
        rec["员工-主管差值"] = diff_value(rec["员工自评层级"], rec["主管层级"])
        rec["自动-主管差值"] = diff_value(rec["自动层级"], rec["主管层级"])
        records.append(rec)
    def stat(col):
        s = pd.Series([num(r.get(col)) for r in records])
        return [fmt(s.mean()), fmt(s.median()), fmt(s.quantile(.75)), fmt(s.max())]
    time_map = [
        ("日类折算日均时长", "每日工作最低负荷（日）", "每日工作最高负荷（日）"),
        ("周类折算日均时长", "每日工作最低负荷（周）", "每日工作最高负荷(周）"),
        ("月类折算日均时长", "每日工作最低负荷（月）", "每日工作最高负荷（月）"),
        ("季类折算日均时长", "每日工作最低负荷（季）", "每日工作最高负荷（季）"),
        ("年类折算日均时长", "每日工作最低负荷（临）", "每日工作最高负荷（临）"),
    ]
    for r in records:
        for label, lo, hi in time_map:
            r[label] = (num(r.get(lo)) + num(r.get(hi))) / 2
        r["折算日均总时长"] = num(r.get("总每日工作量平均负荷（小时）"))
        r["时长波动区间"] = num(r.get("总每日工作量最高负荷（小时）")) - num(r.get("总每日工作量最低负荷（小时）"))
    time_rows = [[label, *stat(label)] for label, _, _ in time_map] + [["折算日均总时长", *stat("折算日均总时长")], ["时长波动区间", *stat("时长波动区间")]]
    summary_rows = []
    for label, field in [("员工自评总体工作负荷", "员工自评层级"), ("自动工作负荷", "自动层级"), ("主管主观工作负荷", "主管层级")]:
        c = Counter(r.get(field) for r in records if r.get(field))
        summary_rows.append([label, sum(c.values()), *[c.get(x, 0) for x in LOAD_ORDER]])
    anomalies = []
    for r in sorted(records, key=lambda x: -x["折算日均总时长"]):
        flags = []
        if r["折算日均总时长"] >= 20:
            flags.append("总日均时长偏高")
        if r["时长波动区间"] >= 10:
            flags.append("时长波动偏大")
        if r["折算日均总时长"] and r["日类折算日均时长"] / r["折算日均总时长"] >= .8:
            flags.append("日类工时占比过高")
        if flags:
            anomalies.append([clean(r.get("部门")), clean(r.get("姓名")), fmt(r["日类折算日均时长"]), fmt(r["周类折算日均时长"]), fmt(r["月类折算日均时长"]), fmt(r["季类折算日均时长"]), fmt(r["年类折算日均时长"]), fmt(r["折算日均总时长"]), fmt(r["时长波动区间"]), "；".join(flags)])
    detail = [[clean(r.get("部门")), clean(r.get("姓名")), fmt(r["日类折算日均时长"]), fmt(r["周类折算日均时长"]), fmt(r["月类折算日均时长"]), fmt(r["季类折算日均时长"]), fmt(r["年类折算日均时长"]), fmt(r["折算日均总时长"]), fmt(r["时长波动区间"]), clean(r.get("总体工作负荷")), clean(r.get("您目前对自己岗位状态的总体态度更接近哪一种？")), clean(r.get("自动工作负荷")), r["主管层级"], r["员工-主管差异"], r["自动-主管差异"]] for r in records]
    emp_sup_large = [
        [clean(r.get("部门")), clean(r.get("姓名")), clean(r.get("总体工作负荷")), clean(r.get("自动工作负荷")), r["主管主观工作负荷原文"] or r["主管层级"], r["员工-主管差值"]]
        for r in records
        if r.get("员工-主管差值") is not None and abs(r["员工-主管差值"]) >= 2
    ]
    auto_sup_large = [
        [clean(r.get("部门")), clean(r.get("姓名")), clean(r.get("总体工作负荷")), clean(r.get("自动工作负荷")), r["主管主观工作负荷原文"] or r["主管层级"], r["自动-主管差值"]]
        for r in records
        if r.get("自动-主管差值") is not None and abs(r["自动-主管差值"]) >= 2
    ]
    emp_sup_large = sorted(emp_sup_large, key=lambda x: (abs(int(x[-1])), x[0], x[1]), reverse=True)
    auto_sup_large = sorted(auto_sup_large, key=lambda x: (abs(int(x[-1])), x[0], x[1]), reverse=True)
    md = f"""# 机关员工工作负荷报告-{DATE}

## 一、报告口径

本报告依据 `机关员工工作负荷认定规则`，结合{scope_label()}工作量摸排表与主管评价表中涉及工作负荷的记录结果形成。

## 二、整体结果

- 统计范围内{scope_label()}员工共 `{TOTAL_STAFF}` 人，涉及部门：`{scope_detail()}`。
- 已完成工作量摸排并形成有效员工自评负荷样本 `{len(workload)}` 人。
- 已完成主管评价并形成有效主管主观负荷样本 `{len(supervisor)}` 人。
- 员工自评与主管主观负荷可直接交叉比对样本 `{sum(1 for r in records if r['主管层级'])}` 人。

### 1. 三类工作负荷分布

{table(["口径", "有效样本", *LOAD_ORDER], summary_rows)}

### 2. 工时折算统计

{table(["时长口径", "平均值", "中位数", "75分位", "最大值"], time_rows)}

## 三、样本分布

### 1. 员工自评总体工作负荷

{distribution(records, "员工自评层级", "姓名", "部门", LOAD_ORDER)}

### 2. 自动工作负荷

{distribution(records, "自动层级", "姓名", "部门", LOAD_ORDER)}

### 3. 主管主观工作负荷

{distribution(records, "主管层级", "姓名", "部门", LOAD_ORDER)}

## 四、工时异常情况

{table(["部门", "姓名", "日类折算日均时长", "周类折算日均时长", "月类折算日均时长", "季类折算日均时长", "年类折算日均时长", "折算日均总时长", "时长波动区间", "异常类型"], anomalies[:60])}

## 五、差异情况

### 1. 员工自评与主管主观认定差异

{table(["差异档位", "人数"], Counter(r["员工-主管差异"] for r in records if r["员工-主管差异"] != "-").items())}

### 2. 自动工作负荷与主管主观认定差异

{table(["差异档位", "人数"], Counter(r["自动-主管差异"] for r in records if r["自动-主管差异"] != "-").items())}

### 3. 员工自评与主管认定差异较大样本

{table(["部门", "姓名", "员工自评总体工作负荷", "自动工作负荷", "主管主观工作负荷", "差值"], emp_sup_large)}

### 4. 自动负荷与主管认定差异较大样本

{table(["部门", "姓名", "员工自评总体工作负荷", "自动工作负荷", "主管主观工作负荷", "差值"], auto_sup_large)}

## 六、人员明细

{table(["部门", "姓名", "日类折算日均时长", "周类折算日均时长", "月类折算日均时长", "季类折算日均时长", "年类折算日均时长", "折算日均总时长", "时长波动区间", "员工自评总体工作负荷", "员工岗位状态态度", "自动工作负荷", "主管主观工作负荷", "员工-主管差异", "自动-主管差异"], detail)}
"""
    return md, records


def load_employee_base() -> pd.DataFrame:
    try:
        roster = pd.read_excel(ROOT / "ERP使用情况.xlsx", sheet_name="职能部门人员台账")
        roster["部门"] = roster["部门"].map(clean)
        roster["姓名"] = roster["姓名"].map(clean)
        roster["职位"] = roster["职位"].map(clean)
        base = roster[
            roster["部门"].isin(SCOPE_DEPTS)
            & roster["姓名"].ne("")
            & ~roster["职位"].str.contains(MIDDLE_RE, na=False)
        ][["部门", "姓名"]].drop_duplicates(["部门", "姓名"])
        if not base.empty:
            base["部门排序"] = base["部门"].map({d: i for i, d in enumerate(SCOPE_DEPTS)})
            return base.sort_values(["部门排序", "姓名"]).drop(columns=["部门排序"]).reset_index(drop=True)
    except Exception:
        pass

    rows = []
    try:
        workload = read_sheet1(latest("职能部门工作量摸排表", ".xlsx"))
        for _, item in workload.iterrows():
            dept = clean(item.get("部门"))
            name = clean(item.get("姓名"))
            if dept in SCOPE_DEPTS and name:
                rows.append({"部门": dept, "姓名": name})
    except FileNotFoundError:
        pass
    try:
        supervisor = read_sheet1(latest("部门成员工作主管评定", ".xlsx"))
        for _, item in supervisor.iterrows():
            dept = clean(item.get("成员部门"))
            name = clean(item.get("成员姓名"))
            if dept in SCOPE_DEPTS and name:
                rows.append({"部门": dept, "姓名": name})
    except FileNotFoundError:
        pass
    if not rows:
        return pd.DataFrame(columns=["部门", "姓名"])
    base = pd.DataFrame(rows).drop_duplicates(["部门", "姓名"])
    base["部门排序"] = base["部门"].map({d: i for i, d in enumerate(SCOPE_DEPTS)})
    return base.sort_values(["部门排序", "姓名"]).drop(columns=["部门排序"]).reset_index(drop=True)


def build_erp_report(supervisor):
    employees = load_employee_base()
    grades = {clean(r.get("成员姓名")): {"按20题计算等级": clean(r.get("规则重算评级")) or clean(r.get("自动生成评级")) or "暂无评级", "主管评价等级": clean(r.get("主管主观评级")) or "暂无评级"} for r in supervisor if clean(r.get("成员姓名"))}
    erp_source = ROOT / "ERP使用情况.xlsx"
    ops = pd.read_excel(erp_source, sheet_name="近三年")
    for col in ["部门", "人员名称", "参与类型"]:
        ops[col] = ops[col].map(clean)
    for col in ["参与流程数量", "处理时间（分钟）", "驳回次数", "有效评论次数"]:
        ops[col] = pd.to_numeric(ops[col], errors="coerce").fillna(0)

    raw_names = set(ops.loc[ops["参与流程数量"].gt(0), "人员名称"])
    roster_names = set(pd.read_excel(erp_source, sheet_name="职能部门人员台账")["姓名"].map(clean))

    role = ops.pivot_table(index="人员名称", columns="参与类型", values="参与流程数量", aggfunc="sum", fill_value=0).reset_index()
    for c in ["发起者", "参与审批者"]:
        if c not in role:
            role[c] = 0
    agg = ops.groupby("人员名称", as_index=False).agg(
        处理时间分钟=("处理时间（分钟）", "sum"),
        驳回次数=("驳回次数", "sum"),
        有效评论次数=("有效评论次数", "sum"),
    )
    person = employees.merge(role, left_on="姓名", right_on="人员名称", how="left").merge(
        agg, left_on="姓名", right_on="人员名称", how="left"
    )
    for c in ["发起者", "参与审批者", "处理时间分钟", "驳回次数", "有效评论次数"]:
        person[c] = pd.to_numeric(person[c], errors="coerce").fillna(0)
    person["有ERP痕迹"] = (person["发起者"] + person["参与审批者"]).gt(0)
    person["在ERP人员台账"] = person["姓名"].isin(roster_names)
    person["按20题计算等级"] = person["姓名"].map(lambda x: grades.get(x, {}).get("按20题计算等级", "暂无评级"))
    person["主管评价等级"] = person["姓名"].map(lambda x: grades.get(x, {}).get("主管评价等级", "暂无评级"))
    person["审批驳回率"] = person.apply(lambda r: pct(r["驳回次数"], r["参与审批者"], 2), axis=1)
    person["审批评论率"] = person.apply(lambda r: pct(r["有效评论次数"], r["参与审批者"], 2), axis=1)
    person["审批平均处理小时"] = person.apply(
        lambda r: fmt(r["处理时间分钟"] / 60 / r["参与审批者"], 2) if r["参与审批者"] else "0", axis=1
    )
    person["流程合计"] = person["发起者"] + person["参与审批者"]
    person["总评论率"] = person.apply(lambda r: pct(r["有效评论次数"], r["流程合计"], 2), axis=1)
    person["平均处理小时"] = person.apply(lambda r: fmt(r["处理时间分钟"] / 60 / r["流程合计"], 2) if r["流程合计"] else "0", axis=1)
    person["平均处理小时数"] = person.apply(lambda r: r["处理时间分钟"] / 60 / r["流程合计"] if r["流程合计"] else 0, axis=1)
    person["审批平均处理小时数"] = person.apply(lambda r: r["处理时间分钟"] / 60 / r["参与审批者"] if r["参与审批者"] else 0, axis=1)

    dept_rows = []
    for dept in SCOPE_DEPTS:
        sub = person[person["部门"] == dept]
        approvals = sub["参与审批者"].sum()
        dept_rows.append(
            [
                dept,
                len(sub),
                int(sub["有ERP痕迹"].sum()),
                int((~sub["有ERP痕迹"]).sum()),
                int(sub["发起者"].gt(0).sum()),
                int(sub["参与审批者"].gt(0).sum()),
                fmt(sub["发起者"].sum(), 0),
                fmt(approvals, 0),
                pct(sub["有效评论次数"].sum(), approvals, 2),
                pct(sub["驳回次数"].sum(), approvals, 2),
            ]
        )

    role_counts = Counter()
    for _, r in person.iterrows():
        if not r["有ERP痕迹"]:
            role_counts["无ERP痕迹"] += 1
        elif r["发起者"] > 0 and r["参与审批者"] > 0:
            role_counts["兼有发起和审批"] += 1
        elif r["发起者"] > 0:
            role_counts["仅发起"] += 1
        else:
            role_counts["仅审批"] += 1

    no_ops = person[~person["有ERP痕迹"]]
    no_ops_roster = no_ops[no_ops["在ERP人员台账"]]
    no_ops_not_roster = no_ops[~no_ops["在ERP人员台账"]]

    grade_sections = []
    for grade in ["S", "A", "B", "C", "D", "E", "F", "暂无评级"]:
        sub = person[person["主管评价等级"] == grade].copy()
        if sub.empty:
            continue
        sub["排序"] = sub["发起者"] + sub["参与审批者"]
        sub = sub.sort_values(["部门", "排序"], ascending=[True, False])
        rows = [
            [
                r["部门"],
                r["姓名"],
                fmt(r["发起者"], 0),
                fmt(r["参与审批者"], 0),
                r["审批驳回率"],
                r["审批评论率"],
                r["审批平均处理小时"],
                r["按20题计算等级"],
                r["主管评价等级"],
            ]
            for _, r in sub.iterrows()
        ]
        grade_sections.append(
            f"### {grade}档人员情况\n\n"
            + table(["部门", "姓名", "发起流程数", "审批流程数", "审批驳回率", "审批评论率", "审批平均处理小时", "按20题计算等级", "主管评价等级"], rows)
        )

    start_people = int(person["发起者"].gt(0).sum())
    approve_people = int(person["参与审批者"].gt(0).sum())
    start_total = person["发起者"].sum()
    approve_total = person["参与审批者"].sum()
    approve_hours = person["处理时间分钟"].sum() / 60
    avg_approve_hours = approve_hours / approve_total if approve_total else 0

    top_start = person[person["发起者"].gt(0)].sort_values("发起者", ascending=False).head(15)
    top_approve = person[person["参与审批者"].gt(0)].sort_values("参与审批者", ascending=False).head(15)
    high_comment = person[person["流程合计"].ge(100)].copy()
    high_comment["总评论率数"] = high_comment.apply(lambda r: num(r["有效评论次数"]) / num(r["流程合计"]) if r["流程合计"] else 0, axis=1)
    high_comment = high_comment.sort_values("总评论率数", ascending=False).head(10)
    low_start = person[person["发起者"].gt(0)].sort_values(["发起者", "部门", "姓名"]).head(10)
    low_approve = person[person["参与审批者"].gt(0)].sort_values(["参与审批者", "部门", "姓名"]).head(10)

    high_flow_low_comment = person[(person["流程合计"].ge(500)) & (person["有效评论次数"] / person["流程合计"].replace(0, pd.NA) < 0.02)].copy()
    high_flow_low_comment = high_flow_low_comment.sort_values("流程合计", ascending=False).head(15)
    high_flow_slow = person[(person["流程合计"].ge(100)) & (person["平均处理小时数"].ge(80))].sort_values("平均处理小时数", ascending=False).head(15)
    slow_approval = person[(person["参与审批者"].gt(0)) & (person["审批平均处理小时数"].ge(120))].sort_values("审批平均处理小时数", ascending=False)
    no_opinion = person[(person["参与审批者"].gt(0)) & (person["有效评论次数"].eq(0)) & (person["驳回次数"].eq(0))].sort_values("参与审批者", ascending=False)
    key_slow = slow_approval[slow_approval["按20题计算等级"].isin(["S", "A", "B"])]

    no_ops_rows = []
    for dept in SCOPE_DEPTS:
        sub = no_ops[no_ops["部门"] == dept]
        if sub.empty:
            continue
        in_roster = sub[sub["在ERP人员台账"]]["姓名"].tolist()
        not_roster = sub[~sub["在ERP人员台账"]]["姓名"].tolist()
        no_ops_rows.append([dept, len(sub), "、".join(sub["姓名"]), "、".join(in_roster) or "-", "、".join(not_roster) or "-"])

    def people_names(df):
        return "、".join(df["姓名"].astype(str).tolist()) or "无"

    dept_flow_summary = "；".join(
        f"`{row[0]}` 参与流程 `{row[7]}` 次、无ERP痕迹 `{row[3]}` 人"
        for row in dept_rows
    )
    conclusion = [
        f"{scope_label()}ERP口径下，{dept_flow_summary}。",
        f"当前统计范围内有 ERP 操作痕迹 `{int(person['有ERP痕迹'].sum())}` 人，无 ERP 操作痕迹 `{len(no_ops)}` 人。",
        f"当前 `参与流程平均审核时间超过120小时` 的人员共有 `{len(slow_approval)}` 人，其中按20题计算等级为 `B` 及以上的有 `{len(key_slow)}` 人，分别为 `{people_names(key_slow)}`。",
        f"当前 `参与流程存在但无任何审批意见和驳回行为` 的人员共有 `{len(no_opinion)}` 人，需要结合其审批职责判断是低风险流程、代办流转，还是过程留痕不足。",
        "当前ERP表未提供单独 `回复率` 字段，因此本轮继续围绕 `发起流程数量`、`参与流程数量`、`评论率`、`驳回率`、`平均处理时长` 和 `是否存在操作痕迹` 做判断。",
    ]

    md = f"""# 机关员工ERP使用情况分析-{DATE}

本报告依据 `ERP使用情况.xlsx` 中 `近三年` 工作簿，并参考{scope_label()}工作摸排表和主管评价表中的人员名单开展分析，用于评判员工和部门数字化使用情况，并作为工作量、工作岗位和工作责任分析的一个辅助维度。

## 一、统计范围

- 统计范围以{scope_label()} `职能部门工作量摸排表` 和 `部门成员工作主管评定` 中出现的人员为基础。
- 本轮自动识别部门为 `{scope_detail()}`，形成统计人数 `{len(person)}` 人。
- 同一员工如存在多条 ERP 记录，按姓名汇总后统一归并到当前台账部门。
- ERP 源表中存在但不属于上述统计范围的流程记录，不纳入本报告人员统计。

## 二、整体结果

- 纳入统计员工：`{len(person)}` 人
- 有 ERP 操作痕迹：`{int(person['有ERP痕迹'].sum())}` 人
- 无 ERP 操作痕迹：`{len(person) - int(person['有ERP痕迹'].sum())}` 人
- 其中，存在于 `职能部门人员台账` 但近三年无任何操作：`{len(no_ops_roster)}` 人
- 其中，未出现在 `职能部门人员台账` 且近三年无任何操作：`{len(no_ops_not_roster)}` 人
- ERP `近三年` 原始操作表中有实际流程量的人员共 `{len(raw_names)}` 人，本报告仅纳入{scope_label()}统计范围内人员。
- 当前源表可直接使用的互动指标为 `有效评论次数` 和 `驳回次数`，未提供单独 `回复率` 字段。

## 三、发起流程与参与流程概况

### 1. 发起流程概况

{table(["口径", "涉及人数", "发起流程总数"], [["发起流程", start_people, fmt(start_total, 0)]])}

### 2. 参与流程概况

{table(["口径", "涉及人数", "参与流程总数", "审核总小时", "平均审核小时", "评论率", "驳回率"], [["参与流程", approve_people, fmt(approve_total, 0), fmt(approve_hours, 2), fmt(avg_approve_hours, 2), pct(person["有效评论次数"].sum(), approve_total, 2), pct(person["驳回次数"].sum(), approve_total, 2)]])}

## 四、部门分布

{table(["部门", "部门人数", "有ERP痕迹人数", "无ERP痕迹人数", "发起人数", "审批人数", "发起流程总数", "参与流程总数", "参与流程评论率", "参与流程驳回率"], dept_rows)}

## 五、角色特征分布

{table(["角色特征", "人数"], [[k, role_counts.get(k, 0)] for k in ["兼有发起和审批", "仅发起", "仅审批", "无ERP痕迹"]])}

## 六、按人员评价等级观察ERP使用情况

{chr(10).join(grade_sections)}

## 七、重点人员表现

### 1. 发起流程数量靠前人员

{table(["部门", "姓名", "发起流程数"], [[r["部门"], r["姓名"], fmt(r["发起者"], 0)] for _, r in top_start.iterrows()])}

### 2. 参与流程数量靠前人员

{table(["部门", "姓名", "审批流程数", "审批评论率", "审批驳回率", "审批平均处理小时"], [[r["部门"], r["姓名"], fmt(r["参与审批者"], 0), r["审批评论率"], r["审批驳回率"], r["审批平均处理小时"]] for _, r in top_approve.iterrows()])}

### 3. 评论率较高人员（总流程数不少于100）

{table(["部门", "姓名", "总流程数", "总评论次数", "加权评论率"], [[r["部门"], r["姓名"], fmt(r["流程合计"], 0), fmt(r["有效评论次数"], 0), r["总评论率"]] for _, r in high_comment.iterrows()])}

### 4. 发起流程最少的10人（剔除完全无发起流程样本）

{table(["部门", "姓名", "发起流程数", "按20题计算等级", "主管评价等级"], [[r["部门"], r["姓名"], fmt(r["发起者"], 0), r["按20题计算等级"], r["主管评价等级"]] for _, r in low_start.iterrows()])}

### 5. 参与流程最少的10人（剔除完全无参与流程样本）

{table(["部门", "姓名", "审批流程数", "审批平均处理小时", "按20题计算等级", "主管评价等级"], [[r["部门"], r["姓名"], fmt(r["参与审批者"], 0), r["审批平均处理小时"], r["按20题计算等级"], r["主管评价等级"]] for _, r in low_approve.iterrows()])}

## 八、无ERP痕迹情况

{table(["部门", "无ERP痕迹人数", "名单", "ERP台账存在但无操作", "ERP台账也不存在"], no_ops_rows)}

## 九、异常信号

### 1. 高流程量但评论率偏低

{table(["部门", "姓名", "总流程数", "总评论次数", "加权评论率", "平均处理小时"], [[r["部门"], r["姓名"], fmt(r["流程合计"], 0), fmt(r["有效评论次数"], 0), r["总评论率"], r["平均处理小时"]] for _, r in high_flow_low_comment.iterrows()])}

### 2. 高流程量且平均处理时长偏高

{table(["部门", "姓名", "总流程数", "平均处理小时", "加权评论率"], [[r["部门"], r["姓名"], fmt(r["流程合计"], 0), r["平均处理小时"], r["总评论率"]] for _, r in high_flow_slow.iterrows()])}

### 3. 参与流程平均审核时间超过120小时

{table(["部门", "姓名", "审批流程数", "审批平均处理小时", "审批评论率", "审批驳回率", "按20题计算等级", "主管评价等级"], [[r["部门"], r["姓名"], fmt(r["参与审批者"], 0), r["审批平均处理小时"], r["审批评论率"], r["审批驳回率"], r["按20题计算等级"], r["主管评价等级"]] for _, r in slow_approval.iterrows()])}

### 4. 参与流程存在，但无任何审批意见和驳回行为

{table(["部门", "姓名", "审批流程数", "审批平均处理小时", "按20题计算等级", "主管评价等级"], [[r["部门"], r["姓名"], fmt(r["参与审批者"], 0), r["审批平均处理小时"], r["按20题计算等级"], r["主管评价等级"]] for _, r in no_opinion.iterrows()])}

### 5. 主要人员与异常情况交叉

以下人员当前 `按20题计算等级` 为 `B` 及以上，但参与流程平均审核时间仍超过 `120` 小时，说明其虽然属于当前较高等级或关键岗位人员，ERP 审批时长口径上仍存在需要复核的异常信号。

{table(["部门", "姓名", "审批流程数", "审批平均处理小时", "按20题计算等级", "主管评价等级"], [[r["部门"], r["姓名"], fmt(r["参与审批者"], 0), r["审批平均处理小时"], r["按20题计算等级"], r["主管评价等级"]] for _, r in key_slow.iterrows()])}

## 十、结论判断

{chr(10).join("- " + item for item in conclusion)}
"""
    return md


def build_cert_report(supervisor):
    def empty_cert(value) -> bool:
        text = clean(value).replace("；", "").replace(";", "").replace("、", "").replace("/", "").strip()
        return text in {"", "无", "没", "没有", "暂无", "无无"}

    def cert_value(value) -> str:
        return "" if empty_cert(value) else clean(value)

    def merge_cert(*values) -> str:
        seen = set()
        merged = []
        for value in values:
            text = cert_value(value)
            if not text:
                continue
            for part in re.split(r"[；;、,，]", text):
                part = cert_value(part)
                if part and part not in seen:
                    seen.add(part)
                    merged.append(part)
        return "；".join(merged)

    roster = pd.read_excel("ERP使用情况.xlsx", sheet_name="职能部门人员台账")
    roster["姓名"] = roster["姓名"].map(clean)
    roster_start = {}
    roster_honor = {}
    for _, item in roster.iterrows():
        name = clean(item.get("姓名"))
        if not name:
            continue
        if name not in roster_start:
            start = item.get("参加工作时间")
            if pd.isna(start):
                start = item.get("到本单位时间")
            roster_start[name] = start
        roster_honor[name] = merge_cert(
            roster_honor.get(name, ""),
            item.get("优秀情况（文本化）"),
            item.get("安全先进个人情况（文本化）"),
        )

    yearend_file = latest("年终总结", ".xlsx")
    yearend = pd.read_excel(yearend_file, sheet_name="Sheet1", header=1)
    yearend["姓名"] = yearend["姓名(必填)"].map(clean)
    yearend_map = {}
    for _, item in yearend.iterrows():
        name = clean(item.get("姓名"))
        if not name or name in yearend_map:
            continue
        yearend_map[name] = {
            "部门": clean(item.get("部门")),
            "职位": clean(item.get("职位")),
            "职称证书": cert_value(item.get("有效职称证书")),
            "资格证书": cert_value(item.get("有效资格证书")),
            "荣誉证书": cert_value(item.get("其他荣誉证书")),
        }

    rows = []
    review = {
        "陈淑怡": "无", "左晋铭": "焊工证", "汤冉": "无", "王培培": "初级会计证",
        "周子豪": "无", "董健": "无", "林威": "无", "王仁斌": "起重工证",
        "徐丹": "无", "杨悦": "无", "徐成威": "无", "沈冲": "焊工证、高压电工证",
        "薛慧": "安全C证", "杨荣康": "无", "赵乐": "助理工程师", "周勋禹": "无",
    }
    for r in supervisor:
        name = clean(r.get("成员姓名"))
        yearend_item = yearend_map.get(name, {})
        title = merge_cert(yearend_item.get("职称证书"), r.get("职称证书"))
        qual = merge_cert(yearend_item.get("资格证书"), r.get("资格证书"))
        honor = merge_cert(yearend_item.get("荣誉证书"), r.get("荣誉证书"), roster_honor.get(name))
        if name in review and review[name] != "无":
            qual = merge_cert(qual, review[name])
        has_any = any([title, qual, honor])
        years = 0
        try:
            start_source = roster_start.get(name) or r.get("成员到本单位时间")
            start = pd.to_datetime(start_source)
            years = (pd.Timestamp(DATE) - start).days / 365.25
        except Exception:
            years = 0
        rows.append({
            **r,
            "姓名": name,
            "部门": clean(r.get("成员部门")),
            "职称证书": title,
            "资格证书": qual,
            "荣誉证书": honor,
            "有证书": has_any,
            "工作年限": years,
            "年终总结匹配": name in yearend_map,
        })
    no_cert = [r for r in rows if not r["有证书"]]
    full2 = [r for r in no_cert if r["工作年限"] >= 2]
    less2 = [r for r in no_cert if r["工作年限"] < 2]
    summary = []
    for grade in GRADE_ORDER:
        for dept in SCOPE_DEPTS:
            sub = [r for r in rows if clean(r.get("规则重算评级")) == grade and r["部门"] == dept]
            if sub:
                no_sub = [r for r in sub if not r["有证书"]]
                summary.append([
                    grade,
                    dept,
                    len(sub),
                    SALARY[grade],
                    sum(1 for r in sub if r["职称证书"]),
                    sum(1 for r in sub if r["资格证书"]),
                    sum(1 for r in sub if r["荣誉证书"]),
                    len(no_sub),
                    sum(1 for r in no_sub if r["工作年限"] < 2),
                    sum(1 for r in no_sub if r["工作年限"] >= 2),
                ])
    grade_index = {g: i for i, g in enumerate(GRADE_ORDER)}
    rows_sorted = sorted(rows, key=lambda r: (grade_index.get(clean(r.get("规则重算评级")), 99), SCOPE_DEPTS.index(r["部门"]) if r["部门"] in SCOPE_DEPTS else 99, -num(r.get("规则重算总分")), r["姓名"]))
    def cert_display(value):
        return clean(value).replace("；", "<br>")

    detail = [[r["部门"], r["姓名"], fmt(r["工作年限"], 1), clean(r.get("规则重算评级")), SALARY.get(clean(r.get("规则重算评级")), ""), clean(r.get("主管主观评级")), SALARY.get(clean(r.get("主管主观评级")), ""), cert_display(r["职称证书"]), cert_display(r["资格证书"]), cert_display(r["荣誉证书"])] for r in rows_sorted]
    current_names = {r["姓名"] for r in rows}
    review_in_scope = {name: cert for name, cert in review.items() if name in current_names}
    review_rows = [[i, name, cert] for i, (name, cert) in enumerate(review_in_scope.items(), 1)]
    review_section = ""
    if review_rows:
        review_section = f"""
## 五、人力资源复核明细

{table(["序号", "人员姓名", "证书核实情况"], review_rows)}
"""
    unmatched = [r for r in rows if not r["年终总结匹配"]]
    no_cert_by_sup = []
    for grade in GRADE_ORDER:
        sub = [r for r in no_cert if clean(r.get("主管主观评级")) == grade]
        if sub:
            grouped = defaultdict(list)
            for r in sub:
                grouped[r["部门"]].append(r["姓名"])
            people = "；".join(f"{dept}：{'、'.join(names)}" for dept, names in grouped.items())
            no_cert_by_sup.append([grade, len(sub), people])
    return f"""# 机关员工证书客观维度分析-{DATE}

## 一、分析口径

本次分析以{scope_label()}已经纳入主管评价的 `{len(rows)}` 名非中层员工为基础。员工 `按20题计算等级` 统一采用{scope_label()} `主管评价20题重算报告-{DATE}.md` 中的规则重算等级；证书信息优先取当前文件夹 `年终总结` 开头文件中的 `有效职称证书`、`有效资格证书`、`其他荣誉证书` 字段，并结合主管评价表中的证书字段合并去重；荣誉证书同步采集 `ERP使用情况.xlsx` 中 `职能部门人员台账` 的 `优秀情况（文本化）`、`安全先进个人情况（文本化）` 信息。若{scope_label()}涉及人力资源 `人员证书复核.png` 中人员，则同步纳入复核口径。

## 二、整体结论

- 纳入分析人数：`{len(rows)}`
- 合并后有职称证书人数：`{sum(1 for r in rows if r['职称证书'])}`
- 合并后有资格证书人数：`{sum(1 for r in rows if r['资格证书'])}`
- 合并后有荣誉证书人数：`{sum(1 for r in rows if r['荣誉证书'])}`
- 无任何证书人数：`{len(no_cert)}`
- 其中工作未满2年且无任何证书人数：`{len(less2)}`
- 其中工作满2年仍无任何证书人数：`{len(full2)}`
- 年终总结未匹配到记录：`{len(unmatched)}`

## 三、按等级部门汇总

{table(["按20题计算等级", "部门", "人数", "等级对应薪资中位数", "职称证书人数", "资格证书人数", "荣誉证书人数", "无任何证书人数", "未满2年无证书人数", "满2年无证书人数"], summary)}

## 四、年终总结未匹配名单

{table(["部门", "姓名", "工作年限", "按20题计算等级", "按20题计算等级对应薪资中位数", "主管评价等级", "主管评价等级对应薪资中位数", "职称证书", "资格证书", "荣誉证书"], [[r["部门"], r["姓名"], fmt(r["工作年限"], 1), clean(r.get("规则重算评级")), SALARY.get(clean(r.get("规则重算评级")), ""), clean(r.get("主管主观评级")), SALARY.get(clean(r.get("主管主观评级")), ""), cert_display(r["职称证书"]), cert_display(r["资格证书"]), cert_display(r["荣誉证书"])] for r in unmatched])}

{review_section}

## 六、按等级人员证书列表

{table(["部门", "姓名", "工作年限", "按20题计算等级", "按20题计算等级对应薪资中位数", "主管评价等级", "主管评价等级对应薪资中位数", "职称证书", "资格证书", "荣誉证书"], detail)}

## 七、无任何证书人员按主管评价等级分布

{table(["主管评价等级", "人数", "部门及人员"], no_cert_by_sup)}

## 八、工作满2年仍无任何证书人员名单

{table(["部门", "姓名", "工作年限", "按20题计算等级", "按20题计算等级对应薪资中位数", "主管评价等级", "主管评价等级对应薪资中位数", "无证书判断", "职称证书", "资格证书", "荣誉证书"], [[r["部门"], r["姓名"], fmt(r["工作年限"], 1), clean(r.get("规则重算评级")), SALARY.get(clean(r.get("规则重算评级")), ""), clean(r.get("主管主观评级")), SALARY.get(clean(r.get("主管主观评级")), ""), "需关注（工作满2年仍无证书）", r["职称证书"], r["资格证书"], r["荣誉证书"]] for r in full2])}

## 九、工作未满2年且无任何证书人员名单

{table(["部门", "姓名", "工作年限", "按20题计算等级", "按20题计算等级对应薪资中位数", "主管评价等级", "主管评价等级对应薪资中位数", "无证书判断", "职称证书", "资格证书", "荣誉证书"], [[r["部门"], r["姓名"], fmt(r["工作年限"], 1), clean(r.get("规则重算评级")), SALARY.get(clean(r.get("规则重算评级")), ""), clean(r.get("主管主观评级")), SALARY.get(clean(r.get("主管主观评级")), ""), "正常（工作未满2年）", r["职称证书"], r["资格证书"], r["荣誉证书"]] for r in less2])}
"""


def grade_idx(grade: str) -> int:
    return GRADE_ORDER.index(grade) if grade in GRADE_ORDER else 999


def worse_grade(*grades: str) -> str:
    valid = [g for g in grades if g in GRADE_ORDER]
    return max(valid, key=grade_idx) if valid else ""


def option_is(value, letters: set[str]) -> bool:
    return option_letter(value) in letters


def grade_distribution_text(records, field: str) -> str:
    counts = Counter(clean(r.get(field)) for r in records)
    return "、".join(f"{g}{counts.get(g, 0)}" for g in GRADE_ORDER if counts.get(g, 0)) or "-"


def build_workload_overview(workload, supervisor):
    base = load_employee_base()
    base_names = set(base["姓名"].map(clean))
    wnames = set(workload["姓名"].map(clean))
    snames = {clean(r.get("成员姓名")) for r in supervisor}
    sup_by_name = {clean(r.get("成员姓名")): r for r in supervisor}
    wl_by_name = {clean(r.get("姓名")): r for _, r in workload.iterrows()}

    fill_rows = []
    eval_rows = []
    missing_workload_rows = []
    missing_eval_rows = []
    for dept in SCOPE_DEPTS:
        dept_people = base[base["部门"] == dept]["姓名"].map(clean).tolist()
        count = len(dept_people)
        missing_w = [n for n in dept_people if n not in wnames]
        missing_s = [n for n in dept_people if n not in snames]
        w_done = count - len(missing_w)
        s_done = count - len(missing_s)
        fill_rows.append([dept, count, w_done, len(missing_w), pct(w_done, count)])
        eval_rows.append([dept, count, s_done, len(missing_s), pct(s_done, count)])
        if missing_w:
            missing_workload_rows.append([dept, len(missing_w), "、".join(missing_w)])
        if missing_s:
            missing_eval_rows.append([dept, len(missing_s), "、".join(missing_s)])

    workload_load_counts = Counter(workload["自动工作负荷"].map(norm_load))
    attitude_counts = Counter(workload["您目前对自己岗位状态的总体态度更接近哪一种？"].map(clean))
    sup_load_counts = Counter(norm_load(r.get("13. 以目前的业务量，您认为该员工的工作忙闲程度如何？")) for r in supervisor)

    auto_counts = Counter(clean(r.get("自动生成评级")) for r in supervisor)
    sup_counts = Counter(clean(r.get("主管主观评级")) for r in supervisor)
    rule_counts = Counter(clean(r.get("规则重算评级")) for r in supervisor)
    auto_sup_diff = [r for r in supervisor if clean(r.get("自动生成评级")) != clean(r.get("主管主观评级"))]
    rule_sup_diff = [r for r in supervisor if clean(r.get("规则重算评级")) != clean(r.get("主管主观评级"))]

    dept_grade_rows = []
    for dept in SCOPE_DEPTS:
        sub = [r for r in supervisor if clean(r.get("成员部门")) == dept]
        if sub:
            dept_grade_rows.append([dept, len(sub), grade_distribution_text(sub, "主管主观评级")])

    salary_rows = []
    comparable = []
    for r in supervisor:
        sup_grade = clean(r.get("主管主观评级"))
        rule_grade = clean(r.get("规则重算评级"))
        if sup_grade in GRADE_ORDER and rule_grade in GRADE_ORDER:
            cmp_value = grade_idx(sup_grade) - grade_idx(rule_grade)
            relation = "主管主观评级与20题重算评级一致" if cmp_value == 0 else ("主管主观评级低于20题重算评级" if cmp_value > 0 else "主管主观评级高于20题重算评级")
            comparable.append((r, relation))
            if relation != "主管主观评级与20题重算评级一致":
                salary_rows.append([clean(r.get("成员部门")), clean(r.get("成员姓名")), rule_grade, sup_grade, relation, clean(r.get("薪资区间-最小值（绩效考核最低值）")), clean(r.get("薪资区间-最大值（绩效考核最高值）"))])
    salary_relation_counts = Counter(rel for _, rel in comparable)

    high_grade = [r for r in supervisor if clean(r.get("主管主观评级")) in {"S", "A", "B"}]
    weak_necessity = [r for r in high_grade if option_is(r.get("该成员所在岗位是否有存在的必要性"), {"B", "C", "D"})]
    weak_leave = [r for r in high_grade if option_is(r.get("20. 如果该员工明天提出辞职，你的第一反应是？"), {"C", "D", "E"})]
    both_weak = [r for r in weak_necessity if r in weak_leave]

    def role_contrast_rows(rows):
        return [[clean(r.get("成员姓名")), clean(r.get("成员部门")), clean(r.get("成员岗位")), clean(r.get("主管主观评级")), clean(r.get("该成员所在岗位是否有存在的必要性")) or "未填写", clean(r.get("20. 如果该员工明天提出辞职，你的第一反应是？")) or "未填写"] for r in rows]

    normal_ranges = {
        "S": (0, 0.10),
        "A": (0.10, 0.20),
        "B": (0.30, 0.50),
        "C": (0.20, 0.30),
        "D": (0.05, 0.15),
        "E": (0, 0.10),
        "F": (0, 0.10),
    }
    overall_dist_rows = []
    for g in GRADE_ORDER:
        count = sup_counts.get(g, 0)
        overall_dist_rows.append([g, count, pct(count, len(supervisor))])

    dept_normal_rows = []
    for dept in SCOPE_DEPTS:
        sub = [r for r in supervisor if clean(r.get("成员部门")) == dept]
        if not sub:
            continue
        if len(sub) < 3:
            dept_normal_rows.append([dept, len(sub), "样本过小，不作判断"])
            continue
        counts = Counter(clean(r.get("主管主观评级")) for r in sub)
        deviations = []
        for g, (low, high) in normal_ranges.items():
            ratio = counts.get(g, 0) / len(sub)
            if ratio < low:
                deviations.append(f"`{g}`偏低")
            elif ratio > high:
                deviations.append(f"`{g}`偏高")
        dept_normal_rows.append([dept, len(sub), "、".join(deviations) or "基本符合"])

    abnormal_fields = [
        ("岗位必要性偏弱", "该成员所在岗位是否有存在的必要性", {"B", "C", "D"}),
        ("离职影响偏弱", "20. 如果该员工明天提出辞职，你的第一反应是？", {"C", "D", "E"}),
        ("沟通偏弱", "8. 你认为该员工在和其他部门沟通事情时，效果如何？", {"C", "D"}),
        ("上手偏弱", "16. 面对新系统、新业务及新技能知识时，该员工的上手速度是？", {"C", "D"}),
        ("效率偏弱", "1.在日常工作中，该成员是否能做到高效、按时完成？", {"C", "D"}),
        ("优化偏弱", "7.他会主动思考如何把现有的工作做得更好吗？", {"C", "D"}),
        ("同岗偏弱", "18.与同岗位的平均水平相比，该员工的胜任力处于什么水平？", {"C", "D"}),
        ("忙闲偏弱", "13. 以目前的业务量，您认为该员工的工作忙闲程度如何？", {"D", "E"}),
    ]

    abnormal_records = []
    for r in supervisor:
        final_grade = worse_grade(clean(r.get("自动生成评级")), clean(r.get("主管主观评级")))
        items = []
        for label, field, letters in abnormal_fields:
            value = clean(r.get(field))
            if option_is(value, letters):
                items.append(f"`{label}：{value}`")
        abnormal_records.append((r, final_grade, items))
    b_plus_abnormal = [(r, fg, items) for r, fg, items in abnormal_records if fg in {"S", "A", "B"} and len(items) >= 3]
    c_abnormal = [(r, fg, items) for r, fg, items in abnormal_records if fg == "C" and len(items) >= 2]

    def abnormal_rows(records):
        records = sorted(records, key=lambda x: (-len(x[2]), grade_idx(x[1]), clean(x[0].get("成员部门")), clean(x[0].get("成员姓名"))))
        return [[clean(r.get("成员姓名")), clean(r.get("成员部门")), clean(r.get("自动生成评级")), clean(r.get("主管主观评级")), fg, len(items), "；".join(items)] for r, fg, items in records]

    low_load_rows = []
    for name, r in sup_by_name.items():
        w = wl_by_name.get(name)
        if w is None:
            continue
        final_grade = worse_grade(clean(r.get("自动生成评级")), clean(r.get("主管主观评级")))
        if final_grade not in {"S", "A", "B", "C"}:
            continue
        emp_load = clean(w.get("总体工作负荷"))
        auto_load = norm_load(w.get("自动工作负荷"))
        signals = []
        if "欠饱和" in emp_load or "非饱和" in emp_load:
            signals.append(f"`员工自评:{emp_load}`")
        if auto_load in {"欠饱和", "非饱和"}:
            signals.append(f"`自动生成:{auto_load}`")
        if signals:
            low_load_rows.append([name, clean(r.get("成员部门")), final_grade, clean(r.get("自动生成评级")), clean(r.get("主管主观评级")), emp_load, auto_load, "；".join(signals)])
    low_load_rows = sorted(low_load_rows, key=lambda x: (grade_idx(x[2]), x[1], x[0]))
    low_load_counts = Counter(row[2] for row in low_load_rows)

    return f"""# 机关工作量摸排报告-{DATE}

## 一、统计口径

本次分析依据最新 `职能部门工作量摸排表`、`部门成员工作主管评定`、`ERP使用情况.xlsx` 及相关规则文件。统计范围为当前表格中自动识别的单一部门或部门集合。

## 二、部门与员工数量

- 部门数量：`{len(SCOPE_DEPTS)}` 个
- 统计员工：`{TOTAL_STAFF}` 人

## 三、各部门工作摸排填报情况

当前 `{TOTAL_STAFF}` 名员工中，已完成工作摸排 `{len(workload)}` 人，未填报 `{TOTAL_STAFF-len(workload)}` 人，整体填报率 `{pct(len(workload), TOTAL_STAFF)}`。

{table(["部门", "员工人数", "已填报", "未填报", "填报率"], fill_rows)}

### 未填报工作摸排人员名单

{table(["部门", "未填报人数", "人员名单"], missing_workload_rows)}

## 四、各部门员工评价完成情况

当前 `{TOTAL_STAFF}` 名员工中，已完成主管评价 `{len(supervisor)}` 人，未完成 `{TOTAL_STAFF-len(supervisor)}` 人，整体完成率 `{pct(len(supervisor), TOTAL_STAFF)}`。

{table(["部门", "员工人数", "已评价", "未评价", "完成率"], eval_rows)}

### 未完成员工评价人员名单

{table(["部门", "未评价人数", "人员名单"], missing_eval_rows)}

## 五、工作负荷与岗位状态

### 1. 自动工作负荷分布

{table(["自动工作负荷", "人数"], [[k, workload_load_counts.get(k, 0)] for k in LOAD_ORDER if workload_load_counts.get(k, 0)])}

### 2. 员工岗位状态态度分布

{table(["岗位状态态度", "人数"], attitude_counts.most_common())}

### 3. 主管忙闲程度分布

{table(["主管忙闲程度", "人数"], [[k, sup_load_counts.get(k, 0)] for k in LOAD_ORDER if sup_load_counts.get(k, 0)])}

### 4. 本维度结论

1. 自动工作负荷中 `超负荷/满负荷` 合计 `{workload_load_counts.get("超负荷", 0) + workload_load_counts.get("满负荷", 0)}` 人，说明{scope_label()}岗位仍存在较明显承压面。
2. 自动工作负荷中 `欠饱和/非饱和` 合计 `{workload_load_counts.get("欠饱和", 0) + workload_load_counts.get("非饱和", 0)}` 人，后续需要与主管评级、ERP痕迹和岗位必要性叠加复核。
3. 主管忙闲判断中 `满负荷/超负荷` 合计 `{sup_load_counts.get("满负荷", 0) + sup_load_counts.get("超负荷", 0)}` 人，与员工填报形成互补口径。

## 六、ABCD 等级分布情况

### 1. 自动评级总体分布

{table(["等级", "人数"], [[g, auto_counts.get(g, 0)] for g in GRADE_ORDER])}

### 2. 主管主观评级总体分布

{table(["等级", "人数"], [[g, sup_counts.get(g, 0)] for g in GRADE_ORDER])}

### 3. 自动评级与主管主观评级一致性

{table(["一致性", "人数"], [["一致", len(supervisor) - len(auto_sup_diff)], ["不一致", len(auto_sup_diff)]])}

### 4. 自动评级与主管主观评级不一致人员名单

{table(["姓名", "部门", "自动生成评级", "主管主观评级"], [[clean(r.get("成员姓名")), clean(r.get("成员部门")), clean(r.get("自动生成评级")), clean(r.get("主管主观评级"))] for r in auto_sup_diff])}

### 5. 20题规则重算评级总体分布

{table(["等级", "人数"], [[g, rule_counts.get(g, 0)] for g in GRADE_ORDER])}

### 6. 20题规则重算评级与主管主观评级一致性

{table(["一致性", "人数"], [["一致", len(supervisor) - len(rule_sup_diff)], ["不一致", len(rule_sup_diff)]])}

### 7. 20题规则重算评级与主管主观评级不一致人员名单

{table(["姓名", "部门", "按20题计算等级", "主管主观评级"], [[clean(r.get("成员姓名")), clean(r.get("成员部门")), clean(r.get("规则重算评级")), clean(r.get("主管主观评级"))] for r in rule_sup_diff])}

### 8. 各部门主管主观评级结构

{table(["部门", "样本数", "主管主观评级分布"], dept_grade_rows)}

### 9. 本维度结论

1. 主管主观评级中 `S/A/B` 合计 `{sum(sup_counts.get(g, 0) for g in ["S", "A", "B"])}` 人，仍是当前评价结果的主体。
2. 自动评级与主管主观评级不一致 `{len(auto_sup_diff)}` 人，20题规则重算评级与主管主观评级不一致 `{len(rule_sup_diff)}` 人，两类差异名单应结合使用。
3. 后续涉及等级判断的报告均优先采用 `20题规则重算评级`，同时保留主管主观评级用于管理判断。

## 七、主管主观评级与20题重算评级一致性

本部分使用主管主观评级与 `20题规则重算评级` 进行对照，不使用 `主管评级对应公司等级` 字段。薪资区间最小值、最大值仅作为背景信息列示，不参与等级判断。

### 1. 总体结果

- 有主管主观评级：`{len([r for r in supervisor if clean(r.get("主管主观评级"))])}` 人
- 已形成20题重算评级、可直接比较：`{len(comparable)}` 人
- 未形成20题重算评级、暂无法比较：`{len(supervisor) - len(comparable)}` 人
- `主管主观评级与20题重算评级一致`：`{salary_relation_counts.get("主管主观评级与20题重算评级一致", 0)}`
- `主管主观评级高于20题重算评级`：`{salary_relation_counts.get("主管主观评级高于20题重算评级", 0)}`
- `主管主观评级低于20题重算评级`：`{salary_relation_counts.get("主管主观评级低于20题重算评级", 0)}`

### 2. 不一致人员名单

{table(["部门", "姓名", "20题重算评级", "主管主观评级", "差异判断", "主管评级薪资区间最小值", "主管评级薪资区间最大值"], salary_rows)}

### 3. 本维度结论

1. 当前可比较样本中，主管主观评级与20题重算评级一致的样本 `{salary_relation_counts.get("主管主观评级与20题重算评级一致", 0)}` 人。
2. 主管主观评级高于20题重算评级的样本 `{salary_relation_counts.get("主管主观评级高于20题重算评级", 0)}` 人，主管主观评级低于20题重算评级的样本 `{salary_relation_counts.get("主管主观评级低于20题重算评级", 0)}` 人。

## 八、岗位必要性、离职影响与评级反差

本部分仅筛选主管主观评级为 `S/A/B` 的样本。

### 1. 总体结果

- 高评级样本：`{len(high_grade)}`
- `岗位必要性偏弱但评级不低`：`{len(weak_necessity)}`
- `离职影响偏弱但评级不低`：`{len(weak_leave)}`
- `两项同时偏弱但评级不低`：`{len(both_weak)}`

### 2. 岗位必要性偏弱但评级不低样本

{table(["姓名", "部门", "岗位", "主管主观评级", "岗位必要性", "离职影响"], role_contrast_rows(weak_necessity))}

### 3. 离职影响偏弱但评级不低样本

{table(["姓名", "部门", "岗位", "主管主观评级", "岗位必要性", "离职影响"], role_contrast_rows(weak_leave))}

### 4. 本维度结论

1. 高评级中存在岗位必要性或离职影响偏弱样本，说明评级不能只看短期工作表现。
2. 同时偏弱的 `{len(both_weak)}` 人应优先纳入人岗适配复核。

## 九、各部门等级结构是否符合常态分布

本部分以主管主观评级对照常态模型：`S 0-10%`、`A 10-20%`、`B 30-50%`、`C 20-30%`、`D 5-15%`、`E/F 0-10%`。

### 1. 整体分布

{table(["等级", "人数", "占比"], overall_dist_rows)}

### 2. 分部门判断

{table(["部门", "样本数", "主要偏离情况"], dept_normal_rows)}

### 3. 本维度结论

1. 整体评级结构以 `B` 档为主峰，符合{scope_label()}员工评价的基本分布特征。
2. 样本数较小部门只作为提示，不单独作为结构异常判断依据。

## 十、取低档后的异常复核名单

本部分按 `自动生成评级` 与 `主管主观评级` 取较低档作为最终评价级别，并叠加关键异常项识别复核范围。

### 1. B类及以上且异常较多人员

{table(["姓名", "部门", "自动生成评级", "主管主观评级", "最终评价级别", "异常项数", "异常选项"], abnormal_rows(b_plus_abnormal))}

### 2. C类且异常更多人员

{table(["姓名", "部门", "自动生成评级", "主管主观评级", "最终评价级别", "异常项数", "异常选项"], abnormal_rows(c_abnormal))}

### 3. 本维度结论

1. `B类及以上且异常较多人员` 共 `{len(b_plus_abnormal)}` 人，应作为高等级边界复核对象。
2. `C类且异常更多人员` 共 `{len(c_abnormal)}` 人，可作为岗位适配和培养转化复核对象。

## 十一、低负荷信号与等级交叉复核范围

本部分将 `员工自评总体工作负荷` 与 `自动工作负荷` 中的 `欠饱和/非饱和` 视为低负荷信号，再与取低档后的最终评价级别交叉。

### 1. 总体结果

- 识别低负荷交叉复核样本：`{len(low_load_rows)}` 人
- `S/A/B/C` 分布：`{"、".join(f"{g}{low_load_counts.get(g, 0)}" for g in ["S", "A", "B", "C"] if low_load_counts.get(g, 0)) or "无"}`

### 2. 名单

{table(["姓名", "部门", "最终评价级别", "自动生成评级", "主管主观评级", "员工自评总体工作负荷", "自动工作负荷", "低负荷信号"], low_load_rows)}

### 3. 本维度结论

1. 低负荷信号不宜直接作为调档依据，但应作为等级复核的重要触发条件。
2. 等级不低且出现低负荷信号的人员，应结合 ERP 使用痕迹、岗位职责和部门实际分工复核。

## 十二、特殊人群与分层管理规则

### 1. E类管理含义

`E类` 不简单理解为“差”，而是指向基础性、支持性、可替代性较强的岗位或任务，后续可纳入人才沉淀池和岗位再配置管理。

### 2. 毕业两年内员工保护

对毕业两年内员工，不因短期表现偏弱被简单压为 `E/F`，原则上优先视为成长观察样本。

### 3. 顾问转岗年龄规则

对男性55岁以上、女性50岁以上员工，应同步考虑顾问、传帮带、经验支持等角色，不简单按高强度主责岗位衡量。

## 十三、综合结论与建议

1. 样本缺口仍主要来自未填报工作摸排和未完成主管评价人员，应优先补齐数据。
2. 等级判断应以 `20题规则重算评级` 为基础，同时将主管主观评级作为管理修正意见保留。
3. 低负荷信号、岗位必要性偏弱、离职影响偏弱和等级不一致人员，是后续复核的主要交叉范围。
4. 对 `B类及以上且异常较多人员`，建议逐人复核其岗位价值、实际工作负荷和可替代性。
5. 对 `C类且异常更多人员`，建议区分“需培养提升”和“岗位可优化合并”两类处理。
6. 对毕业两年内、年龄偏大顾问型人员等特殊样本，应先套用分层管理规则，再作最终等级判断。
"""


def valid_text(v) -> bool:
    text = clean(v)
    if not text:
        return False
    bad_exact = {"无", "暂无", "没有", "无建议", "无意见", "不涉及", "N/A", "NA", "无不妥", "/", "无。"}
    bad_contains = ["暂时未想到", "暂时没想到", "暂时没有想到", "暂无建议", "没有觉得不妥"]
    return text.upper() not in bad_exact and not any(item in text for item in bad_contains)


def build_idea_mom_report(workload, supervisor):
    sup_grade = {clean(r.get("成员姓名")): clean(r.get("规则重算评级")) or clean(r.get("主管主观评级")) for r in supervisor}

    def initial_grade(name: str) -> str:
        return sup_grade.get(clean(name)) or "未评级"

    idea_fields = [
        "您认为最应取消、合并或简化的工作是什么？",
        "您认为最值得推动但尚未落地的优化建议是什么？",
        "问卷中提出的问题您觉得有不妥的情况么，若有请列明",
        "请列出对于本问卷中您觉得需要优化的问题和建议",
    ]
    idea_people = []
    for _, r in workload.iterrows():
        name = clean(r.get("姓名"))
        details = [(field, clean(r.get(field))) for field in idea_fields if valid_text(r.get(field))]
        if details:
            idea_people.append({
                "部门": clean(r.get("部门")),
                "姓名": name,
                "等级": initial_grade(name),
                "明细": details,
            })

    dept_order = {dept: i for i, dept in enumerate(SCOPE_DEPTS)}
    idea_people.sort(key=lambda r: (dept_order.get(r["部门"], 99), grade_idx(r["等级"]), r["姓名"]))
    by_dept = defaultdict(list)
    for item in idea_people:
        by_dept[item["部门"]].append(item)

    idea_lines = []
    if not idea_people:
        idea_lines.append("当前未识别出有效主观想法人员。")
    else:
        for dept in sorted(by_dept, key=lambda d: dept_order.get(d, 99)):
            idea_lines.append(f"### {dept}")
            idea_lines.append("")
            for item in by_dept[dept]:
                idea_lines.append(f"- `{item['姓名']}`")
                idea_lines.append(f"  初步评价：`{item['等级']}`")
                for field, text in item["明细"]:
                    idea_lines.append(f"  {field}：{text}")
                idea_lines.append("")

    mom_field = "若公司推行针对学龄期员工的‘妈妈岗’照顾政策，您是否有意向参与"
    mom_people = []
    for _, r in workload.iterrows():
        intent = clean(r.get(mom_field))
        if option_letter(intent) in {"A", "B"}:
            name = clean(r.get("姓名"))
            mom_people.append({
                "部门": clean(r.get("部门")),
                "姓名": name,
                "等级": initial_grade(name),
                "意向": intent,
                "婚育": clean(r.get("您的婚育情况")),
                "照护": clean(r.get("您目前是否承担较重家庭照护责任")),
                "子女": clean(r.get("您家中是否有12岁以下子女需要较多照护")),
                "配偶": clean(r.get("是否因配偶同在公司且因公无法居家，导致您需独自承担较多家庭责任")),
                "优化": clean(r.get("如果有岗位优化或调整机会，您更倾向于哪一种工作状态？")),
            })
    mom_people.sort(key=lambda r: (dept_order.get(r["部门"], 99), grade_idx(r["等级"]), r["姓名"]))

    mom_lines = []
    if not mom_people:
        mom_lines.append("当前未识别出对“妈妈岗”表达明确意向或初步意向的人员。")
    else:
        for item in mom_people:
            mom_lines.extend([
                f"- `{item['部门']} - {item['姓名']}`",
                f"  初步评价：`{item['等级']}`",
                f"  妈妈岗意向：{item['意向']}",
                f"  您的婚育情况：{item['婚育'] or '未填写'}",
                f"  您目前是否承担较重家庭照护责任：{item['照护'] or '未填写'}",
                f"  您家中是否有12岁以下子女需要较多照护：{item['子女'] or '未填写'}",
                f"  是否因配偶同在公司且因公无法居家，导致您需独自承担较多家庭责任：{item['配偶'] or '未填写'}",
                f"  如果有岗位优化或调整机会，您更倾向于哪一种工作状态？：{item['优化'] or '未填写'}",
                "",
            ])

    return f"""# 有想法人员及妈妈岗意向名单-{DATE}

## 一、有主观想法人员名单

本次在工作摸排开放题中，共识别出 `{len(idea_people)}` 名存在有效主观想法的员工。以下按 `部门 + 初步评价等级 + 人员 + 想法明细` 列示。

{chr(10).join(idea_lines).rstrip()}

## 二、妈妈岗有意向人员名单

当前共有 `{len(mom_people)}` 名员工对“妈妈岗”表达出明确意向或初步意向，名单如下。

{chr(10).join(mom_lines).rstrip()}
"""


def build_feedback(workload, supervisor):
    sup_grade = {clean(r.get("成员姓名")): clean(r.get("规则重算评级")) or clean(r.get("主管主观评级")) for r in supervisor}

    def initial_grade(name: str) -> str:
        return sup_grade.get(clean(name)) or "未评级"

    idea_fields = [
        ("员工认为最应取消、合并或简化的工作", "您认为最应取消、合并或简化的工作是什么？"),
        ("员工认为最值得推动但尚未落地的优化建议", "您认为最值得推动但尚未落地的优化建议是什么？"),
        ("员工认为问卷中存在的不妥情况", "问卷中提出的问题您觉得有不妥的情况么，若有请列明"),
        ("员工对问卷需要优化的问题和建议", "请列出对于本问卷中您觉得需要优化的问题和建议"),
    ]

    idea_by_person = {}
    idea_section_rows = {label: [] for label, _ in idea_fields}
    for _, r in workload.iterrows():
        name = clean(r.get("姓名"))
        dept = clean(r.get("部门"))
        grade = initial_grade(name)
        details = []
        for label, field in idea_fields:
            text = clean(r.get(field))
            if valid_text(text):
                details.append(f"{label}：{text}")
                idea_section_rows[label].append([dept, grade, name, text])
        if details:
            idea_by_person[name] = {"部门": dept, "等级": grade, "明细": "<br>".join(details)}

    high_no_idea = []
    for r in supervisor:
        grade = clean(r.get("规则重算评级")) or clean(r.get("主管主观评级"))
        name = clean(r.get("成员姓名"))
        if grade in {"S", "A", "B"} and name not in idea_by_person:
            high_no_idea.append([clean(r.get("成员部门")), grade, name])
    other_with_idea = [
        [v["部门"], v["等级"], name, v["明细"]]
        for name, v in idea_by_person.items()
        if v["等级"] not in {"S", "A", "B"}
    ]

    mom_rows = []
    for _, r in workload.iterrows():
        intent = clean(r.get("若公司推行针对学龄期员工的‘妈妈岗’照顾政策，您是否有意向参与"))
        if option_letter(intent) in {"A", "B"}:
            name = clean(r.get("姓名"))
            mom_rows.append([clean(r.get("部门")), initial_grade(name), name, clean(r.get("您的婚育情况")), clean(r.get("您目前是否承担较重家庭照护责任")), clean(r.get("您家中是否有12岁以下子女需要较多照护")), intent])

    family_fields = [
        "您的婚育情况",
        "您目前是否承担较重家庭照护责任",
        "您家中是否有12岁以下子女需要较多照护",
        "是否因配偶同在公司且因公无法居家，导致您需独自承担较多家庭责任",
        "您是否有弹性工作、稳定作息、兼顾家庭的岗位需求",
        "若公司推行针对学龄期员工的‘妈妈岗’照顾政策，您是否有意向参与",
        "如果有岗位优化或调整机会，您更倾向于哪一种工作状态？",
    ]
    family_rows = []
    for field in family_fields:
        counts = Counter((clean(v) or "未填写") for v in workload[field])
        for option, count in counts.most_common():
            family_rows.append([field, option, count])

    optimize_field = "如果有岗位优化或调整机会，您更倾向于哪一种工作状态？"
    optimize_counter = Counter()
    for _, r in workload.iterrows():
        option = clean(r.get(optimize_field)) or "未填写"
        optimize_counter[(initial_grade(clean(r.get("姓名"))), option)] += 1
    optimize_rows = [[grade, option, count] for (grade, option), count in sorted(optimize_counter.items(), key=lambda x: (grade_idx(x[0][0]) if x[0][0] != "未评级" else -1, x[0][1]))]

    sup_rows = []
    for r in supervisor:
        details = []
        for label, field in [
            ("工作评价", "工作评价"),
            ("问卷不妥", "问卷中提出的问题您觉得有不妥的情况么，若有请列明"),
            ("问卷优化建议", "请列出对于本问卷中您觉得需要优化的问题和建议"),
        ]:
            text = clean(r.get(field))
            if valid_text(text):
                details.append(f"{label}：{text}")
        if details:
            grade = clean(r.get("规则重算评级")) or clean(r.get("主管主观评级"))
            sup_rows.append([clean(r.get("成员部门")), grade, clean(r.get("成员姓名")), "<br>".join(details)])

    summary_rows = []
    for label, field in idea_fields:
        summary_rows.append(["员工主观想法", field, len(idea_section_rows[label])])
    summary_rows.extend([
        ["主管评价反馈", "工作评价", sum(1 for r in supervisor if valid_text(r.get("工作评价")))],
        ["主管评价反馈", "问卷中提出的问题您觉得有不妥的情况么，若有请列明", sum(1 for r in supervisor if valid_text(r.get("问卷中提出的问题您觉得有不妥的情况么，若有请列明")))],
        ["主管评价反馈", "请列出对于本问卷中您觉得需要优化的问题和建议", sum(1 for r in supervisor if valid_text(r.get("请列出对于本问卷中您觉得需要优化的问题和建议")))],
        ["家庭照护与岗位意向", "若公司推行针对学龄期员工的‘妈妈岗’照顾政策，您是否有意向参与", len(mom_rows)],
    ])

    md = f"""# 重点问题反馈汇总-{DATE}

本稿整合员工工作摸排主观反馈、主管补充评价以及“妈妈岗”相关意向名单，统一按表格展示。

## 一、汇总

{table(["类别", "问题", "有效人数"], summary_rows)}

## 二、分层关注名单

### S/A/B但未提出任何员工建议或问卷意见的人员

{table(["部门", "初步评价", "人员"], high_no_idea)}

### 其他等级但提出员工建议或问卷意见的人员

{table(["部门", "初步评价", "人员", "主观想法明细"], other_with_idea)}

## 三、员工主观想法名单

### 员工认为最应取消、合并或简化的工作

{table(["部门", "初步评价", "人员", "想法明细"], idea_section_rows["员工认为最应取消、合并或简化的工作"])}

### 员工认为最值得推动但尚未落地的优化建议

{table(["部门", "初步评价", "人员", "想法明细"], idea_section_rows["员工认为最值得推动但尚未落地的优化建议"])}

### 员工认为问卷中存在的不妥情况

{table(["部门", "初步评价", "人员", "想法明细"], idea_section_rows["员工认为问卷中存在的不妥情况"])}

### 员工对问卷需要优化的问题和建议

{table(["部门", "初步评价", "人员", "想法明细"], idea_section_rows["员工对问卷需要优化的问题和建议"])}

## 四、主管评价补充反馈

{table(["部门", "初步评价", "人员", "主管评价反馈明细"], sup_rows)}

## 五、妈妈岗有意向人员名单

{table(["部门", "初步评价", "姓名", "婚育情况", "家庭照护责任", "12岁以下子女照护", "妈妈岗意向"], mom_rows)}

## 六、家庭照护与岗位优化倾向统计

{table(["问题", "选项", "人数"], family_rows)}

## 七、岗位优化或调整机会下的工作状态倾向

{table(["等级", "工作状态倾向", "人数"], optimize_rows)}
"""

    def count_by_grade(rows: list[list]) -> list[list]:
        counts = Counter(row[1] for row in rows)
        order = ["未评级", *GRADE_ORDER]
        return [[g, counts.get(g, 0)] for g in order if counts.get(g, 0)]

    family_counter = {
        field: Counter((clean(v) or "未填写") for v in workload[field])
        for field in family_fields
    }
    mom_counter = family_counter["若公司推行针对学龄期员工的‘妈妈岗’照顾政策，您是否有意向参与"]
    mom_intent_count = len(mom_rows)
    optimize_total_rows = [[option, count] for option, count in Counter((clean(v) or "未填写") for v in workload[optimize_field]).most_common()]
    optimize_grade_counts = Counter()
    for grade, _, count in optimize_rows:
        optimize_grade_counts[grade] += count

    cancel_dept_rows = [
        ["技术中心", "会议与考核耗时、审批流程精简、重复性工作较多、前期资料不完整导致返工"],
        ["资产财务部", "ERP流程复杂、付款与报销流程可合并、资料传递与保管可电子化"],
        ["南京永利重工制造有限公司", "工作日志与任务重复、跨部门协调链条长、缺少统一提交格式"],
        ["综合办公室", "无纸化报销、用印和证书借用流程简化"],
        ["信息化中心", "跨部门需求沟通、系统需求确认和流程文档沉淀"],
    ]
    typical_suggestion_rows = [
        ["技术中心", "材料代码核对和质检自动化、AI处理设计录入、按职能定位建立更专业的分工体系"],
        ["资产财务部", "减少兼岗、推进居家办公和弹性上下班、推动 ERP 文档化管理"],
        ["信息化中心", "建立跨部门标准化流程文档，减少沟通返工"],
        ["南京永利重工制造有限公司", "建立统一的标准化提交流程、责任清单和唯一对接人机制"],
        ["综合办公室", "推动电子化流转、减少线下签批和重复资料提交"],
    ]
    cancel_dept_rows = [row for row in cancel_dept_rows if row[0] in SCOPE_DEPTS]
    typical_suggestion_rows = [row for row in typical_suggestion_rows if row[0] in SCOPE_DEPTS]
    cancel_dept_section = ""
    if cancel_dept_rows:
        cancel_dept_section = f"""
典型部门与反馈方向如下：

{table(["部门", "主要反馈方向"], cancel_dept_rows)}
"""
    typical_suggestion_section = ""
    if typical_suggestion_rows:
        typical_suggestion_section = f"""
其中较为典型的建议包括：

{table(["部门", "典型建议方向"], typical_suggestion_rows)}
"""
    questionnaire_focus_rows = [
        ["加班原因", "现有选项覆盖不够完整，建议增加指导、答疑、协助等消耗项"],
        ["流程审批项", "流程审批类工作性质的下拉选项数量不足，难以覆盖实际场景"],
        ["部门工作分类", "部分分类与实际业务不够贴合，建议按部门工作类型补充选项"],
        ["问卷名词", "个别名词偏专业，建议用更直白的表述降低理解差异"],
        ["填报方式", "建议增加电子签名、附件上传或说明补充能力"],
    ]
    cancel_count = len(idea_section_rows["员工认为最应取消、合并或简化的工作"])
    suggestion_count = len(idea_section_rows["员工认为最值得推动但尚未落地的优化建议"])
    questionnaire_bad_count = len(idea_section_rows["员工认为问卷中存在的不妥情况"])
    questionnaire_opt_count = len(idea_section_rows["员工对问卷需要优化的问题和建议"])
    cancel_summary = "本轮未识别出有效反馈。" if cancel_count == 0 else f"""主要集中在：

- `流程审批、ERP流转、报销类流程`：如简化 ERP 流程、合并付款及报销流转、减少纸质流转、减少重复接口。
- `重复性统计、台账、文档整理类工作`：如安全台账统计、文档整理、日志与任务重复填报、资料传递与保管。
- `跨部门协调和资料补要`：如前期资料不完整导致后续反复沟通、审批链条长、协调节点多。

分等级看，有效反馈主要集中在：

{table(["等级", "有效反馈人数"], count_by_grade(idea_section_rows["员工认为最应取消、合并或简化的工作"]))}

{cancel_dept_section}"""
    suggestion_summary = "本轮未识别出有效反馈。" if suggestion_count == 0 else f"""主要方向包括：

- `ERP及系统化优化`
- `跨部门标准化流程与唯一对接人机制`
- `AI或自动化在录入、核对、整理中的应用`
- `岗位职责明晰和专业化分工`
- `弹性上下班、居家办公、打卡地点优化`

分等级看，有效反馈主要集中在：

{table(["等级", "有效反馈人数"], count_by_grade(idea_section_rows["员工认为最值得推动但尚未落地的优化建议"]))}

{typical_suggestion_section}"""
    questionnaire_bad_summary = "本轮未识别出员工对问卷本身提出明确“不妥”意见。" if questionnaire_bad_count == 0 else """主要意见包括：

- 部分问题选项覆盖不全
- 部分问题与实际部门工作类型不匹配
- 个别问题涉及个人信息，建议改为选填
- 部分表述不够直白，个别名词较专业化，不易理解"""
    questionnaire_opt_summary = "本轮未识别出员工对问卷提出明确优化建议。" if questionnaire_opt_count == 0 else f"""主要集中在补充选项、优化题目表述、增加电子签名上传等方向。

具体较集中的反馈有：

{table(["反馈主题", "集中意见"], questionnaire_focus_rows)}"""
    mom_cross_summary = """- 已婚已育及存在家庭照护责任的人员中，本轮未出现 `A.有强烈意向` 或 `B.有意向，但是不迫切，需考虑` 的明确表达。
- 当前更多表现为 `未填写` 或 `无意向`，尚未形成可直接支撑岗位政策需求的样本。
- 家庭照护压力可以作为持续观察项，但本轮不能据此推导出“妈妈岗”明确需求规模。""" if mom_intent_count == 0 else """- 已婚已育人员中，确有少量员工表现出 `强烈意向` 或 `有意向但需考虑`。
- 但更大比例仍然是 `未填写` 或 `无意向`。
- 家庭照护压力和“妈妈岗”意向之间存在关联，但尚不足以直接推导出明确岗位政策需求规模。"""
    mom_judgement = """- `本轮未识别出明确意向人员`
- `仅具备持续观察意义`
- 后续如推进政策，应先通过访谈或补充调研确认真实需求，避免仅依据前置家庭照护条件直接扩大政策口径""" if mom_intent_count == 0 else """- `已具备观察意义`
- `尚不足以直接推导出明确岗位政策需求规模`
- 后续如推进政策，应优先对明确有意向人员进行一对一访谈，避免仅依据问卷选项直接扩大政策口径"""
    mom_final_note = "家庭照护与“妈妈岗”相关问题当前更多体现为观察项，本轮未识别出明确有意向人员，不宜直接放大为岗位政策需求。" if mom_intent_count == 0 else "家庭照护与“妈妈岗”相关问题当前更多体现为观察项，明确有意向人员规模较小，不宜直接放大为整体性岗位政策需求。"

    conclusion = f"""# 重点问题反馈结论分析-{DATE}

## 一、总体情况

本次基于最新工作摸排表与主管评价表，对{scope_label()}人员的开放题反馈进行整理。重点关注员工在工作摸排中的主观想法、主管在评价中的文字反馈，以及家庭照护和岗位优化意向等内容。

{scope_label()}内：

- 纳入员工总人数 `{TOTAL_STAFF}` 人
- 已完成工作摸排 `{len(workload)}` 人
- 已完成主管评价 `{len(supervisor)}` 人

在员工主观反馈中，已将 `无`、`暂无`、`暂无建议`、`暂时没想到`、`没有觉得不妥`、`/` 等表述统一视为“无有效想法”，不计入有效反馈人数。

## 二、员工主观想法情况

### 1. 员工最应取消、合并或简化的工作

当前共有 `{cancel_count}` 人提出了有效想法，{cancel_summary}

### 2. 员工最值得推动但尚未落地的优化建议

当前共有 `{suggestion_count}` 人提出了有效想法，{suggestion_summary}

### 3. 员工对问卷本身的不妥反馈和优化建议

对问卷本身提出明确“不妥”意见的共有 `{questionnaire_bad_count}` 人，{questionnaire_bad_summary}

对问卷提出明确优化建议的共有 `{questionnaire_opt_count}` 人，{questionnaire_opt_summary}

## 三、主管评价反馈情况

主管侧文字反馈主要分三类：

{table(["类别", "有效反馈人数"], [["工作评价", sum(1 for r in supervisor if valid_text(r.get("工作评价")))], ["问卷不妥情况", sum(1 for r in supervisor if valid_text(r.get("问卷中提出的问题您觉得有不妥的情况么，若有请列明")))], ["问卷优化建议", sum(1 for r in supervisor if valid_text(r.get("请列出对于本问卷中您觉得需要优化的问题和建议")))]] )}

其中：

- `工作评价` 基本已形成较高覆盖，可作为后续逐人判断的重要文字依据。
- 主管对问卷本身的不妥反馈数量不多，但与员工侧意见有一定一致性，主要还是集中在 `选项覆盖不足`、`问题表述不够贴合实际工作`。
- 主管对问卷优化提出的明确建议较少，说明主管更偏向于完成评价本身，而非主动参与问卷结构设计。

因此，主管侧信息更适合用于：

- 补充员工个人工作表现的定性描述
- 识别员工工作表现与主观想法之间是否存在反差
- 作为后续逐人复核的背景材料

## 四、家庭照护与“妈妈岗”相关情况

### 1. 前置条件填写情况

工作摸排中与家庭照护相关的前置条件填写情况如下：

{table(["问题", "主要分布"], [[field, "、".join(f"`{k} {v}`" for k, v in family_counter[field].most_common())] for field in family_fields if field not in {optimize_field, "若公司推行针对学龄期员工的‘妈妈岗’照顾政策，您是否有意向参与"}])}

整体看，这一组问题中，婚育情况和家庭照护责任填写较完整；涉及配偶是否为公司家属及因公无法居家导致家庭责任变化的问题，仍有较多 `未填写`，说明该类问题相对敏感。

### 2. “妈妈岗”意向情况

关于 `若公司推行针对学龄期员工的“妈妈岗”照顾政策，您是否有意向参与`：

{table(["选项", "人数"], [[k, v] for k, v in mom_counter.most_common()])}

当前直接表达出明确意向的人数为 `{mom_intent_count}` 人，其中 `A.有强烈意向` `{mom_counter.get("A.有强烈意向", 0)}` 人，`B.有意向，但是不迫切，需考虑` `{mom_counter.get("B.有意向，但是不迫切，需考虑", 0)}` 人。更多人员为未填写或无意向，说明这类政策问题在现阶段仍偏观察项。

### 3. 前置条件与“妈妈岗”意向交叉判断

从交叉结果看：

{mom_cross_summary}

因此，这一块现阶段更适合判断为：

{mom_judgement}

## 五、岗位优化或调整机会下的工作状态倾向

对 `如果有岗位优化或调整机会，您更倾向于哪一种工作状态？` 的填写情况看，当前主要集中在以下几类：

{table(["工作状态倾向", "人数"], optimize_total_rows)}

按等级分布看：

{table(["等级", "记录数"], [[g, optimize_grade_counts.get(g, 0)] for g in ["未评级", *GRADE_ORDER] if optimize_grade_counts.get(g, 0)])}

总体判断是：

1. `稳定保持型` 和 `挑战发展型` 是当前两类主流倾向。
2. `职责边界更清晰` 也是一类较明显诉求，说明部分员工并不单纯追求降压，而是希望责任边界更明确。
3. 从等级看，各等级内部均存在不同岗位倾向，后续如推进岗位优化，不宜按单一假设处理。

## 六、综合判断

1. 当前员工开放题反馈中，真正有价值的信息主要集中在 `流程精简`、`ERP及审批优化`、`减少重复统计与纸面流转`、`跨部门对接标准化` 四条线上。
2. 员工对于调研问卷本身的反馈总体不多，但方向较一致，主要问题在于 `选项覆盖不全`、`表述不够直白`、`与实际业务匹配度不足`。
3. 主管侧的 `工作评价` 已具有较高参考价值，后续更适合作为逐人分析时的定性支撑，而不是单独做大规模统计。
4. {mom_final_note}
5. 关于岗位优化后的工作状态倾向，员工内部并未呈现单一取向，而是同时存在 `稳定型` 与 `发展型` 两大群体，后续如推进岗位优化，应结合等级、负荷和岗位必要性分层处理。
"""
    return md, conclusion


def build_conclusions(workload_records, supervisor):
    c_load = Counter(r.get("自动层级") for r in workload_records)
    c_grade = Counter(r.get("规则重算评级") for r in supervisor)
    return f"""# 机关工作量摸排结论分析-{DATE}

## 一、总体结论

本轮{scope_label()}摸排样本为：工作量摸排 `{len(workload_records)}` 人，主管评价 `{len(supervisor)}` 人。自动工作负荷中 `超负荷` `{c_load.get('超负荷',0)}` 人、`满负荷` `{c_load.get('满负荷',0)}` 人，说明{scope_label()}岗位仍存在一定承压面；同时也存在 `欠饱和/非饱和` 样本，需要结合岗位职责和ERP实际承担度复核。

## 二、评价结论

20题规则重算后，S/A/B合计 `{sum(c_grade.get(g,0) for g in ['S','A','B'])}` 人，C/D/E/F合计 `{sum(c_grade.get(g,0) for g in ['C','D','E','F'])}` 人。高等级人员主要作为关键岗位和项目指导型骨干观察，低等级人员应结合证书、ERP、工作负荷和主管原文评价开展岗位适配复核。

## 三、后续建议

1. 对自动超负荷且主管也认定满负荷以上人员，优先开展流程减负和资源配置复核。
2. 对工作满2年仍无证书且评级偏高人员，按人力资源复核后的名单收紧证书维度。
3. 对员工反馈中反复出现的重复填报、流程审批、职责边界问题，建立专项整改台账。
"""


def build_grade_heatmap(supervisor) -> Path:
    from PIL import Image, ImageDraw, ImageFont

    out_dir = ROOT / "outputs"
    out_dir.mkdir(exist_ok=True)
    path = out_dir / f"各部门等级热力矩阵对比-{DATE}.png"
    depts = [d for d in SCOPE_DEPTS if any(clean(r.get("成员部门")) == d for r in supervisor)]

    def matrix(field: str):
        rows: list[list[int]] = []
        for dept in depts:
            sub = [r for r in supervisor if clean(r.get("成员部门")) == dept]
            counts = Counter(clean(r.get(field)) for r in sub)
            rows.append([counts.get(g, 0) for g in GRADE_ORDER])
        return rows

    panels = [
        ("主管原始评价等级分布", matrix("主管主观评级")),
        ("规则重算评级分布", matrix("规则重算评级")),
    ]
    vmax = max([value for _, data in panels for row in data for value in row] + [1])

    font_path = Path("/System/Library/Fonts/STHeiti Light.ttc")
    if not font_path.exists():
        font_path = Path("/System/Library/Fonts/Supplemental/Arial Unicode.ttf")
    title_font = ImageFont.truetype(str(font_path), 30)
    sub_font = ImageFont.truetype(str(font_path), 22)
    label_font = ImageFont.truetype(str(font_path), 18)
    small_font = ImageFont.truetype(str(font_path), 16)

    width, height = 1800, 820
    img = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(img)

    def text_center(text: str, box, font, fill="#111827"):
        bbox = draw.textbbox((0, 0), text, font=font)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        x0, y0, x1, y1 = box
        draw.text((x0 + (x1 - x0 - tw) / 2, y0 + (y1 - y0 - th) / 2), text, font=font, fill=fill)

    def blue(value: int) -> tuple[int, int, int]:
        if value <= 0:
            return (242, 248, 252)
        ratio = value / vmax
        start = (222, 238, 247)
        end = (36, 96, 160)
        return tuple(int(start[i] + (end[i] - start[i]) * ratio) for i in range(3))

    text_center("各部门等级区间热力分布对比", (0, 12, width, 58), title_font)
    panel_w = 760
    top = 105
    lefts = [80, 940]
    cell_w = 62
    cell_h = 44
    name_w = 260
    for left, (title, data) in zip(lefts, panels):
        text_center(title, (left, 66, left + panel_w, 98), sub_font)
        for j, grade in enumerate(GRADE_ORDER):
            text_center(grade, (left + name_w + j * cell_w, top - 34, left + name_w + (j + 1) * cell_w, top - 4), label_font)
        for i, dept in enumerate(depts):
            y = top + i * cell_h
            draw.text((left, y + 10), dept, font=small_font, fill="#111827")
            for j, value in enumerate(data[i]):
                x = left + name_w + j * cell_w
                draw.rectangle([x, y, x + cell_w, y + cell_h], fill=blue(value), outline="white", width=2)
                if value:
                    fill = "white" if value / vmax > 0.55 else "#111827"
                    text_center(str(value), (x, y, x + cell_w, y + cell_h), label_font, fill=fill)
        draw.text((left + name_w + 170, top + len(depts) * cell_h + 24), "等级区间", font=label_font, fill="#111827")
    img.save(path)
    return path


def build_integrated(paths, workload_records, supervisor):
    matrix_path = build_grade_heatmap(supervisor)
    base = load_employee_base()
    base_names = set(base["姓名"].map(clean))
    workload_names = {clean(r.get("姓名")) for r in workload_records}
    supervisor_names = {clean(r.get("成员姓名")) for r in supervisor}
    missing_workload = base[~base["姓名"].isin(workload_names)]
    missing_supervisor = base[~base["姓名"].isin(supervisor_names)]

    def missing_rows(df: pd.DataFrame) -> list[list]:
        rows = []
        for dept in SCOPE_DEPTS:
            names = df[df["部门"].eq(dept)]["姓名"].map(clean).tolist()
            if names:
                rows.append([dept, len(names), "、".join(names)])
        return rows or [["-", 0, "无"]]

    def names_text(records, name_col="成员姓名", dept_col="成员部门") -> str:
        grouped = defaultdict(list)
        for r in records:
            name = clean(r.get(name_col))
            dept = clean(r.get(dept_col))
            if name:
                grouped[dept].append(name)
        parts = []
        for dept in SCOPE_DEPTS:
            names = grouped.get(dept, [])
            if names:
                parts.append(f"{dept}：{'、'.join(names)}")
        return "；".join(parts) or "-"

    def dist_rows(field: str) -> list[list]:
        rows = []
        for grade in GRADE_ORDER:
            sub = [r for r in supervisor if clean(r.get(field)) == grade]
            if sub:
                rows.append([f"`{grade}`", f"`{len(sub)}`", f"`{names_text(sub)}`"])
        return rows

    auto_rule_diff = [r for r in supervisor if clean(r.get("自动生成评级")) != clean(r.get("规则重算评级"))]
    auto_rule_preview = "、".join(
        f"{clean(r.get('成员姓名'))}（表内{clean(r.get('自动生成评级')) or '-'} / 重算{clean(r.get('规则重算评级')) or '-'}）"
        for r in auto_rule_diff[:20]
    )
    if len(auto_rule_diff) > 20:
        auto_rule_preview += f"等{len(auto_rule_diff)}人"

    detail_rows = []
    for r in sorted(supervisor, key=lambda x: (SCOPE_DEPTS.index(clean(x.get("成员部门"))) if clean(x.get("成员部门")) in SCOPE_DEPTS else 99, grade_idx(clean(x.get("规则重算评级"))), clean(x.get("成员姓名")))):
        rule_grade = clean(r.get("规则重算评级"))
        sup_grade = clean(r.get("主管主观评级"))
        detail_rows.append([
            clean(r.get("成员部门")),
            clean(r.get("成员姓名")),
            fmt(r.get("汇总分值")),
            fmt(r.get("纯20题总分")),
            fmt(r.get("岗位必要性修正分")),
            fmt(r.get("规则重算总分")),
            clean(r.get("自动生成评级")),
            rule_grade,
            sup_grade,
            clean(r.get("规则匹配情况")),
        ])

    diff_by_dept = []
    for dept in SCOPE_DEPTS:
        sub = []
        for r in supervisor:
            if clean(r.get("成员部门")) != dept:
                continue
            rule_grade = clean(r.get("规则重算评级"))
            sup_grade = clean(r.get("主管主观评级"))
            if rule_grade and sup_grade and rule_grade != sup_grade:
                sub.append(f"{clean(r.get('成员姓名'))}（重算{fmt(r.get('规则重算总分'))}分，规则{rule_grade}/主管{sup_grade}）")
        if sub:
            diff_by_dept.append([dept, "、".join(sub)])

    def empty_cert(value) -> bool:
        text = clean(value).replace("；", "").replace(";", "").replace("、", "").replace("/", "").strip()
        return text in {"", "无", "没", "没有", "暂无", "无无"}

    def cert_value(value) -> str:
        return "" if empty_cert(value) else clean(value)

    def merge_cert(*values) -> str:
        seen = set()
        merged = []
        for value in values:
            text = cert_value(value)
            if not text:
                continue
            for part in re.split(r"[；;、,，]", text):
                part = cert_value(part)
                if part and part not in seen:
                    seen.add(part)
                    merged.append(part)
        return "；".join(merged)

    roster = pd.read_excel(ROOT / "ERP使用情况.xlsx", sheet_name="职能部门人员台账")
    roster["姓名"] = roster["姓名"].map(clean)
    roster_start = {}
    roster_honor = {}
    for _, item in roster.iterrows():
        name = clean(item.get("姓名"))
        if not name:
            continue
        if name not in roster_start:
            roster_start[name] = item.get("参加工作时间") if not pd.isna(item.get("参加工作时间")) else item.get("到本单位时间")
        roster_honor[name] = merge_cert(
            roster_honor.get(name, ""),
            item.get("优秀情况（文本化）"),
            item.get("安全先进个人情况（文本化）"),
        )

    yearend_map = {}
    yearend_file = latest_optional_file("年终总结", ".xlsx")
    if yearend_file:
        yearend = pd.read_excel(yearend_file, sheet_name="Sheet1", header=1)
        yearend["姓名"] = yearend["姓名(必填)"].map(clean)
        for _, item in yearend.iterrows():
            name = clean(item.get("姓名"))
            if name and name not in yearend_map:
                yearend_map[name] = {
                    "职称证书": cert_value(item.get("有效职称证书")),
                    "资格证书": cert_value(item.get("有效资格证书")),
                    "荣誉证书": cert_value(item.get("其他荣誉证书")),
                }
    hr_review = {
        "陈淑怡": "无", "左晋铭": "焊工证", "汤冉": "无", "王培培": "初级会计证",
        "周子豪": "无", "董健": "无", "林威": "无", "王仁斌": "起重工证",
        "徐丹": "无", "杨悦": "无", "徐成威": "无", "沈冲": "焊工证、高压电工证",
        "薛慧": "安全C证", "杨荣康": "无", "赵乐": "助理工程师", "周勋禹": "无",
    }
    cert_rows = []
    for r in supervisor:
        name = clean(r.get("成员姓名"))
        yearend_item = yearend_map.get(name, {})
        title = merge_cert(yearend_item.get("职称证书"), r.get("职称证书"))
        qual = merge_cert(yearend_item.get("资格证书"), r.get("资格证书"))
        honor = merge_cert(yearend_item.get("荣誉证书"), r.get("荣誉证书"), roster_honor.get(name))
        original_has_any = any([title, qual, honor])
        if name in hr_review and hr_review[name] != "无":
            qual = merge_cert(qual, hr_review[name])
        try:
            start = pd.to_datetime(roster_start.get(name) or r.get("成员到本单位时间"))
            years = (pd.Timestamp(DATE) - start).days / 365.25
        except Exception:
            years = 0
        cert_rows.append({
            **r,
            "姓名": name,
            "部门": clean(r.get("成员部门")),
            "工作年限": years,
            "职称证书": title,
            "资格证书": qual,
            "荣誉证书": honor,
            "台账荣誉证书": roster_honor.get(name, ""),
            "原始有证书": original_has_any,
            "有证书": any([title, qual, honor]),
        })

    no_cert_original = [r for r in cert_rows if not r["原始有证书"]]
    no_cert = [r for r in cert_rows if not r["有证书"]]
    no_cert_full2 = [r for r in no_cert if r["工作年限"] >= 2]
    no_cert_less2 = [r for r in no_cert if r["工作年限"] < 2]
    no_cert_d_above = [r for r in no_cert_full2 if grade_idx(clean(r.get("规则重算评级"))) <= grade_idx("D")]
    roster_honor_rows = [r for r in cert_rows if clean(r.get("台账荣誉证书"))]

    def cert_list(rows):
        return "、".join(clean(r.get("姓名")) for r in rows) or "无"

    def honor_cert_list(rows, field="荣誉证书"):
        return "、".join(f"{clean(r.get('姓名'))}（{clean(r.get(field))}）" for r in rows if clean(r.get(field))) or "无"

    current_names = {clean(r.get("姓名")) for r in cert_rows}
    hr_review_in_scope = {name: result for name, result in hr_review.items() if name in current_names}
    review_rows = [[i, name, "原无证书人员复核", result] for i, (name, result) in enumerate(hr_review_in_scope.items(), start=1)]
    reviewed_has_cert = [r for r in cert_rows if clean(r.get("姓名")) in hr_review_in_scope and r["有证书"]]
    reviewed_no_cert = [r for r in cert_rows if clean(r.get("姓名")) in hr_review_in_scope and not r["有证书"]]
    hr_review_section = ""
    if hr_review_in_scope:
        hr_review_section = f"""
#### 人力资源复核后的收紧说明

在上述原始信息采集口径基础上，根据人力资源反馈的《人员证书复核》截图，本轮对原“无任何证书人员”名单作进一步核验。{scope_label()}涉及复核人员 `{len(hr_review_in_scope)}` 人，其中核实仍无证书 `{len(reviewed_no_cert)}` 人，补充确认已有证书 `{len(reviewed_has_cert)}` 人。

复核后仍无任何证书人员包括：`{cert_list(no_cert)}`。

人力资源补充确认已有证书人员包括：`{"、".join(f"{clean(r.get('姓名'))}（{hr_review_in_scope.get(clean(r.get('姓名')))}）" for r in reviewed_has_cert) or "无"}`。

其中，工作满 `2` 年仍无任何证书，且当前基础分级为 `D` 及以上的人员包括：`{cert_list(no_cert_d_above)}`。

有该截图作为佐证时，证书维度建议采用收紧后的复核口径；若无该截图或同等证据材料，则维持上方原始信息采集口径，不自动收紧。

人力资源证书复核明细如下：

{table(["序号", "人员姓名", "情况说明", "证书核实情况"], review_rows)}
"""

    sup_load_low = [r for r in supervisor if norm_load(r.get("13. 以目前的业务量，您认为该员工的工作忙闲程度如何？")) in {"欠饱和", "非饱和"}]
    workload_by_name = {clean(r.get("姓名")): r for r in workload_records}
    low_load_review = []
    low_load_other = []
    busy_self_low_sup = []
    for r in sup_load_low:
        name = clean(r.get("成员姓名"))
        w = workload_by_name.get(name)
        emp_load = clean(w.get("总体工作负荷")) if w else ""
        auto_load = norm_load(w.get("自动工作负荷")) if w else ""
        item = {**r, "员工自评": emp_load, "自动负荷": auto_load}
        if grade_idx(clean(r.get("规则重算评级"))) <= grade_idx("D") and (not w or auto_load in {"欠饱和", "非饱和", ""}):
            low_load_review.append(item)
        else:
            low_load_other.append(item)
        if any(x in emp_load for x in ["满负荷", "超负荷", "繁忙"]) and norm_load(r.get("13. 以目前的业务量，您认为该员工的工作忙闲程度如何？")) in {"欠饱和", "非饱和"}:
            busy_self_low_sup.append(item)

    erp_source = ROOT / "ERP使用情况.xlsx"
    ops = pd.read_excel(erp_source, sheet_name="近三年")
    for col in ["部门", "人员名称", "参与类型"]:
        ops[col] = ops[col].map(clean)
    for col in ["参与流程数量", "处理时间（分钟）", "驳回次数", "有效评论次数"]:
        ops[col] = pd.to_numeric(ops[col], errors="coerce").fillna(0)
    role = ops.pivot_table(index="人员名称", columns="参与类型", values="参与流程数量", aggfunc="sum", fill_value=0).reset_index()
    for c in ["发起者", "参与审批者"]:
        if c not in role:
            role[c] = 0
    agg = ops.groupby("人员名称", as_index=False).agg(处理时间分钟=("处理时间（分钟）", "sum"), 驳回次数=("驳回次数", "sum"), 有效评论次数=("有效评论次数", "sum"))
    employees = load_employee_base()
    person = employees.merge(role, left_on="姓名", right_on="人员名称", how="left").merge(agg, left_on="姓名", right_on="人员名称", how="left")
    for c in ["发起者", "参与审批者", "处理时间分钟", "驳回次数", "有效评论次数"]:
        person[c] = pd.to_numeric(person[c], errors="coerce").fillna(0)
    grade_map = {clean(r.get("成员姓名")): clean(r.get("规则重算评级")) for r in supervisor}
    sup_map = {clean(r.get("成员姓名")): clean(r.get("主管主观评级")) for r in supervisor}
    person["规则重算评级"] = person["姓名"].map(lambda n: grade_map.get(clean(n), "未评级"))
    person["主管主观评级"] = person["姓名"].map(lambda n: sup_map.get(clean(n), "未评级"))
    person["流程合计"] = person["发起者"] + person["参与审批者"]
    person["审批平均处理小时数"] = person.apply(lambda r: r["处理时间分钟"] / 60 / r["参与审批者"] if r["参与审批者"] else 0, axis=1)
    slow_approval = person[(person["参与审批者"].gt(0)) & (person["审批平均处理小时数"].ge(120))].sort_values("审批平均处理小时数", ascending=False)
    slow_key = slow_approval[slow_approval["规则重算评级"].isin(["S", "A", "B"])]
    low_erp_d_above = person[(person["流程合计"].le(52)) & (person["规则重算评级"].isin(["S", "A", "B", "C", "D"]))].sort_values(["部门", "流程合计"])
    low_erp_other = person[(person["流程合计"].le(52)) & (~person["规则重算评级"].isin(["S", "A", "B", "C", "D"]))].sort_values(["部门", "流程合计"])
    top_flow = person.sort_values("流程合计", ascending=False).head(1)
    top_comment = person.sort_values("有效评论次数", ascending=False).head(1)
    top_reject = person.sort_values("驳回次数", ascending=False).head(1)

    abnormal_fields = [
        ("岗位必要性偏弱", "该成员所在岗位是否有存在的必要性", {"B", "C", "D"}),
        ("离职影响偏弱", "20. 如果该员工明天提出辞职，你的第一反应是？", {"C", "D", "E"}),
        ("沟通偏弱", "8. 你认为该员工在和其他部门沟通事情时，效果如何？", {"C", "D"}),
        ("上手偏弱", "16. 面对新系统、新业务及新技能知识时，该员工的上手速度是？", {"C", "D"}),
        ("效率偏弱", "1.在日常工作中，该成员是否能做到高效、按时完成？", {"C", "D"}),
        ("优化偏弱", "7.他会主动思考如何把现有的工作做得更好吗？", {"C", "D"}),
        ("同岗偏弱", "18.与同岗位的平均水平相比，该员工的胜任力处于什么水平？", {"C", "D"}),
    ]
    key_abnormal = []
    for r in supervisor:
        items = []
        for label, field, letters in abnormal_fields:
            value = clean(r.get(field))
            if option_is(value, letters):
                items.append(f"{label}：{value}")
        if items:
            key_abnormal.append((r, items))
    role_value_abnormal = [(r, items) for r, items in key_abnormal if any(x.startswith("岗位必要性偏弱") or x.startswith("离职影响偏弱") for x in items)]
    role_value_c_above = [r for r, items in role_value_abnormal if grade_idx(clean(r.get("规则重算评级"))) <= grade_idx("C")]

    under1 = [r for r in cert_rows if r["工作年限"] < 1]
    under1_low = [r for r in under1 if clean(r.get("规则重算评级")) in {"E", "F"}]

    def sup_list(rows):
        return "、".join(clean(r.get("成员姓名")) for r in rows) or "无"

    def erp_list(df, include_hours=False):
        if df.empty:
            return "无"
        values = []
        for _, r in df.iterrows():
            if include_hours:
                values.append(f"{r['姓名']}（{fmt(r['审批平均处理小时数'], 1)}小时，{r['规则重算评级']}）")
            else:
                values.append(f"{r['姓名']}（{fmt(r['流程合计'], 0)}次，{r['规则重算评级']}）")
        return "、".join(values)

    short_depts = []
    for dept in SCOPE_DEPTS:
        count = STAFF_COUNTS.get(dept, 0)
        miss_w = len(missing_workload[missing_workload["部门"] == dept])
        miss_s = len(missing_supervisor[missing_supervisor["部门"] == dept])
        if miss_w or miss_s:
            short_depts.append((dept, miss_w + miss_s))
    short_dept_text = "、".join(d for d, _ in sorted(short_depts, key=lambda x: x[1], reverse=True)[:5]) or "暂无明显短板部门"

    return f"""# 机关工作量摸排综合报告-{DATE}

## 一、总体业务流程

本次机关工作量摸排和人员评价，整体按“表单采集、规则赋分、辅助校正、形成建议”的流程开展。

具体做法是：先在钉钉搭建员工工作摸排表单和主管评价表单，围绕员工日常交付、工作质量、协同能力、成长潜力、岗位价值和工作负荷等内容设计题目，再形成主管评价赋分规则，对已完成主管评价的员工先做基础分级；在此基础上，再结合主管评价中的工作负荷、员工工作摸排中的工作负荷、ERP 使用情况、员工证书情况以及重点问题反馈，对基础分级结果做局部复核和调整建议。

## 二、填写情况

当前报告范围内，共涉及 `{len(SCOPE_DEPTS)}` 个部门、`{TOTAL_STAFF}` 名员工，部门为：`{scope_detail()}`。

其中：

- 已完成工作摸排 `{len(workload_records)}` 人，填报率 `{pct(len(workload_records), TOTAL_STAFF)}`
- 已完成主管评价 `{len(supervisor)}` 人，完成率 `{pct(len(supervisor), TOTAL_STAFF)}`

### 1. 未填报工作摸排人员名单

{table(["部门", "未填报人数", "人员名单"], missing_rows(missing_workload))}

### 2. 未完成主管评价人员名单

{table(["部门", "未完成人数", "人员名单"], missing_rows(missing_supervisor))}

总体上，当前样本已经具备开展阶段性分级和局部人员调整建议的基础。当前影响整体判断完整性的短板，仍然主要集中在 `{short_dept_text}`。

## 三、主管评价基础分级结果

当前主管评价基础分级，按《机关员工主管评价赋分规则》执行。本轮分值测算不再直接采用表内 `题1分值` 至 `题20分值`，而是读取主管评价表中每题的原始选项文本，按规则文件中的选项-分值表逐题重新赋分，并用规则档位区间复核基础分级。分级标准如下：

{table(["分数区间", "等级"], [["95分及以上", "S"], ["85-94.5分", "A"], ["75-84.5分", "B"], ["60-74.5分", "C"], ["50-59.5分", "D"], ["40-49.5分", "E"], ["40分以下", "F"]])}

规则重算校验结果：按当前规则文件读取题1至题20原始选项文本、逐题匹配规则分值并套用档位区间后，发现 `{len(auto_rule_diff)}` 人与表内自动生成评级不一致：{auto_rule_preview or "无"}。建议优先复核上述人员的题项分值或等级公式。

当前 `{len(supervisor)}` 名已完成主管评价员工的基础分级分布如下：

{table(["档位", "人数", "部门及人员"], dist_rows("规则重算评级"))}

当前 `{len(supervisor)}` 名已完成主管评价员工的主管原始评价等级分布如下：

{table(["档位", "人数", "部门及人员"], dist_rows("主管主观评级"))}

当前主管原始评价等级与规则重算评级不一致的人员，单独按部门列示如下。

### 1. 基础分级明细

![各部门等级热力矩阵对比]({matrix_path})

{table(["部门", "姓名", "表内20题总分", "纯20题重算总分", "岗位必要性修正分", "综合重算总分", "表内基础分级", "综合重算评级", "主管评价等级", "规则匹配情况"], detail_rows)}

### 2. 规则重算评级与主管评价差异人员

{table(["部门", "差异人员"], diff_by_dept)}

## 四、局部复核和调整建议

局部复核和调整，当前优先看五个辅助维度：

1. `证书情况`：重点识别“工作满2年仍无证书”与“工作未满1年暂不宜按低档固化”两类情形。
2. `工作负荷度`：同时看员工工作摸排负荷、自动工作负荷和主管主观工作负荷，判断是否存在低负荷、错配或明显承压。
3. `ERP 使用情况`：看是否存在稳定发起、审批或持续参与痕迹，用于辅助判断岗位承担度和实际业务参与度。
4. `岗位必要性和关键题目异常`：看岗位是否需要保留，以及效率、同岗胜任力、上手速度、离职影响、沟通和优化意识等题目是否集中偏弱。
5. `重点问题反馈`：结合员工开放题、主管补充反馈和家庭照护/岗位优化意向，识别流程简化、岗位边界和特殊人群管理需求。

### 1. 证书维度信号

当前无任何证书人员包括：`{cert_list(no_cert_original)}`。

人员台账补充采集到优秀员工或安全先进个人荣誉的人员包括：`{honor_cert_list(roster_honor_rows, "台账荣誉证书")}`。

其中，工作满 `2` 年仍无任何证书，且当前基础分级为 `D` 及以上的人员包括：`{cert_list(no_cert_d_above)}`。

工作满 `2` 年仍无任何证书的其他人员包括：`{cert_list([r for r in no_cert_full2 if r not in no_cert_d_above])}`。

工作不满 `2` 年且当前无任何证书的人员包括：`{cert_list(no_cert_less2)}`。

按当前复核口径，工作满 `2` 年仍无任何证书的人员，原则上最高不宜超过 `D` 级。结合当前基础分级，建议纳入下调复核的人员包括：`{cert_list([r for r in no_cert_d_above if clean(r.get("规则重算评级")) in {"S", "A", "B", "C"}])}`。

{hr_review_section}

### 2. 工作负荷维度信号

当前主管主观工作负荷偏低的人员包括：`{sup_list(sup_load_low)}`。

其中，员工自评或自动工作负荷也未显示明显繁忙，且当前基础分级为 `D` 及以上的人员包括：`{sup_list(low_load_review)}`。

员工自评或自动工作负荷未显示明显繁忙的其他人员包括：`{sup_list(low_load_other)}`。

主管评价不忙、但员工自评仍显示繁忙的人员包括：`{sup_list(busy_self_low_sup)}`。

按当前复核口径，主管主观工作负荷已明确偏低的人员，原则上最高不宜超过 `D` 级。结合当前基础分级，建议纳入下调复核的人员包括：`{sup_list([r for r in low_load_review if clean(r.get("规则重算评级")) in {"S", "A", "B", "C"}])}`。

### 3. ERP 使用情况信号

本轮ERP维度直接引用 `ERP使用情况.xlsx` 中 `近三年` 工作簿，并与 `职能部门人员台账` 口径匹配。ERP信号主要用于观察流程承担度、审批端集中度、评论和驳回留痕情况。

{table(["ERP信号", "人员", "复核含义"], [["流程量靠前", erp_list(top_flow), "可作为工作承担度、流程节点占用和岗位饱和度的正向佐证；对基础分级偏低但ERP承担显著的人员，建议避免简单下调。"], ["驳回次数靠前", "、".join(f"{r['姓名']}（驳回{fmt(r['驳回次数'], 0)}次）" for _, r in top_reject.iterrows()), "反映审批把关、资料校验或流程纠偏责任较重；需结合驳回质量判断，不宜直接理解为负面。"], ["有效评论靠前", "、".join(f"{r['姓名']}（评论{fmt(r['有效评论次数'], 0)}次）" for _, r in top_comment.iterrows()), "反映审批意见留痕、跨岗位解释和流程推动工作量；可作为隐性协调负荷补充信号。"], ["ERP痕迹较少且基础分级D及以上", erp_list(low_erp_d_above.head(30)), "对应岗位是否确需ERP参与需要部门解释；若岗位职责应使用ERP但记录很少，可作为岗位职责复核信号。"]])}

平均审核周期超过 `120` 小时，且当前基础分级为 `B` 及以上的人员包括：`{erp_list(slow_key, include_hours=True)}`。

其他平均审核周期超过 `120` 小时的人员包括：`{erp_list(slow_approval[~slow_approval["姓名"].isin(set(slow_key["姓名"]))], include_hours=True)}`。

没有参与过任何 ERP 痕迹，或 ERP 发起加审批总数不超过 `52` 次，且当前基础分级为 `D` 及以上的人员包括：`{erp_list(low_erp_d_above)}`。

其他 ERP 痕迹较少的人员包括：`{erp_list(low_erp_other)}`。

### 4. 岗位价值和关键题目异常信号

当前岗位必要性、离职影响和关键题目异常最突出的人员包括：`{sup_list([r for r, _ in sorted(key_abnormal, key=lambda x: -len(x[1]))[:20]])}`。

其中，离职影响和岗位关键价值信号不高、但当前基础分级仍为 `C` 类及以上的人员包括：`{sup_list(role_value_c_above)}`。按当前复核口径，上述人员原则上不宜超过 `D` 级，建议纳入下调复核或逐人说明。

### 5. 未满1年人员保护原则

对于工作未满 `1` 年的员工，本轮不直接纳入 `E/F` 调整建议，统一作为培养观察样本处理。当前涉及人员主要为：

{chr(10).join(f'- `{clean(r.get("姓名"))}（{fmt(r.get("工作年限"), 1)}年，基础分级{clean(r.get("规则重算评级"))}）`' for r in under1) or "- `无`"}

其中：

- `{cert_list(under1_low)}` 当前不建议直接按低档固化，应先列入试用/培养期观察名单。
- 对未满一年但已经达到 `B/C` 档的人员，应结合岗位实际承担、ERP痕迹和主管原文评价继续观察。

## 五、综合结论

1. 主管评价20题重算后，`S/A/B` 合计 `{sum(1 for r in supervisor if clean(r.get("规则重算评级")) in {"S", "A", "B"})}` 人，`C/D/E/F` 合计 `{sum(1 for r in supervisor if clean(r.get("规则重算评级")) in {"C", "D", "E", "F"})}` 人，整体评价结构反映{scope_label()}样本的分布情况。
2. 工作摸排已完成 `{len(workload_records)}` 人，较参考稿覆盖面提升，但仍存在未填报人员，应继续补齐后再作最终定档。
3. 证书、工作负荷、ERP 和岗位价值四类辅助维度应作为“复核信号”，不宜单独替代主管评价，但可以用于识别评级偏高、岗位低饱和或客观支撑不足人员。
4. 对工作满2年仍无证书、主管负荷偏低且评级偏高、ERP痕迹明显不足、岗位必要性和离职影响偏弱的交叉人员，应优先进入逐人复核。
5. 对未满1年人员、家庭照护需求人员和员工反馈中明确表达岗位优化诉求的人员，应以培养观察、岗位边界澄清和流程减负为主，不宜简单按低档固化。

## 六、附件建议

本报告正文只保留结论性结果，详细规则、重算结果和专项分析可配套查阅：

{chr(10).join(f'{i}. [{p.name}]({p})' for i, p in enumerate(paths, start=1))}
"""


def inline_markup(text: str) -> str:
    text = html.escape(str(text or "").replace("<br>", "\n"))
    text = re.sub(r"`([^`]+)`", r"\1", text)
    text = re.sub(r"\*\*([^*]+)\*\*", r"<b>\1</b>", text)
    return text.replace("\n", "<br/>")


def split_table_row(line: str) -> list[str]:
    return [cell.strip() for cell in line.strip().strip("|").split("|")]


def is_separator(line: str) -> bool:
    return bool(re.match(r"^\|\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?$", line.strip()))


def chunk_text(text: str, size: int) -> list[str]:
    text = str(text or "")
    if len(text) <= size:
        return [text]
    chunks = []
    while text:
        cut = size
        for sep in ["；", "。", "，", "、", ";", ","]:
            pos = text.rfind(sep, 0, size)
            if pos > size * 0.55:
                cut = pos + 1
                break
        chunks.append(text[:cut])
        text = text[cut:]
    return chunks


def expand_long_rows(rows: list[list[str]], max_chars: int) -> list[list[str]]:
    if not rows:
        return rows
    expanded = [rows[0]]
    width = len(rows[0])
    for row in rows[1:]:
        row = row + [""] * (width - len(row))
        chunks_by_cell = [chunk_text(cell, max_chars) for cell in row[:width]]
        parts = max(len(chunks) for chunks in chunks_by_cell)
        for i in range(parts):
            new_row = []
            for chunks in chunks_by_cell:
                new_row.append(chunks[i] if i < len(chunks) else "")
            expanded.append(new_row)
    return expanded


def col_widths(headers: list[str], page_width: float) -> list[float]:
    col_count = len(headers)
    if col_count <= 2:
        if headers in (["题目", "分值"], ["选项", "分值"]):
            return [page_width - 22 * mm, 22 * mm]
        if headers in (["文件类型", "主要用途"], ["规则", "说明"], ["字段", "用途"]):
            return [50 * mm, page_width - 50 * mm]
        if headers == ["部门", "差异人员"]:
            return [34 * mm, page_width - 34 * mm]
        compact_headers = {
            "人数",
            "数值",
            "有效反馈人数",
            "记录数",
            "发起流程总数",
            "涉及人数",
            "角色特征",
            "负荷层级",
            "差异档位",
        }
        if col_count == 2 and any(h in compact_headers for h in headers):
            return [page_width / 2] * 2
        return [36 * mm, page_width - 36 * mm]
    if col_count == 3:
        if headers == ["等级", "工作状态倾向", "人数"]:
            return [24 * mm, page_width - 48 * mm, 24 * mm]
        if headers in (["档位", "人数", "部门及人员"], ["等级", "人数", "部门及人员"]):
            return [18 * mm, 18 * mm, page_width - 36 * mm]
        if headers == ["ERP信号", "人员", "复核含义"]:
            return [30 * mm, page_width * 0.42, page_width - 30 * mm - page_width * 0.42]
        return [page_width / 3] * 3
    if col_count == 4:
        return [28 * mm, 24 * mm, page_width * 0.24, page_width - 52 * mm - page_width * 0.24]
    if col_count == 5:
        return [28 * mm, 18 * mm, 32 * mm, 44 * mm, page_width - 122 * mm]
    if col_count == 7:
        if headers == ["姓名", "部门", "自动生成评级", "主管主观评级", "最终评价级别", "异常项数", "异常选项"]:
            return [18 * mm, 24 * mm, 18 * mm, 20 * mm, 18 * mm, 16 * mm, page_width - 98 * mm]
        if headers == ["部门", "姓名", "20题重算评级", "主管主观评级", "差异判断", "主管评级薪资区间最小值", "主管评级薪资区间最大值"]:
            return [28 * mm, 16 * mm, 17 * mm, 17 * mm, page_width - 108 * mm, 27 * mm, 27 * mm]
    if col_count == 10:
        if headers == ["部门", "姓名", "工作年限", "按20题计算等级", "按20题计算等级对应薪资中位数", "主管评价等级", "主管评价等级对应薪资中位数", "职称证书", "资格证书", "荣誉证书"]:
            return [16 * mm, 11 * mm, 8 * mm, 10 * mm, 12 * mm, 10 * mm, 12 * mm, 16 * mm, 16 * mm, page_width - 111 * mm]
    return [page_width / col_count] * col_count


def markdown_to_pdf(md_path: Path, pdf_path: Path) -> None:
    pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    styles = getSampleStyleSheet()
    body = ParagraphStyle("body", parent=styles["Normal"], fontName="STSong-Light", fontSize=10.5, leading=16, spaceAfter=7)
    h1 = ParagraphStyle("h1", parent=body, fontSize=20, leading=28, alignment=1, spaceAfter=14)
    h2 = ParagraphStyle("h2", parent=body, fontSize=16, leading=22, spaceBefore=12, spaceAfter=8)
    h3 = ParagraphStyle("h3", parent=body, fontSize=12.5, leading=18, spaceBefore=10, spaceAfter=6)
    h4 = ParagraphStyle("h4", parent=body, fontSize=11.2, leading=16, spaceBefore=8, spaceAfter=5)
    bullet = ParagraphStyle("bullet", parent=body, leftIndent=12)

    pagesize = A4
    doc = SimpleDocTemplate(
        str(pdf_path),
        pagesize=pagesize,
        leftMargin=22 * mm,
        rightMargin=22 * mm,
        topMargin=16 * mm,
        bottomMargin=16 * mm,
    )
    page_width = pagesize[0] - 44 * mm
    story = []
    lines = md_path.read_text(encoding="utf-8").splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        if not line:
            story.append(Spacer(1, 3))
            i += 1
            continue
        if line.startswith("|"):
            table_lines = []
            while i < len(lines) and lines[i].startswith("|"):
                if not is_separator(lines[i]):
                    table_lines.append(lines[i])
                i += 1
            rows = [split_table_row(x) for x in table_lines]
            if rows:
                max_cols = max(len(r) for r in rows)
                rows = [(r + [""] * (max_cols - len(r)))[:max_cols] for r in rows]
                font_size = 8.2 if max_cols <= 5 else 6.6 if max_cols <= 8 else 5.4
                cell = ParagraphStyle("cell", parent=body, fontSize=font_size, leading=font_size + 3.4, spaceAfter=0)
                head = ParagraphStyle("head", parent=cell, fontSize=font_size + 0.3, leading=font_size + 3.8, spaceAfter=0)
                data = [[Paragraph(inline_markup(c), head if ridx == 0 else cell) for c in row] for ridx, row in enumerate(rows)]
                tbl = Table(data, colWidths=col_widths(rows[0], page_width), repeatRows=1, splitByRow=1)
                tbl.setStyle(
                    TableStyle(
                        [
                            ("FONTNAME", (0, 0), (-1, -1), "STSong-Light"),
                            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9EAF7")),
                            ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#BFBFBF")),
                            ("VALIGN", (0, 0), (-1, -1), "TOP"),
                            ("LEFTPADDING", (0, 0), (-1, -1), 3),
                            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                            ("TOPPADDING", (0, 0), (-1, -1), 5),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                        ]
                    )
                )
                story.append(tbl)
                story.append(Spacer(1, 6))
            continue
        if line.startswith("# "):
            if story:
                story.append(PageBreak())
            story.append(Paragraph(inline_markup(line[2:]), h1))
        elif line.startswith("## "):
            story.append(Paragraph(inline_markup(line[3:]), h2))
        elif line.startswith("#### "):
            story.append(Paragraph(inline_markup(line[5:]), h4))
        elif line.startswith("### "):
            story.append(Paragraph(inline_markup(line[4:]), h3))
        elif line.startswith("- "):
            story.append(Paragraph("• " + inline_markup(line[2:]), bullet))
        elif re.match(r"^\d+\.\s+", line):
            story.append(Paragraph(inline_markup(line), bullet))
        elif line.startswith("> "):
            story.append(Paragraph(inline_markup(line[2:]), bullet))
        elif line.startswith("!["):
            match = re.match(r"!\[[^\]]*\]\((.+?)\)", line)
            image_path = Path(match.group(1)) if match else None
            if image_path and image_path.exists():
                img = RLImage(str(image_path))
                ratio = img.imageHeight / img.imageWidth if img.imageWidth else 1
                img.drawWidth = page_width
                img.drawHeight = page_width * ratio
                max_height = pagesize[1] - 58 * mm
                if img.drawHeight > max_height:
                    img.drawHeight = max_height
                    img.drawWidth = max_height / ratio
                story.append(img)
                story.append(Spacer(1, 6))
            else:
                story.append(Paragraph("图片引用：" + inline_markup(line), body))
        else:
            story.append(Paragraph(inline_markup(line), body))
        i += 1
    doc.build(story)


def write_report(stem: str, md: str) -> Path:
    path = ROOT / f"{OUTPUT_PREFIX}{stem}-{DATE}.md"
    path.write_text(md, encoding="utf-8")
    markdown_to_pdf(path, path.with_suffix(".pdf"))
    return path


def main():
    global SCOPE_DEPTS, STAFF_COUNTS, TOTAL_STAFF, OUTPUT_PREFIX
    SCOPE_DEPTS, STAFF_COUNTS = infer_department_scope()
    TOTAL_STAFF = sum(STAFF_COUNTS.values())
    OUTPUT_PREFIX = scope_output_prefix()
    supervisor = load_supervisor()
    supervisor_evaluations = load_supervisor_evaluations()
    workload = load_workload()

    outputs = []
    items = [
        ("机关员工主管评价20题重算报告", build_supervisor_report(supervisor_evaluations)),
        ("机关员工ERP使用情况分析", build_erp_report(supervisor)),
    ]
    workload_md, workload_records = build_workload_report(workload, supervisor)
    items.append(("机关员工工作负荷报告", workload_md))
    items.append(("机关员工证书客观维度分析", build_cert_report(supervisor)))
    items.append(("机关工作量摸排报告", build_workload_overview(workload, supervisor)))
    items.append(("机关工作量摸排结论分析", build_conclusions(workload_records, supervisor)))
    items.append(("有想法人员及妈妈岗意向名单", build_idea_mom_report(workload, supervisor)))
    feedback_md, feedback_conclusion = build_feedback(workload, supervisor)
    items.append(("重点问题反馈汇总", feedback_md))
    items.append(("重点问题反馈结论分析", feedback_conclusion))

    for stem, md in items:
        outputs.append(write_report(stem, md))

    integrated = write_report("机关工作量摸排综合报告", build_integrated(outputs, workload_records, supervisor))
    outputs.append(integrated)

    print("生成文件：")
    for p in outputs:
        print(p.name)
    print(f"识别部门：{', '.join(SCOPE_DEPTS)}")
    print(f"统计员工：{TOTAL_STAFF}")
    print(f"工作摸排样本：{len(workload)}")
    print(f"主管评价人员：{len(supervisor)}")
    print(f"主管评价记录：{len(supervisor_evaluations)}")


if __name__ == "__main__":
    main()
