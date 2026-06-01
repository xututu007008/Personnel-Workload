from __future__ import annotations

import html
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
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


ROOT = Path(__file__).resolve().parent
DATE = datetime.now().strftime("%Y-%m-%d")
REPORT_STEM = "机关员工主管评价20题重算报告"
SUPERVISOR_PREFIX = "部门成员工作主管评定"
RULE_PREFIX = "机关员工主管评价赋分规则"
SCOPE_DEPTS = [
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
GRADE_ORDER = ["S", "A", "B", "C", "D", "E", "F"]
SALARY = {"S": 10000, "A": 9000, "B": 8000, "C": 6500, "D": 5000, "E": 4000, "F": 3000}
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


def latest(prefix: str, suffix: str) -> Path:
    files = [p for p in ROOT.iterdir() if p.name.startswith(prefix) and p.suffix.lower() == suffix]
    if not files:
        raise FileNotFoundError(f"未找到文件：{prefix}*{suffix}")
    return max(files, key=lambda p: (p.stat().st_mtime, p.name))


def clean(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    return "" if text.lower() == "nan" else text


def num(value) -> float:
    try:
        if pd.isna(value):
            return 0.0
        return float(value)
    except Exception:
        return 0.0


def fmt(value, digits: int = 2) -> str:
    try:
        number = float(value)
    except Exception:
        return "-"
    if abs(number - round(number)) < 1e-9:
        return str(int(round(number)))
    return f"{number:.{digits}f}".rstrip("0").rstrip(".")


def table(headers, rows) -> str:
    result = ["| " + " | ".join(headers) + " |", "| " + " | ".join(["---"] * len(headers)) + " |"]
    for row in rows:
        result.append("| " + " | ".join(str(item).replace("\n", "<br>") for item in row) + " |")
    return "\n".join(result)


def normalize_option_text(value) -> str:
    text = str(value or "").strip().replace("`", "")
    text = re.sub(r"\s+", "", text)
    return text.replace("：", ":").replace("。", "")


def option_letter(value) -> str:
    match = re.match(r"^([A-Z])\s*[.．、]", str(value or "").strip())
    return match.group(1) if match else ""


def grade_from_score(score) -> str:
    value = float(score)
    for grade, low, high in GRADE_SCORE_RULES:
        if low <= value < high:
            return grade
    return "-"


def load_rule_scores() -> dict[int, dict]:
    rule_path = latest(RULE_PREFIX, ".md")
    markdown = rule_path.read_text(encoding="utf-8")
    rules: dict[int, dict] = {}
    for qno in range(1, 21):
        pattern = rf"### 第{qno}题.*?(?=\n### 第{qno + 1}题|\n## |\Z)"
        match = re.search(pattern, markdown, flags=re.S)
        if not match:
            continue
        exact: dict[str, float] = {}
        by_letter: dict[str, float] = {}
        for line in match.group(0).splitlines():
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


def question_header(record: dict, qno: int) -> str | None:
    for key in record:
        text = str(key).strip()
        if text.startswith(f"{qno}.") or text.startswith(f"{qno}．"):
            return key
    return None


def enrich_score(record: dict, rules: dict[int, dict]) -> None:
    total = 0.0
    exact_count = 0
    letter_count = 0
    missing = []
    for qno in range(1, 21):
        header = question_header(record, qno)
        answer = record.get(header) if header else None
        rule = rules.get(qno, {})
        score = rule.get("exact", {}).get(normalize_option_text(answer))
        if score is not None:
            exact_count += 1
        else:
            letter = option_letter(answer)
            score = rule.get("by_letter", {}).get(letter)
            if score is not None:
                letter_count += 1
        if score is None:
            score = 0.0
            missing.append(f"题{qno}")
        total += score
    pure_total = round(total, 1)
    necessity_adjustment = NECESSITY_SCORE.get(option_letter(record.get("该成员所在岗位是否有存在的必要性")), 0)
    total = round(pure_total + necessity_adjustment, 1)
    record["纯20题总分"] = int(pure_total) if float(pure_total).is_integer() else pure_total
    record["岗位必要性修正分"] = necessity_adjustment
    record["规则重算总分"] = int(total) if float(total).is_integer() else total
    record["规则重算评级"] = grade_from_score(total)
    record["规则匹配情况"] = f"精确{exact_count}题/字母{letter_count}题/未匹配{len(missing)}题"
    record["规则未匹配题目"] = "、".join(missing) if missing else "无"


def load_records() -> tuple[list[dict], Path, Path]:
    excel_path = latest(SUPERVISOR_PREFIX, ".xlsx")
    rule_path = latest(RULE_PREFIX, ".md")
    frame = pd.read_excel(excel_path, sheet_name="Sheet1", header=1)
    frame = frame.loc[:, [col for col in frame.columns if not str(col).startswith("Unnamed:")]]
    frame = frame[frame["成员姓名"].notna() & frame["成员部门"].isin(SCOPE_DEPTS)].copy()
    frame["评价时间(必填)"] = frame["评价时间(必填)"].astype(str)
    frame = frame.sort_values("评价时间(必填)").drop_duplicates(["成员部门", "成员姓名"], keep="last")
    rules = load_rule_scores()
    records = frame.to_dict("records")
    for record in records:
        enrich_score(record, rules)
    return records, excel_path, rule_path


def names_by_dept(records: list[dict]) -> str:
    grouped = defaultdict(list)
    for record in records:
        grouped[clean(record.get("成员部门"))].append(clean(record.get("成员姓名")))
    parts = []
    for dept in SCOPE_DEPTS:
        names = [name for name in grouped.get(dept, []) if name]
        if names:
            parts.append(f"{dept}：{'、'.join(names)}")
    return "；".join(parts) or "-"


def distribution(records: list[dict], field: str) -> str:
    rows = []
    for grade in GRADE_ORDER:
        subset = [record for record in records if clean(record.get(field)) == grade]
        if subset:
            rows.append([f"`{grade}`", f"`{len(subset)}`", f"`{names_by_dept(subset)}`"])
    return table(["档位", "人数", "部门及人员"], rows)


def build_report(records: list[dict], excel_path: Path, rule_path: Path) -> str:
    scores = [num(record.get("规则重算总分")) for record in records]
    rule_counts = Counter(clean(record.get("规则重算评级")) for record in records)
    supervisor_counts = Counter(clean(record.get("主管主观评级")) for record in records)
    auto_counts = Counter(clean(record.get("自动生成评级")) for record in records)
    grade_rows = [
        ["综合重算等级（含岗位必要性修正）", *[rule_counts.get(grade, 0) for grade in GRADE_ORDER]],
        ["主管评价等级", *[supervisor_counts.get(grade, 0) for grade in GRADE_ORDER]],
        ["表内自动等级", *[auto_counts.get(grade, 0) for grade in GRADE_ORDER]],
    ]
    detail = []
    for record in sorted(records, key=lambda item: (SCOPE_DEPTS.index(clean(item.get("成员部门"))), -num(item.get("规则重算总分")))):
        grade = clean(record.get("规则重算评级"))
        supervisor_grade = clean(record.get("主管主观评级"))
        detail.append([
            clean(record.get("成员部门")),
            clean(record.get("成员姓名")),
            fmt(record.get("纯20题总分")),
            fmt(record.get("岗位必要性修正分")),
            fmt(record.get("规则重算总分")),
            grade,
            SALARY.get(grade, ""),
            supervisor_grade,
            SALARY.get(supervisor_grade, ""),
        ])
    return f"""# {REPORT_STEM}-{DATE}

## 一、测算口径

本报告依据 `{rule_path.name}`，结合 `{excel_path.name}` 中主管对员工 `20` 道问卷题及 `岗位必要性` 独立观察项的评价结果测算形成。本轮先读取20题原始选项文本逐题重算，再按岗位必要性观察项执行扣分修正，不覆盖表内原公式。

## 二、整体结果

- 纳入测算样本：`{len(records)}` 人

### 1. 三类等级分布

{table(["口径", *GRADE_ORDER], grade_rows)}

## 三、样本校核结果

### 1. 当前总分统计

{table(["指标", "数值"], [["最低分", fmt(min(scores))], ["25分位", fmt(pd.Series(scores).quantile(.25))], ["中位数", fmt(pd.Series(scores).median())], ["75分位", fmt(pd.Series(scores).quantile(.75))], ["最高分", fmt(max(scores))]])}

### 2. 当前样本分布

{distribution(records, "规则重算评级")}

## 四、人员明细

{table(["部门", "姓名", "纯20题总分", "岗位必要性修正分", "综合重算总分", "综合重算等级", "综合重算等级对应薪资中位数", "主管评价等级", "主管评价等级对应薪资中位数"], detail)}
"""


def inline_markup(text: str) -> str:
    text = html.escape(str(text or "").replace("<br>", "\n"))
    text = re.sub(r"`([^`]+)`", r"\1", text)
    return text.replace("\n", "<br/>")


def split_table_row(line: str) -> list[str]:
    return [cell.strip() for cell in line.strip().strip("|").split("|")]


def is_separator(line: str) -> bool:
    return bool(re.match(r"^\|\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?$", line.strip()))


def col_widths(headers: list[str], page_width: float) -> list[float]:
    if headers == ["档位", "人数", "部门及人员"]:
        return [18 * mm, 18 * mm, page_width - 36 * mm]
    if len(headers) == 2:
        return [page_width / 2] * 2
    return [page_width / len(headers)] * len(headers)


def markdown_to_pdf(md_path: Path, pdf_path: Path) -> None:
    pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    styles = getSampleStyleSheet()
    body = ParagraphStyle("body", parent=styles["Normal"], fontName="STSong-Light", fontSize=10.5, leading=16, spaceAfter=7)
    h1 = ParagraphStyle("h1", parent=body, fontSize=20, leading=28, alignment=1, spaceAfter=14)
    h2 = ParagraphStyle("h2", parent=body, fontSize=16, leading=22, spaceBefore=12, spaceAfter=8)
    h3 = ParagraphStyle("h3", parent=body, fontSize=12.5, leading=18, spaceBefore=10, spaceAfter=6)
    doc = SimpleDocTemplate(str(pdf_path), pagesize=A4, leftMargin=18 * mm, rightMargin=18 * mm, topMargin=16 * mm, bottomMargin=16 * mm)
    page_width = A4[0] - 36 * mm
    story = []
    lines = md_path.read_text(encoding="utf-8").splitlines()
    index = 0
    while index < len(lines):
        line = lines[index].rstrip()
        if not line:
            story.append(Spacer(1, 3))
            index += 1
            continue
        if line.startswith("|"):
            table_lines = []
            while index < len(lines) and lines[index].startswith("|"):
                if not is_separator(lines[index]):
                    table_lines.append(lines[index])
                index += 1
            rows = [split_table_row(item) for item in table_lines]
            font_size = 8.2 if len(rows[0]) <= 5 else 6.2
            cell = ParagraphStyle("cell", parent=body, fontSize=font_size, leading=font_size + 3.4, spaceAfter=0)
            data = [[Paragraph(inline_markup(value), cell) for value in row] for row in rows]
            report_table = Table(data, colWidths=col_widths(rows[0], page_width), repeatRows=1, splitByRow=1)
            report_table.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9EAF7")),
                ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#BFBFBF")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]))
            story.extend([report_table, Spacer(1, 6)])
            continue
        if line.startswith("# "):
            story.append(Paragraph(inline_markup(line[2:]), h1))
        elif line.startswith("## "):
            story.append(Paragraph(inline_markup(line[3:]), h2))
        elif line.startswith("### "):
            story.append(Paragraph(inline_markup(line[4:]), h3))
        elif line.startswith("- "):
            story.append(Paragraph("• " + inline_markup(line[2:]), body))
        else:
            story.append(Paragraph(inline_markup(line), body))
        index += 1
    doc.build(story)


def main() -> None:
    records, excel_path, rule_path = load_records()
    md_path = ROOT / f"{REPORT_STEM}-{DATE}.md"
    pdf_path = md_path.with_suffix(".pdf")
    md_path.write_text(build_report(records, excel_path, rule_path), encoding="utf-8")
    markdown_to_pdf(md_path, pdf_path)
    print(f"主管评价文件：{excel_path.name}")
    print(f"赋分规则文件：{rule_path.name}")
    print(f"纳入测算样本：{len(records)}")
    print(f"生成文件：{md_path.name}")
    print(f"生成文件：{pdf_path.name}")


if __name__ == "__main__":
    main()
