#!/usr/bin/env python3
"""Add BD merchant detail sections to the current dashboard.

This script reads the latest local merchant daily workbook, keeps the current
subsidy/dashboard modules intact, and appends BD-level merchant detail data.
"""

from __future__ import annotations

import html
import json
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
REPORT_DIR = ROOT / "reports"
DASHBOARD_DIR = REPORT_DIR / "dashboard"
HISTORY_DIR = REPORT_DIR / "history"
V10_PATH = REPORT_DIR / "BD运营日报_问题清单版_V10_最新数据.xlsx"


def clean(value: Any) -> Any:
    if pd.isna(value):
        return ""
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return value


def records(frame: pd.DataFrame) -> list[dict[str, Any]]:
    if frame is None or frame.empty:
        return []
    return [{str(k): clean(v) for k, v in row.items()} for row in frame.to_dict("records")]


def esc(value: Any) -> str:
    value = clean(value)
    if isinstance(value, float):
        return html.escape(f"{value:,.2f}")
    return html.escape(str(value))


def html_table(frame: pd.DataFrame) -> str:
    if frame is None or frame.empty:
        return '<p class="empty">暂无数据</p>'
    headers = "".join(f"<th>{esc(col)}</th>" for col in frame.columns)
    rows = []
    for _, row in frame.iterrows():
        rows.append("<tr>" + "".join(f"<td>{esc(row[col])}</td>" for col in frame.columns) + "</tr>")
    return f'<div class="table-wrap"><table><thead><tr>{headers}</tr></thead><tbody>{"".join(rows)}</tbody></table></div>'


def find_source_file() -> Path:
    files = sorted(ROOT.glob("商户日报*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not files:
        raise FileNotFoundError("未找到商户日报*.xlsx")
    return files[0]


def read_sheet(path: Path, sheet_name: str) -> pd.DataFrame:
    try:
        return pd.read_excel(path, sheet_name=sheet_name)
    except Exception:
        return pd.DataFrame()


def text(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip().replace({"nan": "", "None": ""})


def number(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0)


def build_problem_lookup() -> pd.DataFrame:
    issues = read_sheet(V10_PATH, "运营问题清单")
    if issues.empty or "商户名" not in issues:
        return pd.DataFrame(columns=["商户名称", "问题类型", "优先级", "建议动作"])
    issues = issues.copy()
    issues["商户名称"] = text(issues["商户名"])
    for column in ["问题类型", "优先级", "整改建议"]:
        if column not in issues:
            issues[column] = ""
        issues[column] = text(issues[column])
    priority_rank = {"P0": 0, "P1": 1, "P2": 2}
    issues["_rank"] = issues["优先级"].map(priority_rank).fillna(9)

    def join_unique(values: pd.Series) -> str:
        out: list[str] = []
        for value in values.astype(str):
            value = value.strip()
            if value and value not in out:
                out.append(value)
        return " | ".join(out)

    grouped = (
        issues.sort_values(["商户名称", "_rank"])
        .groupby("商户名称", as_index=False)
        .agg(
            问题类型=("问题类型", join_unique),
            优先级=("优先级", join_unique),
            建议动作=("整改建议", join_unique),
        )
    )
    return grouped


def build_bd_merchant_details() -> tuple[pd.DataFrame, dict[str, pd.DataFrame], Path]:
    source = find_source_file()
    excel = pd.ExcelFile(source)
    sheet_name = "源数据" if "源数据" in excel.sheet_names else excel.sheet_names[0]
    raw = pd.read_excel(source, sheet_name=sheet_name)
    required = {
        "Q": "业务线含tn",
        "BI": "总订单",
        "DF": "代理商总补贴金额",
        "DL": "总b补代理商补贴金额",
        "DU": "餐饮b补_减配总补贴金额",
        "HW": "总c补代理商补贴金额",
        "LV": "餐饮爆单红包代理商补贴金额",
        "OS": "餐饮拼团代理商补贴金额",
        "RA": "一口价代理商补贴金额",
    }
    missing = [name for name in required.values() if name not in raw.columns]
    if missing:
        raise ValueError("日报缺少字段：" + "，".join(missing))

    frame = pd.DataFrame(
        {
            "商户名称": text(raw["商户名称"]),
            "BD名称": text(raw["bd名称"]).replace({"": "未分配BD"}),
            "Q列值": text(raw["业务线含tn"]),
            "总订单 BI": number(raw["总订单"]),
            "代理商总补贴金额 DF": number(raw["代理商总补贴金额"]),
            "B补 DL": number(raw["总b补代理商补贴金额"]),
            "C补 HW": number(raw["总c补代理商补贴金额"]),
            "减配 DU": number(raw["餐饮b补_减配总补贴金额"]),
            "爆单红包 LV": number(raw["餐饮爆单红包代理商补贴金额"]),
            "拼团 OS": number(raw["餐饮拼团代理商补贴金额"]),
            "一口价 RA": number(raw["一口价代理商补贴金额"]),
            "毛G": number(raw["毛g"]),
        }
    )
    frame = frame[frame["商户名称"].ne("") & frame["Q列值"].eq("FML")].copy()
    problems = build_problem_lookup()
    frame = frame.merge(problems, on="商户名称", how="left")
    for column in ["问题类型", "优先级", "建议动作"]:
        frame[column] = frame[column].fillna("")

    detail_columns = [
        "商户名称",
        "BD名称",
        "总订单 BI",
        "代理商总补贴金额 DF",
        "B补 DL",
        "C补 HW",
        "减配 DU",
        "爆单红包 LV",
        "拼团 OS",
        "一口价 RA",
        "毛G",
        "问题类型",
        "优先级",
        "建议动作",
    ]
    frame = frame[detail_columns].sort_values(["BD名称", "总订单 BI", "代理商总补贴金额 DF"], ascending=[True, False, False])

    summary_rows = []
    detail_by_bd: dict[str, pd.DataFrame] = {}
    for bd_name, group in frame.groupby("BD名称", dropna=False):
        bd_name = str(bd_name) if str(bd_name).strip() else "未分配BD"
        detail_by_bd[bd_name] = group.reset_index(drop=True)
        problem_mask = group["问题类型"].astype(str).str.strip().ne("")
        p0_mask = group["优先级"].astype(str).str.contains("P0", na=False)
        p1_mask = group["优先级"].astype(str).str.contains("P1", na=False)
        subsidy = float(group["代理商总补贴金额 DF"].sum())
        gmv = float(group["毛G"].sum())
        summary_rows.append(
            {
                "BD名称": bd_name,
                "FML商户数": int(group["商户名称"].nunique()),
                "总订单": float(group["总订单 BI"].sum()),
                "代理商总补贴": subsidy,
                "问题商户数": int(group.loc[problem_mask, "商户名称"].nunique()),
                "P0数量": int(group.loc[p0_mask, "商户名称"].nunique()),
                "P1数量": int(group.loc[p1_mask, "商户名称"].nunique()),
                "ROI": round(gmv / subsidy, 2) if subsidy else "",
            }
        )
    summary = pd.DataFrame(summary_rows).sort_values(["代理商总补贴", "总订单"], ascending=[False, False]).reset_index(drop=True)
    return summary, detail_by_bd, source


def build_bd_merchant_html(summary: pd.DataFrame, detail_by_bd: dict[str, pd.DataFrame]) -> str:
    parts = ['<h2>BD商户明细</h2><details open><summary>BD商户明细分析 <span>按Q=FML分组</span></summary>']
    parts.append(html_table(summary))
    for _, row in summary.iterrows():
        bd_name = str(row["BD名称"])
        title = (
            f"{bd_name}  FML商户{int(row['FML商户数'])}家  "
            f"总订单{int(row['总订单'])}  补贴{float(row['代理商总补贴']):,.2f}  "
            f"问题{int(row['问题商户数'])}"
        )
        parts.append(f"<details><summary>{esc(title)}</summary>{html_table(detail_by_bd.get(bd_name, pd.DataFrame()))}</details>")
    parts.append("</details>")
    return "".join(parts)


def archive_outputs() -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    archive = HISTORY_DIR / f"bd_merchant_detail_{stamp}"
    archive.mkdir(parents=True, exist_ok=True)
    for path in [
        DASHBOARD_DIR / "dashboard.html",
        DASHBOARD_DIR / "dashboard_data.json",
        DASHBOARD_DIR / "summary.json",
        DASHBOARD_DIR / "ai_report.json",
        DASHBOARD_DIR / "subsidy.json",
        REPORT_DIR / "dashboard.html",
    ]:
        if path.exists():
            shutil.copy2(path, archive / path.name)
    return archive


def main_cli() -> int:
    archive = archive_outputs()
    summary, detail_by_bd, source = build_bd_merchant_details()
    html_block = build_bd_merchant_html(summary, detail_by_bd)

    dashboard_html = DASHBOARD_DIR / "dashboard.html"
    root_dashboard_html = REPORT_DIR / "dashboard.html"
    if root_dashboard_html.exists() and (
        not dashboard_html.exists() or root_dashboard_html.stat().st_mtime >= dashboard_html.stat().st_mtime
    ):
        DASHBOARD_DIR.mkdir(parents=True, exist_ok=True)
        shutil.copy2(root_dashboard_html, dashboard_html)
    html_text = dashboard_html.read_text(encoding="utf-8")
    # Remove a previous generated merchant-detail block if present.
    start = html_text.find("<h2>BD商户明细</h2>")
    if start != -1:
        end = html_text.find("<script>", start)
        if end != -1:
            html_text = html_text[:start] + html_text[end:]
    insert_at = html_text.rfind("<script>")
    if insert_at == -1:
        insert_at = html_text.rfind("</main>")
    html_text = html_text[:insert_at] + html_block + html_text[insert_at:]
    dashboard_html.write_text(html_text, encoding="utf-8")
    root_dashboard_html.write_text(html_text, encoding="utf-8")

    data_path = DASHBOARD_DIR / "dashboard_data.json"
    data = json.loads(data_path.read_text(encoding="utf-8"))
    data["generated_at"] = datetime.now().isoformat(timespec="seconds")
    data["bd_merchant_summary"] = records(summary)
    data["bd_merchant_details"] = {bd: records(frame) for bd, frame in detail_by_bd.items()}
    data["bd_merchant_detail_source"] = {
        "source_file": str(source),
        "fml_rule": "Q列 业务线含tn = FML",
        "order_field": "BI 总订单",
        "agent_subsidy_field": "DF 代理商总补贴金额",
    }
    data_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    summary_path = DASHBOARD_DIR / "summary.json"
    summary_data = json.loads(summary_path.read_text(encoding="utf-8")) if summary_path.exists() else {}
    summary_data["bd_merchant_summary"] = records(summary)
    summary_data["bd_merchant_detail_source"] = data["bd_merchant_detail_source"]
    summary_path.write_text(json.dumps(summary_data, ensure_ascii=False, indent=2), encoding="utf-8")

    result = {
        "archive_dir": str(archive),
        "source_file": str(source),
        "dashboard": str(dashboard_html),
        "dashboard_data": str(data_path),
        "bd_count": int(len(summary)),
        "bd_summary": records(summary),
        "dashboard_updated_at": datetime.fromtimestamp(dashboard_html.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main_cli())
