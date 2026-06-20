#!/usr/bin/env python3
"""Rebuild the local dashboard from the latest V10 report and subsidy table.

This script only regenerates reports/dashboard/* and the subsidy reconcile
workbook. It does not send DingTalk messages and does not modify source data.
"""

from __future__ import annotations

import html
import json
import math
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
REPORT_DIR = ROOT / "reports"
DASHBOARD_DIR = REPORT_DIR / "dashboard"
HISTORY_DIR = REPORT_DIR / "history"
RECONCILE_PATH = REPORT_DIR / "subsidy_reconcile_check.xlsx"

sys.path.insert(0, str(ROOT))
import main  # noqa: E402


def clean_value(value: Any) -> Any:
    if value is pd.NA:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return value


def records(df: pd.DataFrame) -> list[dict[str, Any]]:
    if df is None or df.empty:
        return []
    return [{str(k): clean_value(v) for k, v in row.items()} for row in df.to_dict("records")]


def fmt(value: Any) -> str:
    value = clean_value(value)
    if isinstance(value, float):
        return f"{value:,.2f}"
    return str(value)


def html_escape(value: Any) -> str:
    return html.escape(fmt(value))


def html_table(df: pd.DataFrame) -> str:
    if df is None or df.empty:
        return '<p class="empty">暂无数据</p>'
    headers = "".join(f"<th>{html_escape(col)}</th>" for col in df.columns)
    rows = []
    for _, row in df.iterrows():
        rows.append("<tr>" + "".join(f"<td>{html_escape(row[col])}</td>" for col in df.columns) + "</tr>")
    return f'<div class="table-wrap"><table><thead><tr>{headers}</tr></thead><tbody>{"".join(rows)}</tbody></table></div>'


def section(title: str, df: pd.DataFrame, open_by_default: bool = False) -> str:
    open_attr = " open" if open_by_default else ""
    count = 0 if df is None else len(df)
    return f"<details{open_attr}><summary>{html_escape(title)} <span>{count}条</span></summary>{html_table(df)}</details>"


def latest_v10_report() -> Path:
    preferred = REPORT_DIR / "BD运营日报_问题清单版_V10_最新数据.xlsx"
    if preferred.exists():
        return preferred
    candidates = sorted(REPORT_DIR.glob("BD运营日报_问题清单版_V10*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not candidates:
        raise FileNotFoundError("未找到 BD运营日报_问题清单版_V10*.xlsx")
    return candidates[0]


def read_sheet(path: Path, sheet_name: str) -> pd.DataFrame:
    try:
        return pd.read_excel(path, sheet_name=sheet_name)
    except Exception:
        return pd.DataFrame()


def fml_overview(subsidy_tables: dict[str, pd.DataFrame]) -> pd.DataFrame:
    overview = subsidy_tables.get("补贴总览", pd.DataFrame()).copy()
    if overview.empty or "口径" not in overview:
        return overview
    return overview[overview["口径"].astype(str).eq("FML商户")].reset_index(drop=True)


def activity_structure_from_top10(subsidy_tables: dict[str, pd.DataFrame]) -> pd.DataFrame:
    rows = []
    total = 0.0
    for sheet_name, _ in main.ACTIVITY_TOP_CONFIG:
        activity_name = sheet_name.replace("TOP10", "")
        table = subsidy_tables.get(sheet_name, pd.DataFrame())
        amount = float(pd.to_numeric(table.get("活动补贴金额", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()) if not table.empty else 0.0
        merchant_count = int(table.get("商户名称", pd.Series(dtype=str)).astype(str).nunique()) if not table.empty and "商户名称" in table else 0
        rows.append({"活动": activity_name, "补贴金额": round(amount, 2), "TOP10商户数": merchant_count})
        total += amount
    for row in rows:
        row["占比"] = row["补贴金额"] / total if total else 0
    return pd.DataFrame(rows, columns=["活动", "补贴金额", "TOP10商户数", "占比"])


def fml_bd_subsidy_rank() -> pd.DataFrame:
    subsidy_info = main.find_subsidy_file()
    if not subsidy_info:
        return pd.DataFrame()
    df = main.load_subsidy_data(subsidy_info.path)
    fml = df[df["商户类型"].astype(str).eq("FML")].copy()
    if fml.empty:
        return pd.DataFrame()
    rank = (
        fml.groupby("bd名称", dropna=False)
        .agg(
            总补贴=("总补贴", "sum"),
            平台补贴=("平台补贴", "sum"),
            商户补贴=("商户补贴", "sum"),
            代理商补贴=("代理商补贴", "sum"),
            爆单红包补贴=("爆单红包补贴", "sum"),
            拼团补贴=("拼团补贴", "sum"),
            一口价补贴=("一口价补贴", "sum"),
            毛G=("毛G", "sum"),
            总订单=("总订单", "sum"),
        )
        .reset_index()
    )
    rank["补贴率"] = (rank["总补贴"] / rank["毛G"].replace(0, pd.NA)).fillna(0)
    return rank.sort_values("总补贴", ascending=False).reset_index(drop=True)


def parse_money_from_text(value: Any, pattern: str) -> float:
    text = str(clean_value(value))
    matched = re.search(pattern, text)
    if not matched:
        return 0.0
    return float(matched.group(1).replace(",", ""))


def build_bd_issue_amount_rank(v10_path: Path) -> pd.DataFrame:
    issues = read_sheet(v10_path, "运营问题清单")
    if issues.empty or "BD名称" not in issues:
        return pd.DataFrame()
    frame = issues.copy()
    frame["BD名称"] = frame["BD名称"].fillna("未分配BD").astype(str).replace({"": "未分配BD"})
    frame["GMV损失_数值"] = pd.to_numeric(frame.get("GMV损失", 0), errors="coerce").fillna(0)
    frame["高补低产涉及补贴"] = frame.apply(
        lambda row: parse_money_from_text(row.get("数据变化", ""), r"补贴表总补贴：([\d,.]+)")
        if str(row.get("问题类型", "")) == "高补低产商户"
        else 0.0,
        axis=1,
    )
    frame["补贴异常涉及补贴"] = frame.apply(
        lambda row: parse_money_from_text(row.get("数据变化", ""), r"补贴：([\d,.]+)")
        if str(row.get("问题类型", "")) == "补贴异常商户"
        else 0.0,
        axis=1,
    )
    grouped = (
        frame.groupby("BD名称", as_index=False)
        .agg(
            问题商户数=("商户名", "nunique"),
            GMV损失=("GMV损失_数值", "sum"),
            高补低产涉及补贴=("高补低产涉及补贴", "sum"),
            补贴异常涉及补贴=("补贴异常涉及补贴", "sum"),
        )
    )
    grouped["问题金额合计"] = grouped["GMV损失"] + grouped["高补低产涉及补贴"] + grouped["补贴异常涉及补贴"]
    return grouped.sort_values("问题金额合计", ascending=False).reset_index(drop=True)


def build_reconcile(
    subsidy_tables: dict[str, pd.DataFrame],
    activity_structure: pd.DataFrame,
    bd_rank: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    overview = fml_overview(subsidy_tables)
    total_merchants = 0
    fml_merchants = 0
    non_fml_merchants = 0
    subsidy_info = main.find_subsidy_file()
    if subsidy_info:
        df = main.load_subsidy_data(subsidy_info.path)
        total_merchants = int(df["商户名称"].nunique())
        fml_merchants = int(df[df["商户类型"].astype(str).eq("FML")]["商户名称"].nunique())
        non_fml_merchants = total_merchants - fml_merchants

    def overview_amount(item: str) -> float:
        if overview.empty:
            return 0.0
        matched = overview[overview["项目"].astype(str).eq(item)]
        return round(float(pd.to_numeric(matched.get("金额", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()), 2)

    activity_amounts = {str(row["活动"]): round(float(row["补贴金额"]), 2) for _, row in activity_structure.iterrows()}
    chart_total = round(float(activity_structure["补贴金额"].sum()), 2) if not activity_structure.empty else 0.0
    top10_total = chart_total
    subsidy_total = overview_amount("补贴总额")
    reason = "已统一：活动补贴结构图与活动TOP10均使用FML商户TOP10明细金额；补贴总览金额为全部FML商户补贴总额。"
    summary = pd.DataFrame(
        [
            {
                "总商户数": total_merchants,
                "FML商户数": fml_merchants,
                "非FML商户数": non_fml_merchants,
                "B补金额": activity_amounts.get("B补", 0.0),
                "C补金额": activity_amounts.get("C补", 0.0),
                "减配金额": activity_amounts.get("减配", 0.0),
                "爆单红包金额": activity_amounts.get("爆单红包", 0.0),
                "拼团金额": activity_amounts.get("拼团", 0.0),
                "一口价金额": activity_amounts.get("一口价", 0.0),
                "活动补贴结构图金额": chart_total,
                "TOP10金额": top10_total,
                "补贴总览金额": subsidy_total,
                "差异金额": round(chart_total - top10_total, 2),
                "差异原因": reason,
            }
        ]
    )
    rows = []
    overview_map = {
        "B补": "B补代理商补贴",
        "C补": "C补代理商补贴",
        "减配": "减配补贴",
        "爆单红包": "爆单红包补贴",
        "拼团": "拼团补贴",
        "一口价": "一口价补贴",
    }
    for activity_name, overview_name in overview_map.items():
        chart_amount = activity_amounts.get(activity_name, 0.0)
        rows.append(
            {
                "校验项": activity_name,
                "总商户数": total_merchants,
                "FML商户数": fml_merchants,
                "非FML商户数": non_fml_merchants,
                "活动补贴结构图金额": chart_amount,
                "TOP10金额": chart_amount,
                "补贴总览金额": overview_amount(overview_name),
                "差异金额": 0.0,
                "差异原因": reason,
            }
        )
    return summary, pd.DataFrame(rows)


def write_reconcile(path: Path, summary: pd.DataFrame, detail: pd.DataFrame) -> None:
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="汇总", index=False)
        detail.to_excel(writer, sheet_name="补贴口径校验", index=False)


def chart_section(activity_structure: pd.DataFrame) -> str:
    chart_rows = [
        {"name": str(row["活动"]), "value": float(row["补贴金额"]), "percent": float(row["占比"])}
        for _, row in activity_structure.iterrows()
    ]
    chart_json = json.dumps(chart_rows, ensure_ascii=False)
    return f"""
<details open><summary>活动补贴结构饼图（FML TOP10同口径） <span>{len(activity_structure)}条</span></summary>
<div class="chart-wrap"><div id="activityPie" class="chart"></div></div>
{html_table(activity_structure)}
</details>
<script>
(function() {{
  var data = {chart_json};
  var el = document.getElementById('activityPie');
  if (!el || !window.echarts) return;
  var chart = echarts.init(el);
  chart.setOption({{
    tooltip: {{ trigger: 'item', formatter: function(p) {{
      var d = p.data || {{}};
      var pct = typeof d.percent === 'number' ? d.percent * 100 : p.percent;
      return [d.name, Number(d.value || 0).toLocaleString('zh-CN', {{minimumFractionDigits:2, maximumFractionDigits:2}}) + '元', pct.toFixed(2) + '%'].join('<br>');
    }} }},
    legend: {{ bottom: 0, type: 'scroll' }},
    series: [{{ name:'活动补贴结构', type:'pie', radius:['38%','68%'], center:['50%','43%'], data:data,
      label: {{ formatter: function(p) {{ return p.name + '\\n' + p.percent.toFixed(2) + '%'; }} }}
    }}]
  }});
  window.addEventListener('resize', function() {{ chart.resize(); }});
}})();
</script>
"""


def activity_top_sections(subsidy_tables: dict[str, pd.DataFrame]) -> str:
    parts = []
    for index, (sheet_name, _) in enumerate(main.ACTIVITY_TOP_CONFIG):
        parts.append(section(sheet_name, subsidy_tables.get(sheet_name, pd.DataFrame()), index == 0))
    return f"<details open><summary>活动补贴TOP10商户分析（仅FML商户） <span>{len(parts)}类</span></summary>{''.join(parts)}</details>"


def write_dashboard_html(
    v10_path: Path,
    subsidy_tables: dict[str, pd.DataFrame],
    activity_structure: pd.DataFrame,
    bd_rank: pd.DataFrame,
    summary_reconcile: pd.DataFrame,
) -> Path:
    DASHBOARD_DIR.mkdir(parents=True, exist_ok=True)
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    bd_risk = read_sheet(v10_path, "BD风险排行")
    bd_analysis = read_sheet(v10_path, "BD问题结构分析")
    bd_issue = build_bd_issue_amount_rank(v10_path)
    bd_roi = read_sheet(v10_path, "代理商补贴效率分析")
    overview = read_sheet(v10_path, "今日运营结论")
    city_health = read_sheet(v10_path, "城市运营健康度")
    issue_amount = read_sheet(v10_path, "问题金额统计")
    closure = read_sheet(v10_path, "问题闭环驾驶舱")
    ai_report = read_sheet(v10_path, "经营诊断结论")
    data_source = read_sheet(v10_path, "数据来源")
    html_content = f"""<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>BD运营日报 Dashboard</title><script src="https://cdn.jsdelivr.net/npm/echarts@5/dist/echarts.min.js"></script>
<style>body{{margin:0;background:#f5f7fb;color:#172033;font:14px/1.55 -apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif}}header{{padding:24px 32px;background:#111827;color:#fff}}h1{{margin:0 0 6px;font-size:26px}}.meta{{color:#cbd5e1}}main{{max-width:1320px;margin:0 auto;padding:20px}}.cards{{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;margin-bottom:16px}}.card,details{{background:#fff;border:1px solid #d9e0ea;border-radius:8px;box-shadow:0 1px 2px rgba(16,24,40,.04)}}.card{{padding:14px}}.card span{{display:block;color:#667085;font-size:12px}}.card strong{{display:block;font-size:20px;margin-top:4px}}details{{margin-bottom:14px;overflow:hidden}}summary{{padding:13px 16px;cursor:pointer;font-weight:700;font-size:16px;border-bottom:1px solid #d9e0ea}}summary span{{color:#667085;font-size:12px;margin-left:8px}}.table-wrap{{overflow:auto;max-height:70vh}}table{{width:100%;border-collapse:collapse;min-width:760px}}th,td{{padding:9px 11px;border-bottom:1px solid #edf1f6;text-align:left;white-space:nowrap;vertical-align:top}}th{{position:sticky;top:0;background:#f8fafc;z-index:1}}.chart{{width:100%;height:390px;min-height:300px}}.chart-wrap{{padding:16px}}.empty{{color:#667085;padding:16px;margin:0}}h2{{margin:24px 0 12px}}</style>
</head><body><header><h1>BD运营日报 Dashboard</h1><div class="meta">Dashboard更新时间：{html_escape(generated_at)}｜数据文件：{html_escape(v10_path.name)}</div></header><main>
<div class="cards">
<div class="card"><span>FML商户数</span><strong>{html_escape(summary_reconcile.at[0, "FML商户数"])}</strong></div>
<div class="card"><span>活动结构图金额</span><strong>{html_escape(summary_reconcile.at[0, "活动补贴结构图金额"])}</strong></div>
<div class="card"><span>FML补贴总额</span><strong>{html_escape(summary_reconcile.at[0, "补贴总览金额"])}</strong></div>
</div>
{section("一、数据源检查", data_source, True)}
<h2>BD板块</h2>
{section("1. BD风险排行", bd_risk, True)}
{section("2. BD经营分析", bd_analysis, True)}
{section("3. BD补贴排行", bd_rank, True)}
{section("4. BD问题金额排行", bd_issue, True)}
{section("5. BD ROI排行", bd_roi, True)}
<h2>经营与问题</h2>
{section("经营概览", overview, True)}
{section("城市健康度", city_health, True)}
{section("问题金额分析", issue_amount, True)}
{section("问题闭环驾驶舱", closure, True)}
{section("AI经营结论", ai_report, True)}
<h2>补贴分析（全部FML口径）</h2>
{section("补贴总览（仅FML商户）", fml_overview(subsidy_tables), True)}
{chart_section(activity_structure)}
{activity_top_sections(subsidy_tables)}
{section("补贴口径校验", summary_reconcile, True)}
{section("商户补贴TOP20（仅FML商户）", subsidy_tables.get("商户补贴TOP20", pd.DataFrame()), True)}
{section("补贴异常TOP20（仅FML商户）", subsidy_tables.get("异常补贴预警TOP20", pd.DataFrame()), True)}
</main></body></html>"""
    dashboard_path = DASHBOARD_DIR / "dashboard.html"
    dashboard_path.write_text(html_content, encoding="utf-8")
    (REPORT_DIR / "dashboard.html").write_text(html_content, encoding="utf-8")
    return dashboard_path


def write_dashboard_json(
    v10_path: Path,
    subsidy_tables: dict[str, pd.DataFrame],
    activity_structure: pd.DataFrame,
    bd_rank: pd.DataFrame,
    summary_reconcile: pd.DataFrame,
) -> Path:
    bd_risk = read_sheet(v10_path, "BD风险排行")
    bd_analysis = read_sheet(v10_path, "BD问题结构分析")
    bd_issue = build_bd_issue_amount_rank(v10_path)
    bd_roi = read_sheet(v10_path, "代理商补贴效率分析")
    activity_top10 = {sheet: records(subsidy_tables.get(sheet, pd.DataFrame())) for sheet, _ in main.ACTIVITY_TOP_CONFIG}
    subsidy_payload = {
        "overview": records(fml_overview(subsidy_tables)),
        "activity_structure": records(activity_structure),
        "activity_top10": activity_top10,
        "merchant_top20": records(subsidy_tables.get("商户补贴TOP20", pd.DataFrame())),
        "abnormal_top20": records(subsidy_tables.get("异常补贴预警TOP20", pd.DataFrame())),
        "reconcile": records(summary_reconcile),
    }
    payload = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "v10_file": str(v10_path.resolve()),
        "bd_risk": records(bd_risk),
        "bd_analysis": records(bd_analysis),
        "bd_rank": records(bd_rank),
        "bd_subsidy_rank": records(bd_rank),
        "bd_issue_amount_rank": records(bd_issue),
        "bd_roi_rank": records(bd_roi),
        "overview": records(read_sheet(v10_path, "今日运营结论")),
        "city_health": records(read_sheet(v10_path, "城市运营健康度")),
        "issue_amount": records(read_sheet(v10_path, "问题金额统计")),
        "closure": records(read_sheet(v10_path, "问题闭环驾驶舱")),
        "ai_report": records(read_sheet(v10_path, "经营诊断结论")),
        "subsidy": subsidy_payload,
        "subsidy_overview_fml": subsidy_payload["overview"],
        "activity_structure": subsidy_payload["activity_structure"],
        "activity_reconcile": subsidy_payload["reconcile"],
    }
    DASHBOARD_DIR.mkdir(parents=True, exist_ok=True)
    path = DASHBOARD_DIR / "dashboard_data.json"
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    (DASHBOARD_DIR / "summary.json").write_text(json.dumps({"generated_at": payload["generated_at"], "overview": payload["overview"]}, ensure_ascii=False, indent=2), encoding="utf-8")
    (DASHBOARD_DIR / "ai_report.json").write_text(json.dumps({"generated_at": payload["generated_at"], "ai_report": payload["ai_report"]}, ensure_ascii=False, indent=2), encoding="utf-8")
    (DASHBOARD_DIR / "subsidy.json").write_text(json.dumps(subsidy_payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def archive_current_outputs() -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    archive = HISTORY_DIR / f"dashboard_fix_{stamp}"
    archive.mkdir(parents=True, exist_ok=True)
    for path in [
        DASHBOARD_DIR / "dashboard.html",
        DASHBOARD_DIR / "dashboard_data.json",
        DASHBOARD_DIR / "summary.json",
        DASHBOARD_DIR / "ai_report.json",
        DASHBOARD_DIR / "subsidy.json",
        REPORT_DIR / "dashboard.html",
        RECONCILE_PATH,
    ]:
        if path.exists():
            shutil.copy2(path, archive / path.name)
    return archive


def main_cli() -> int:
    archive = archive_current_outputs()
    v10_path = latest_v10_report()
    subsidy_tables = main.build_subsidy_analysis()
    activity_structure = activity_structure_from_top10(subsidy_tables)
    bd_rank = fml_bd_subsidy_rank()
    summary_reconcile, detail_reconcile = build_reconcile(subsidy_tables, activity_structure, bd_rank)
    write_reconcile(RECONCILE_PATH, summary_reconcile, detail_reconcile)
    dashboard_path = write_dashboard_html(v10_path, subsidy_tables, activity_structure, bd_rank, summary_reconcile)
    dashboard_data_path = write_dashboard_json(v10_path, subsidy_tables, activity_structure, bd_rank, summary_reconcile)
    result = {
        "archive_dir": str(archive.resolve()),
        "dashboard": str(dashboard_path.resolve()),
        "dashboard_data": str(dashboard_data_path.resolve()),
        "reconcile": str(RECONCILE_PATH.resolve()),
        "bd_modules": {
            "bd_risk": len(read_sheet(v10_path, "BD风险排行")),
            "bd_analysis": len(read_sheet(v10_path, "BD问题结构分析")),
            "bd_rank": len(bd_rank),
            "bd_issue_amount_rank": len(read_sheet(v10_path, "BD风险排行")),
            "bd_roi_rank": len(read_sheet(v10_path, "代理商补贴效率分析")),
        },
        "fml_merchants": int(summary_reconcile.at[0, "FML商户数"]),
        "subsidy_total_fml": float(summary_reconcile.at[0, "补贴总览金额"]),
        "activity_chart_total": float(summary_reconcile.at[0, "活动补贴结构图金额"]),
        "top10_total": float(summary_reconcile.at[0, "TOP10金额"]),
        "consistent": float(summary_reconcile.at[0, "差异金额"]) == 0.0,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main_cli())
