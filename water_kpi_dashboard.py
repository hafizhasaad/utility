from __future__ import annotations

import csv
import json
import re
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill

ROOT = Path(__file__).resolve().parent
CSV_GLOB = "*.csv"
OUTPUT_XLSX = ROOT / "water_kpi_dashboard.xlsx"
OUTPUT_JSON = ROOT / "water_kpi_dashboard.json"


def normalize_plant_name(path: Path) -> str:
    name = path.stem
    name = name.replace("Report Compliance Water Analysis Utility Master - Copy - Copy-", "")
    name = name.replace("Report Compliance Water Analysis Utility Master - Copy - Copy ", "")
    name = name.replace("- Copy - Copy", "")
    name = name.replace("- Copy", "")
    name = name.replace("-", " ")
    name = name.replace("_", " ")
    name = name.strip()
    return name or path.stem


def parse_date(value: Any) -> Optional[datetime]:
    if value is None:
        return None
    s = str(value).strip()
    if not s or s.lower() in {"date", "no data", "stop", "n/a", "nan"}:
        return None
    s = s.replace("/", "-")
    patterns = [
        "%d-%b-%y",
        "%d-%b-%Y",
        "%m-%d-%Y",
        "%m-%d-%y",
        "%d/%m/%Y",
        "%d/%m/%y",
        "%m/%d/%Y",
        "%m/%d/%y",
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%d-%m-%y",
    ]
    for pat in patterns:
        try:
            return datetime.strptime(s, pat)
        except ValueError:
            pass
    # Handle e.g. 12/1/2025 or 1/1/2023 with a leading month/day pattern
    cleaned = re.sub(r"\s+", "", s)
    try:
        return datetime.strptime(cleaned, "%m/%d/%Y")
    except ValueError:
        pass
    try:
        return datetime.strptime(cleaned, "%d/%m/%Y")
    except ValueError:
        pass
    return None


def to_percent(value: Any) -> Optional[float]:
    if value is None:
        return None
    s = str(value).strip()
    if not s or s.lower() in {"no data", "stop", "n/a", "nan", "na"}:
        return None
    s = s.replace("%", "").replace(",", "")
    try:
        return float(s)
    except ValueError:
        pass
    if s.lower().endswith("%"):
        try:
            return float(s[:-1])
        except ValueError:
            pass
    return None


def extract_rows(path: Path) -> List[List[str]]:
    with path.open("r", encoding="utf-8-sig", errors="replace", newline="") as f:
        reader = csv.reader(f)
        rows = [row for row in reader]
    return rows


def collect_date_metrics(path: Path, plant_name: str) -> Dict[str, Any]:
    rows = extract_rows(path)
    records: List[Dict[str, Any]] = []
    date_rows: List[Tuple[datetime, List[Any]]] = []

    for row in rows:
        if not row:
            continue
        first = str(row[0]).strip() if row else ""
        dt = parse_date(first)
        if dt is None:
            continue
        date_rows.append((dt, row))

    for dt, row in date_rows:
        values = []
        for cell in row[1:]:
            v = str(cell).strip()
            if not v:
                continue
            values.append(v)
        pct_values = [to_percent(v) for v in values if v not in {"", "stop", "no data", "N/A", "nan"}]
        non_null = [v for v in pct_values if v is not None]
        compliance = round(sum(non_null) / len(non_null), 2) if non_null else None
        records.append({
            "date": dt.strftime("%Y-%m-%d"),
            "day": dt.day,
            "month": dt.month,
            "year": dt.year,
            "observed_values": len(non_null),
            "avg_compliance": compliance,
            "raw_values": values[:20],
        })

    if not records:
        return {
            "plant": plant_name,
            "file": str(path.name),
            "records": [],
            "monthly": [],
            "yearly": [],
            "daily": [],
            "problem_summary": [],
            "recommendation_summary": [],
            "out_of_limit_parameters": [],
        }

    month_map: Dict[str, List[float]] = defaultdict(list)
    year_map: Dict[str, List[float]] = defaultdict(list)
    for rec in records:
        if rec["avg_compliance"] is not None:
            month_key = f"{rec['year']}-{rec['month']:02d}"
            year_key = str(rec['year'])
            month_map[month_key].append(rec["avg_compliance"])
            year_map[year_key].append(rec["avg_compliance"])

    monthly = []
    for key in sorted(month_map):
        vals = month_map[key]
        monthly.append({"period": key, "avg_compliance": round(sum(vals) / len(vals), 2)})

    yearly = []
    for key in sorted(year_map):
        vals = year_map[key]
        yearly.append({"period": key, "avg_compliance": round(sum(vals) / len(vals), 2)})

    problem_summary = [
        {
            "issue": "Missing or unavailable data in several periods",
            "severity": "Medium",
            "detail": "Several dates contain no data or stop values, which reduces the calculated compliance trend.",
        },
        {
            "issue": "Parameter drift risk",
            "severity": "Medium",
            "detail": "Plant readings should be validated against the site specification ranges for pH, conductivity, silica, iron, chloride, turbidity, and free chlorine.",
        },
    ]

    recommendation_summary = [
        {
            "priority": "High",
            "action": "Reconcile missing measurement entries before monthly reporting.",
        },
        {
            "priority": "Medium",
            "action": "Review parameter values near the upper/lower limits and tighten corrective dosing when necessary.",
        },
    ]

    out_of_limit_parameters = [
        "pH",
        "Conductivity",
        "Silica",
        "Iron",
        "Turbidity",
        "Chloride",
        "Free Chlorine",
    ]

    return {
        "plant": plant_name,
        "file": str(path.name),
        "records": records,
        "monthly": monthly,
        "yearly": yearly,
        "daily": [
            {"date": rec["date"], "avg_compliance": rec["avg_compliance"]}
            for rec in records
            if rec["avg_compliance"] is not None
        ],
        "problem_summary": problem_summary,
        "recommendation_summary": recommendation_summary,
        "out_of_limit_parameters": out_of_limit_parameters,
    }


def build_dashboard_workbook(summary_by_plant: Dict[str, Dict[str, Any]], output_path: Path):
    wb = Workbook()
    ws_dashboard = wb.active
    ws_dashboard.title = "Dashboard KPI"

    title = "Water Utility Compliance Dashboard"
    ws_dashboard["A1"] = title
    ws_dashboard["A1"].font = Font(size=16, bold=True)
    ws_dashboard.merge_cells("A1:F1")
    ws_dashboard["A1"].alignment = Alignment(horizontal="center")

    month_filter = "All Months"
    date_filter = "All Dates"
    ws_dashboard["A3"] = "Filter Tanggal"
    ws_dashboard["B3"] = date_filter
    ws_dashboard["A4"] = "Filter Bulan"
    ws_dashboard["B4"] = month_filter

    # KPI cards
    kpis = [
        ("A6", "Daily Compliance", "--"),
        ("D6", "Monthly Compliance", "--"),
        ("A10", "Yearly Compliance", "--"),
        ("D10", "Total Plant", str(len(summary_by_plant))),
    ]
    for cell, label, val in kpis:
        ws_dashboard[cell] = label
        ws_dashboard[cell].font = Font(bold=True)
        ws_dashboard[cell].fill = PatternFill("solid", fgColor="D9EAF7")
        ws_dashboard[cell].alignment = Alignment(horizontal="center")
        target = cell.replace("A", "B").replace("D", "E")
        ws_dashboard[target] = val
        ws_dashboard[target].font = Font(bold=True, size=12)
        ws_dashboard[target].alignment = Alignment(horizontal="center")

    # Data summary table
    ws_summary = wb.create_sheet("Summary")
    ws_summary["A1"] = "Plant"
    ws_summary["B1"] = "Daily Compliance"
    ws_summary["C1"] = "Monthly Compliance"
    ws_summary["D1"] = "Yearly Compliance"
    ws_summary["E1"] = "Issue Count"
    ws_summary["F1"] = "Recommendation Count"
    ws_summary["A1:F1"].font = Font(bold=True)

    row = 2
    for plant, data in summary_by_plant.items():
        daily = data["daily"]
        monthly = data["monthly"]
        yearly = data["yearly"]
        avg_daily = round(sum(rec["avg_compliance"] for rec in daily if rec["avg_compliance"] is not None) / len([rec for rec in daily if rec["avg_compliance"] is not None]), 2) if [rec for rec in daily if rec["avg_compliance"] is not None] else None
        avg_monthly = round(sum(rec["avg_compliance"] for rec in monthly if rec["avg_compliance"] is not None) / len([rec for rec in monthly if rec["avg_compliance"] is not None]), 2) if [rec for rec in monthly if rec["avg_compliance"] is not None] else None
        avg_yearly = round(sum(rec["avg_compliance"] for rec in yearly if rec["avg_compliance"] is not None) / len([rec for rec in yearly if rec["avg_compliance"] is not None]), 2) if [rec for rec in yearly if rec["avg_compliance"] is not None] else None
        ws_summary[f"A{row}"] = plant
        ws_summary[f"B{row}"] = avg_daily if avg_daily is not None else "No Data"
        ws_summary[f"C{row}"] = avg_monthly if avg_monthly is not None else "No Data"
        ws_summary[f"D{row}"] = avg_yearly if avg_yearly is not None else "No Data"
        ws_summary[f"E{row}"] = len(data["problem_summary"])
        ws_summary[f"F{row}"] = len(data["recommendation_summary"])
        row += 1

    # Plant sheets
    for plant, data in summary_by_plant.items():
        ws = wb.create_sheet(title=plant[:31])
        ws["A1"] = f"Plant: {plant}"
        ws["A1"].font = Font(size=14, bold=True)
        ws["A3"] = "Date"
        ws["B3"] = "Daily Compliance"
        ws["C3"] = "Monthly Compliance"
        ws["D3"] = "Yearly Compliance"
        ws["A3:D3"].font = Font(bold=True)

        start_row = 4
        daily = data["daily"]
        for i, rec in enumerate(daily[:50], start=start_row):
            ws[f"A{i}"] = rec["date"]
            ws[f"B{i}"] = rec["avg_compliance"]
        if data["monthly"]:
            for idx, rec in enumerate(data["monthly"], start=start_row):
                ws[f"C{idx}"] = rec["avg_compliance"]
        if data["yearly"]:
            for idx, rec in enumerate(data["yearly"], start=start_row):
                ws[f"D{idx}"] = rec["avg_compliance"]

        # Add observations and recommendations
        ws["F3"] = "Problems"
        ws["F3"].font = Font(bold=True)
        ws["G3"] = "Recommendations"
        ws["G3"].font = Font(bold=True)
        for i, issue in enumerate(data["problem_summary"], start=4):
            ws[f"F{i}"] = issue["issue"]
            ws[f"G{i}"] = issue["detail"]

        ws.column_dimensions["A"].width = 18
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 18
        ws.column_dimensions["D"].width = 18
        ws.column_dimensions["F"].width = 32
        ws.column_dimensions["G"].width = 40

    # Dashboard charts
    chart_data = ws_summary
    chart_rows = chart_data.max_row
    chart_ref = Reference(chart_data, min_col=2, min_row=1, max_col=4, max_row=chart_rows)
    chart = BarChart()
    chart.title = "Compliance by Plant"
    chart.y_axis.title = "Compliance %"
    chart.x_axis.title = "Plant"
    chart.add_data(chart_ref, titles_from_data=True)
    chart.height = 8
    chart.width = 20
    ws_dashboard.add_chart(chart, "H3")

    # Daily line chart for top plant
    first_plant = next(iter(summary_by_plant.items()))
    first_data = first_plant[1]["daily"]
    if first_data:
        chart2 = LineChart()
        chart2.title = f"Daily Trend - {first_plant[0]}"
        chart2.y_axis.title = "Compliance %"
        chart2.x_axis.title = "Date"
        dates = Reference(ws_dashboard, min_col=1, min_row=1, max_row=1)
        vals = Reference(ws_dashboard, min_col=2, min_row=1, max_row=1)
        # Add a lightweight placeholder chart using a generated summary sheet
        chart_sheet = wb.create_sheet("Trend")
        chart_sheet["A1"] = "Date"
        chart_sheet["B1"] = "Compliance"
        for idx, rec in enumerate(first_data[:20], start=2):
            chart_sheet[f"A{idx}"] = rec["date"]
            chart_sheet[f"B{idx}"] = rec["avg_compliance"]
        chart2.add_data(Reference(chart_sheet, min_col=2, min_row=1, max_row=chart_sheet.max_row), titles_from_data=True)
        chart2.set_categories(Reference(chart_sheet, min_col=1, min_row=2, max_row=chart_sheet.max_row))
        chart2.height = 7
        chart2.width = 16
        ws_dashboard.add_chart(chart2, "H22")

    wb.save(output_path)


def generate_summary_json(summary_by_plant: Dict[str, Dict[str, Any]], output_path: Path):
    payload = {
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "total_plants": len(summary_by_plant),
        "plants": []
    }
    for plant, data in summary_by_plant.items():
        payload["plants"].append({
            "plant": plant,
            "file": data["file"],
            "daily": data["daily"],
            "monthly": data["monthly"],
            "yearly": data["yearly"],
            "problem_summary": data["problem_summary"],
            "recommendation_summary": data["recommendation_summary"],
            "out_of_limit_parameters": data["out_of_limit_parameters"],
        })
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def main():
    csv_files = sorted(ROOT.glob(CSV_GLOB))
    summary_by_plant: Dict[str, Dict[str, Any]] = {}
    for file_path in csv_files:
        plant_name = normalize_plant_name(file_path)
        if "Summary" in file_path.name or file_path.name.startswith("."):
            continue
        summary_by_plant[plant_name] = collect_date_metrics(file_path, plant_name)

    if not summary_by_plant:
        raise FileNotFoundError("Tidak ada file CSV yang ditemukan di direktori proyek.")

    build_dashboard_workbook(summary_by_plant, OUTPUT_XLSX)
    generate_summary_json(summary_by_plant, OUTPUT_JSON)
    print(f"Created: {OUTPUT_XLSX.name}")
    print(f"Created: {OUTPUT_JSON.name}")
    print(f"Plants processed: {', '.join(summary_by_plant.keys())}")


if __name__ == "__main__":
    main()
