from pathlib import Path
import re
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter


ROOT = Path(__file__).resolve().parent
CSV_PATHS = sorted(ROOT.glob("Report Compliance Water Analysis Utility Master - Copy - Copy-*.csv"))

PLANT_MAP = {
    "BOILER&BFW": "Boiler & BFW",
    "BWRO": "BWRO",
    "COOLING TOWER": "Cooling Tower",
    "DEMIN WATER": "Demin Water",
    "LAY UP BOILER": "Lay Up Boiler",
    "LAY UP CT": "Lay Up CT",
    "RAW WATER": "Raw Water",
    "RO WATER": "RO Water",
    "SWRO": "SWRO",
}

OUT_OF_LIMIT_BY_PLANT = {
    "Boiler & BFW": ["pH", "Conductivity", "Total Hardness", "Silica", "Iron", "Turbidity", "Phosphate", "DO"],
    "BWRO": ["pH", "Conductivity", "Total Hardness", "Silica", "Iron", "Chloride", "Turbidity", "Free Chlorine"],
    "Cooling Tower": ["pH", "Conductivity", "Calcium Hardness", "Iron", "Silica", "Orthophosphate", "Free Chlorine", "Chloride", "Turbidity"],
    "Demin Water": ["pH", "Conductivity", "Total Hardness", "Silica", "Iron", "Turbidity", "Free Chlorine"],
    "Lay Up Boiler": ["pH", "Conductivity", "Silica", "Phosphate", "P-Alkalinity", "M-Alkalinity", "O-Alkalinity", "Sulphite"],
    "Lay Up CT": ["pH", "Conductivity", "Iron", "Orthophosphate", "Free Chlorine", "Turbidity"],
    "Raw Water": ["pH", "Conductivity", "Total Hardness", "Silica", "Iron", "Chloride", "Turbidity"],
    "RO Water": ["pH", "Conductivity", "Total Hardness", "Silica", "Iron", "Chloride", "Turbidity"],
    "SWRO": ["pH", "Conductivity", "Total Hardness", "Silica", "Iron", "Chloride", "Turbidity", "Free Chlorine"],
}

RECOMMENDATION_BY_PLANT = {
    "Boiler & BFW": "Tingkatkan kontrol blowdown, cek pengendapan kerak dan pH feedwater, serta review dosing phosphate/oxygen scavenger untuk menjaga kualitas boiler dan BFW.",
    "BWRO": "Lakukan backwash dan review media filter, kontrol SDI/klorinasi, serta periksa pressure drop pada membran agar kinerja RO tetap stabil.",
    "Cooling Tower": "Pantau konsentrasi blowdown, tingkatkan kontrol dosis inhibitor dan klorin, serta cek kebersihan tower dan drift eliminator untuk menurunkan scaling dan korosi.",
    "Demin Water": "Pantau regenerasi resin, kebersihan sand filter, serta kontrol pH dan silica agar kualitas demin tetap sesuai spesifikasi sebelum masuk ke sistem.",
    "Lay Up Boiler": "Pastikan lay-up chemistry terjaga saat idle dengan pH, phosphate, dan sulphite yang sesuai agar tidak terjadi korosi internal boiler.",
    "Lay Up CT": "Periksa kondisi inhibitor dan free chlorine agar pH, iron, serta kualitas air tower tetap aman saat sistem lay-up.",
    "Raw Water": "Perbaiki koagulan, kontrol turbidity dan hardness, serta evaluasi pretreatment seperti sand filter agar feed lebih konsisten.",
    "RO Water": "Review antiscalant dan flushing sistem RO, termasuk pH, silica, hardness, serta kondensor/pretreatment untuk menjaga permeate stabil.",
    "SWRO": "Evaluasi pretreatment, recovery rate, dan dosing antiscalant serta lakukan cleaning membrane bila silika/hardness mulai naik.",
}


def normalize_plant_name(path: Path) -> str:
    name = path.stem.replace("Report Compliance Water Analysis Utility Master - Copy - Copy-", "")
    name = name.replace("-", " ").strip()
    return PLANT_MAP.get(name.replace(" & ", "&").replace(" &", "&").replace("& ", "&"), name)


def parse_date(value: str):
    if not value:
        return None
    s = value.strip()
    for pattern in [r"^(\d{1,2})/(\d{1,2})/(\d{2,4})$", r"^(\d{1,2})-(\w{3})-(\d{2,4})$"]:
        m = re.match(pattern, s)
        if m:
            if "/" in s:
                month, day, year = m.groups()
                dt = datetime.strptime(f"{month}/{day}/{year}", "%m/%d/%Y")
            else:
                day, month_name, year = m.groups()
                dt = datetime.strptime(f"{day}-{month_name}-{year}", "%d-%b-%y")
            return dt.strftime("%Y-%m-%d")
    return None


def parse_float(value):
    if value is None:
        return None
    s = str(value).strip()
    if not s or s.lower() in {"no data", "stop", "n/a", "nan"}:
        return None
    s = s.replace("%", "").replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def percent_from_cell(value):
    v = parse_float(value)
    if v is None:
        return None
    return v if v <= 100 else v / 100 if 0 < v <= 1 else v


def extract_daily_kpi_rows(path: Path):
    rows = []
    text = path.read_text(encoding="utf-8-sig", errors="ignore")
    raw = [line for line in text.splitlines()]

    for idx, line in enumerate(raw):
        cells = [c.strip() for c in line.split(",")]
        if len(cells) < 2:
            continue
        date_value = cells[0].strip()
        parsed_date = parse_date(date_value)
        if not parsed_date:
            continue

        vals = []
        for cell in cells[1:]:
            p = percent_from_cell(cell)
            if p is not None:
                vals.append(p)

        if vals:
            avg = round(sum(vals) / len(vals), 2)
            rows.append({
                "Plant": normalize_plant_name(path),
                "Date": parsed_date,
                "Month": parsed_date[:7],
                "Year": parsed_date[:4],
                "Daily KPI": avg,
                "Source": path.name,
            })
    return rows


def build_summary_dataset():
    rows = []
    for path in CSV_PATHS:
        try:
            rows.extend(extract_daily_kpi_rows(path))
        except Exception:
            continue
    return rows


def style_header(ws, cell_range):
    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(color="FFFFFF", bold=True)
    for row in ws[cell_range]:
        for cell in row:
            cell.fill = fill
            cell.font = font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(
                left=Side(style="thin", color="D9D9D9"),
                right=Side(style="thin", color="D9D9D9"),
                top=Side(style="thin", color="D9D9D9"),
                bottom=Side(style="thin", color="D9D9D9"),
            )


def add_card(ws, title, cell, value, color):
    ws[cell] = title
    ws[cell].font = Font(bold=True, color="FFFFFF")
    ws[cell].fill = PatternFill("solid", fgColor=color)
    ws[cell].alignment = Alignment(horizontal="center", vertical="center")
    ws[cell].border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    target = ws.cell(row=ws[cell].row + 1, column=ws[cell].column)
    target.value = value
    target.font = Font(bold=True, size=16, color="1F1F1F")
    target.alignment = Alignment(horizontal="center", vertical="center")
    target.fill = PatternFill("solid", fgColor="F2F2F2")
    target.border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )


def make_dashboard(data_rows):
    wb = Workbook()
    dash = wb.active
    dash.title = "Dashboard"
    dash.sheet_view.showGridLines = False

    # filter cells
    dash["A1"] = "Utility KPI Dashboard"
    dash["A1"].font = Font(size=18, bold=True, color="1F4E78")
    dash["A3"] = "SELECT DATE"
    dash["B3"] = "2025-12-10"
    dash["A4"] = "SELECT MONTH"
    dash["B4"] = "2025-12"
    dash["A5"] = "SELECT PLANT"
    dash["B5"] = "All"

    dates = sorted({r["Date"] for r in data_rows})
    months = sorted({r["Month"] for r in data_rows})
    plants = sorted({r["Plant"] for r in data_rows})

    dv_date = DataValidation(type="list", formula1=f'"{",".join(dates)}"', allow_blank=False)
    dv_month = DataValidation(type="list", formula1=f'"{",".join(months)}"', allow_blank=False)
    dv_plant = DataValidation(type="list", formula1=f'"All,{",".join(plants)}"', allow_blank=False)
    dash.add_data_validation(dv_date)
    dash.add_data_validation(dv_month)
    dash.add_data_validation(dv_plant)
    dv_date.add("B3")
    dv_month.add("B4")
    dv_plant.add("B5")

    # hidden data sheet
    data_sheet = wb.create_sheet("Data")
    data_sheet.append(["Plant", "Date", "Month", "Year", "Daily KPI", "Monthly KPI", "Yearly KPI"])
    for r in data_rows:
        data_sheet.append([r["Plant"], r["Date"], r["Month"], r["Year"], r["Daily KPI"], "", ""])

    # calculate monthly and yearly values on data sheet
    for plant in plants:
        plant_dates = [r["Date"] for r in data_rows if r["Plant"] == plant]
        for month in months:
            vals = [r["Daily KPI"] for r in data_rows if r["Plant"] == plant and r["Month"] == month]
            if vals:
                data_sheet.append([plant, "", month, month[:4], "", round(sum(vals) / len(vals), 2), ""])
        for year in sorted({r["Year"] for r in data_rows if r["Plant"] == plant}):
            vals = [r["Daily KPI"] for r in data_rows if r["Plant"] == plant and r["Year"] == year]
            if vals:
                data_sheet.append([plant, "", "", year, "", "", round(sum(vals) / len(vals), 2)])

    # formulas for dashboard KPIs
    dash["D3"] = '=IFERROR(AVERAGEIFS(Data!$E$2:$E$1000,Data!$B$2:$B$1000,$B$3,Data!$A$2:$A$1000,IF($B$5=""All"",Data!$A$2:$A$1000,$B$5)),0)'
    dash["D4"] = '=IFERROR(AVERAGEIFS(Data!$F$2:$F$1000,Data!$C$2:$C$1000,$B$4,Data!$A$2:$A$1000,IF($B$5=""All"",Data!$A$2:$A$1000,$B$5)),0)'
    dash["D5"] = '=IFERROR(AVERAGEIFS(Data!$G$2:$G$1000,Data!$D$2:$D$1000,YEAR(TODAY()),Data!$A$2:$A$1000,IF($B$5=""All"",Data!$A$2:$A$1000,$B$5)),0)'

    dash["C3"] = "Daily compliance"
    dash["C4"] = "Monthly compliance"
    dash["C5"] = "Yearly compliance"
    for cell in ["C3", "C4", "C5"]:
        dash[cell].font = Font(bold=True, color="1F1F1F")

    # summary cards on dashboard
    headers = ["A8", "D8", "G8", "J8"]
    colors = ["5B9BD5", "70AD47", "ED7D31", "A5A5A5"]
    for title_cell, value_cell, color in zip(headers, ["A9", "D9", "G9", "J9"], colors):
        dash[title_cell] = "Plant"
        dash[title_cell].font = Font(bold=True)
        dash[value_cell] = "0%"
        dash[value_cell].font = Font(size=14, bold=True)
        dash[value_cell].fill = PatternFill("solid", fgColor=color)

    # create plant table for selected date month
    start_row = 12
    dash[f"A{start_row}"] = "Plant"
    dash[f"B{start_row}"] = "Daily KPI"
    dash[f"C{start_row}"] = "Monthly KPI"
    dash[f"D{start_row}"] = "Yearly KPI"
    style_header(dash, f"A{start_row}:D{start_row}")

    row_idx = start_row + 1
    for plant in plants:
        dash[f"A{row_idx}"] = plant
        dash[f"B{row_idx}"] = f'=IFERROR(AVERAGEIFS(Data!$E$2:$E$1000,Data!$A$2:$A$1000,"{plant}",Data!$B$2:$B$1000,$B$3),0)'
        dash[f"C{row_idx}"] = f'=IFERROR(AVERAGEIFS(Data!$F$2:$F$1000,Data!$A$2:$A$1000,"{plant}",Data!$C$2:$C$1000,$B$4),0)'
        dash[f"D{row_idx}"] = f'=IFERROR(AVERAGEIFS(Data!$G$2:$G$1000,Data!$A$2:$A$1000,"{plant}",Data!$D$2:$D$1000,YEAR(TODAY())),0)'
        row_idx += 1

    # chart daily KPI
    chart = BarChart()
    chart.type = "bar"
    chart.style = 10
    chart.title = "Daily KPI per Plant"
    chart.height = 7
    chart.width = 11
    data = Reference(dash, min_col=2, min_row=start_row, max_col=2, max_row=row_idx - 1)
    cats = Reference(dash, min_col=1, min_row=start_row + 1, max_row=row_idx - 1)
    chart.add_data(data, titles_from_data=False)
    chart.set_categories(cats)
    chart.y_axis.title = "Compliance %"
    chart.x_axis.title = "Plant"
    chart.series[0].graphicalProperties.solidFill = "5B9BD5"
    chart.series[0].dataLabels = DataLabelList()
    chart.series[0].dataLabels.showVal = True
    dash.add_chart(chart, "F12")

    # monthly chart
    monthly = BarChart()
    monthly.type = "bar"
    monthly.style = 10
    monthly.title = "Monthly KPI per Plant"
    monthly.height = 7
    monthly.width = 11
    m_data = Reference(dash, min_col=3, min_row=start_row, max_col=3, max_row=row_idx - 1)
    m_cats = Reference(dash, min_col=1, min_row=start_row + 1, max_row=row_idx - 1)
    monthly.add_data(m_data, titles_from_data=False)
    monthly.set_categories(m_cats)
    monthly.y_axis.title = "Compliance %"
    monthly.x_axis.title = "Plant"
    monthly.series[0].graphicalProperties.solidFill = "70AD47"
    monthly.series[0].dataLabels = DataLabelList()
    monthly.series[0].dataLabels.showVal = True
    dash.add_chart(monthly, "R12")

    # yearly chart
    yearly = BarChart()
    yearly.type = "bar"
    yearly.style = 10
    yearly.title = "Yearly KPI per Plant"
    yearly.height = 7
    yearly.width = 11
    y_data = Reference(dash, min_col=4, min_row=start_row, max_col=4, max_row=row_idx - 1)
    y_cats = Reference(dash, min_col=1, min_row=start_row + 1, max_row=row_idx - 1)
    yearly.add_data(y_data, titles_from_data=False)
    yearly.set_categories(y_cats)
    yearly.y_axis.title = "Compliance %"
    yearly.x_axis.title = "Plant"
    yearly.series[0].graphicalProperties.solidFill = "ED7D31"
    yearly.series[0].dataLabels = DataLabelList()
    yearly.series[0].dataLabels.showVal = True
    dash.add_chart(yearly, "F32")

    # add KPI summary sheet for each plant
    for plant in plants:
        ws = wb.create_sheet(plant)
        ws.sheet_view.showGridLines = False
        ws["A1"] = f"{plant} - Water Quality KPI"
        ws["A1"].font = Font(size=18, bold=True, color="1F4E78")

        ws["A3"] = "Daily Compliance"
        ws["B3"] = f'=IFERROR(AVERAGEIFS(Data!$E$2:$E$1000,Data!$A$2:$A$1000,"{plant}",Data!$B$2:$B$1000,$B$3),0)'
        ws["A4"] = "Monthly Compliance"
        ws["B4"] = f'=IFERROR(AVERAGEIFS(Data!$F$2:$F$1000,Data!$A$2:$A$1000,"{plant}",Data!$C$2:$C$1000,$B$4),0)'
        ws["A5"] = "Yearly Compliance"
        ws["B5"] = f'=IFERROR(AVERAGEIFS(Data!$G$2:$G$1000,Data!$A$2:$A$1000,"{plant}",Data!$D$2:$D$1000,YEAR(TODAY())),0)'

        for r in [3,4,5]:
            ws[f"A{r}"].font = Font(bold=True)
            ws[f"B{r}"].font = Font(bold=True, size=12)
            ws[f"B{r}"].number_format = '0.00%'

        ws["D3"] = "Out of Limit Parameters"
        ws["D3"].font = Font(bold=True)
        out_params = OUT_OF_LIMIT_BY_PLANT.get(plant, ["pH", "Conductivity", "Silica", "Turbidity"])
        for idx, param in enumerate(out_params, start=5):
            ws[f"D{idx}"] = param

        ws["G3"] = "Problem"
        ws["G3"].font = Font(bold=True)
        ws["G5"] = "Plant issue summary"
        ws["G6"] = RECOMMENDATION_BY_PLANT.get(plant, "Lakukan review parameter utama dan lakukan preventive maintenance secara rutin.")
        ws["G6"].alignment = Alignment(wrap_text=True)

        ws["K3"] = "Recommendation"
        ws["K3"].font = Font(bold=True)
        ws["K5"] = RECOMMENDATION_BY_PLANT.get(plant, "Lakukan supervisi kualitas air dan maintenance berkala.")
        ws["K5"].alignment = Alignment(wrap_text=True)

        # Format
        for cell in ["A1", "D3", "G3", "K3"]:
            ws[cell].fill = PatternFill("solid", fgColor="D9EAF7")

        ws["A7"] = "Date"
        ws["B7"] = "KPI"
        ws["C7"] = "Status"
        style_header(ws, "A7:C7")
        row = 8
        for record in sorted([r for r in data_rows if r["Plant"] == plant], key=lambda x: x["Date"]):
            ws[f"A{row}"] = record["Date"]
            ws[f"B{row}"] = record["Daily KPI"] / 100
            ws[f"C{row}"] = "OK" if record["Daily KPI"] >= 90 else "WATCH" if record["Daily KPI"] >= 75 else "ACTION"
            row += 1

        ws.column_dimensions["A"].width = 18
        ws.column_dimensions["B"].width = 14
        ws.column_dimensions["C"].width = 14
        ws.column_dimensions["D"].width = 24
        ws.column_dimensions["G"].width = 45
        ws.column_dimensions["K"].width = 45

    wb.save(ROOT / "Utility_Compliance_Dashboard.xlsx")
    return wb


if __name__ == "__main__":
    dataset = build_summary_dataset()
    if not dataset:
        raise RuntimeError("Tidak ada data KPI yang bisa diproses. Periksa nama file CSV di folder repository.")
    make_dashboard(dataset)
    print(f"Workbook generated: {ROOT / 'Utility_Compliance_Dashboard.xlsx'}")
    print(f"Total KPI rows parsed: {len(dataset)}")
    print("Finished.")
