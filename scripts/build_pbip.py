# -*- coding: utf-8 -*-
"""Generate the Colorado Motor Vehicle Sales Power BI project (PBIP / PBIR + TMDL) — 7 pages with sidebar navigation."""
import json, os, shutil
from PIL import Image, ImageDraw, ImageFont

OUT = r"C:\Users\Asus\Documents\Colorado_PowerBI_Dashboard"
NAME = "Colorado_MV_Sales"
CSV_SRC = r"C:\Users\Asus\Downloads\Analysis-Projects\colorado_motor_vehicle_sales.csv"

# ---------- palette (validated with dataviz validator, dark surface #121C30) ----------
SERIES = ["#E85555", "#3987E5", "#C98500", "#199E70", "#9085E9"]
PAGE_BG = "#0B1220"
PANEL = "#121C30"
PANEL_ALT = "#172440"
PANEL_HEAD = "#1B2A48"
PANEL_BORDER = "#1F2B45"
SIDEBAR = "#0E1729"
INK = "#FFFFFF"
INK2 = "#C3C2B7"
MUTED = "#898781"
GRID = "#26324A"
GOOD = "#0CA30C"
BAD = "#D03B3B"
ACCENT = SERIES[0]
GOLD = SERIES[2]
BN = 1000000000

REP = os.path.join(OUT, f"{NAME}.Report")
SM = os.path.join(OUT, f"{NAME}.SemanticModel")
DEF = os.path.join(REP, "definition")
PAGES_DIR = os.path.join(DEF, "pages")
RES = os.path.join(REP, "StaticResources", "RegisteredResources")
DATA = os.path.join(OUT, "data")

shutil.rmtree(PAGES_DIR, ignore_errors=True)          # drop stale pages/visuals
for d in [OUT, REP, SM, DEF, PAGES_DIR, RES, DATA]:
    os.makedirs(d, exist_ok=True)
for f in os.listdir(RES):
    if f.endswith(".png"):
        os.remove(os.path.join(RES, f))
shutil.copy(CSV_SRC, os.path.join(DATA, "colorado_motor_vehicle_sales.csv"))


def dump(path, obj):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, indent=2, ensure_ascii=False)


# =====================================================================================
# 1. SEMANTIC MODEL (model.bim, TMSL)
# =====================================================================================
sales_m = [
    "let",
    '    Source = Csv.Document(File.Contents(DataFolder & "\\colorado_motor_vehicle_sales.csv"), [Delimiter = ",", Columns = 4, Encoding = 65001, QuoteStyle = QuoteStyle.None]),',
    "    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),",
    '    Typed = Table.TransformColumnTypes(Headers, {{"year", Int64.Type}, {"quarter", Int64.Type}, {"county", type text}, {"sales", Int64.Type}}),',
    '    Renamed = Table.RenameColumns(Typed, {{"year", "Year"}, {"quarter", "Quarter No"}, {"county", "County"}, {"sales", "Sales"}}),',
    '    AddDate = Table.AddColumn(Renamed, "Date", each #date([Year], ([Quarter No] - 1) * 3 + 1, 1), type date),',
    '    AddQ = Table.AddColumn(AddDate, "Quarter", each "Q" & Text.From([Quarter No]), type text),',
    '    AddPeriod = Table.AddColumn(AddQ, "Period", each Text.From([Year]) & " Q" & Text.From([Quarter No]), type text),',
    '    AddPeriodSort = Table.AddColumn(AddPeriod, "Period Sort", each [Year] * 10 + [Quarter No], Int64.Type),',
    '    AddRegion = Table.AddColumn(AddPeriodSort, "Region", each',
    '        if List.Contains({"Adams", "Arapahoe", "Denver", "Douglas", "Jefferson", "Boulder", "Broomfield", "Boulder/Broomfield"}, [County]) then "Denver Metro"',
    '        else if List.Contains({"Larimer", "Weld"}, [County]) then "N. Front Range"',
    '        else if List.Contains({"El Paso", "Pueblo", "Fremont"}, [County]) then "S. Front Range"',
    '        else if List.Contains({"Mesa", "Garfield", "La Plata"}, [County]) then "Western Slope"',
    '        else "Rest of State", type text)',
    "in",
    "    AddRegion",
]

calendar_m = [
    "let",
    "    Dates = List.Dates(#date(2008, 1, 1), Duration.Days(#date(2015, 12, 31) - #date(2008, 1, 1)) + 1, #duration(1, 0, 0, 0)),",
    '    T = Table.FromList(Dates, Splitter.SplitByNothing(), {"Date"}),',
    '    T1 = Table.TransformColumnTypes(T, {{"Date", type date}}),',
    '    T2 = Table.AddColumn(T1, "Year", each Date.Year([Date]), Int64.Type),',
    '    T3 = Table.AddColumn(T2, "Quarter No", each Date.QuarterOfYear([Date]), Int64.Type),',
    '    T4 = Table.AddColumn(T3, "Quarter", each "Q" & Text.From(Date.QuarterOfYear([Date])), type text),',
    '    T5 = Table.AddColumn(T4, "Year Quarter", each Text.From(Date.Year([Date])) & " Q" & Text.From(Date.QuarterOfYear([Date])), type text),',
    '    T6 = Table.AddColumn(T5, "Year Quarter Sort", each Date.Year([Date]) * 10 + Date.QuarterOfYear([Date]), Int64.Type),',
    '    T7 = Table.AddColumn(T6, "Fiscal Year", each "FY" & Text.From(Date.Year([Date])), type text),',
    '    T8 = Table.AddColumn(T7, "Quarter Start", each Date.StartOfQuarter([Date]), type date)',
    "in",
    "    T8",
]


def col(name, dtype, fmt=None, summarize="none", hidden=False, sort_by=None, key=False):
    c = {"name": name, "dataType": dtype, "sourceColumn": name, "summarizeBy": summarize}
    if fmt:
        c["formatString"] = fmt
    if hidden:
        c["isHidden"] = True
    if sort_by:
        c["sortByColumn"] = sort_by
    if key:
        c["isKey"] = True
    return c


B_FMT = '$#,0.0"B"'
B2_FMT = '$#,0.00"B"'
PCT = "0.0%"

measures = [
    ("Total Sales", "SUM ( Sales[Sales] )", '$#,0'),
    ("Total Sales (B)", f"DIVIDE ( [Total Sales], {BN} )", B2_FMT),
    ("Sales $B", f"DIVIDE ( [Total Sales], {BN} )", B_FMT),
    ("Sales $M", "DIVIDE ( [Total Sales], 1000000 )", '$#,0.0"M"'),
    ("Sales PY", "CALCULATE ( [Total Sales], SAMEPERIODLASTYEAR ( 'Calendar'[Date] ) )", '$#,0'),
    ("Sales PY $B", f"DIVIDE ( [Sales PY], {BN} )", B_FMT),
    ("YoY %", "DIVIDE ( [Total Sales] - [Sales PY], [Sales PY] )", PCT),
    ("Sales PQ", "CALCULATE ( [Total Sales], DATEADD ( 'Calendar'[Date], -1, QUARTER ) )", '$#,0'),
    ("QoQ %", "DIVIDE ( [Total Sales] - [Sales PQ], [Sales PQ] )", PCT),
    ("Rolling 4Q Sales $B", f"VAR w = DATESINPERIOD ( 'Calendar'[Date], MAX ( 'Calendar'[Date] ), -1, YEAR ) VAR q = CALCULATE ( DISTINCTCOUNT ( Sales[Period Sort] ), w ) RETURN IF ( q = 4, CALCULATE ( [Sales $B], w ) )", B_FMT),
    ("Quarters", "DISTINCTCOUNT ( Sales[Period Sort] )", "#,0"),
    ("Rows", "COUNTROWS ( Sales )", "#,0"),
    ("Avg Quarterly Sales", f"DIVIDE ( [Total Sales], [Quarters] * {BN} )", B2_FMT),
    ("Latest Year", "MAX ( Sales[Year] )", "0"),
    ("First Year", "MIN ( Sales[Year] )", "0"),
    ("Latest Year Sales", f"VAR y = [Latest Year] RETURN DIVIDE ( CALCULATE ( [Total Sales], Sales[Year] = y ), {BN} )", B2_FMT),
    ("Prior Year Sales", f"VAR y = [Latest Year] - 1 RETURN DIVIDE ( CALCULATE ( [Total Sales], Sales[Year] = y ), {BN} )", B2_FMT),
    ("Latest YoY %", "DIVIDE ( [Latest Year Sales] - [Prior Year Sales], [Prior Year Sales] )", PCT),
    ("CAGR", "VAR y1 = [First Year] VAR y2 = [Latest Year] VAR s1 = CALCULATE ( [Total Sales], Sales[Year] = y1 ) VAR s2 = CALCULATE ( [Total Sales], Sales[Year] = y2 ) RETURN IF ( y2 > y1 && s1 > 0, DIVIDE ( s2, s1 ) ^ ( 1 / ( y2 - y1 ) ) - 1 )", PCT),
    ("Best Quarter Sales", f"DIVIDE ( MAXX ( VALUES ( Sales[Period] ), [Total Sales] ), {BN} )", B2_FMT),
    ("Best Quarter", "MAXX ( TOPN ( 1, VALUES ( Sales[Period] ), [Total Sales] ), Sales[Period] )", None),
    ("Counties", "DISTINCTCOUNT ( Sales[County] )", "#,0"),
    ("Years", "DISTINCTCOUNT ( Sales[Year] )", "#,0"),
    ("Top County", "MAXX ( TOPN ( 1, VALUES ( Sales[County] ), [Total Sales] ), Sales[County] )", None),
    ("Top County Share", "DIVIDE ( MAXX ( VALUES ( Sales[County] ), [Total Sales] ), [Total Sales] )", PCT),
    ("Share of State %", "DIVIDE ( [Total Sales], CALCULATE ( [Total Sales], REMOVEFILTERS ( Sales[County], Sales[Region] ) ) )", PCT),
    ("Share of Year %", "DIVIDE ( [Total Sales], CALCULATE ( [Total Sales], REMOVEFILTERS ( 'Calendar'[Quarter], 'Calendar'[Quarter No], Sales[Quarter], Sales[Quarter No] ) ) )", PCT),
    # KPI sub-labels (text)
    ("KPI Total Sales Sub", 'IF ( ISBLANK ( [CAGR] ), "single year selected", ( IF ( [CAGR] >= 0, "▲ ", "▼ " ) & FORMAT ( ABS ( [CAGR] ), "0.0%" ) & " CAGR FY" & RIGHT ( [First Year], 2 ) & "–" & RIGHT ( [Latest Year], 2 ) ) )', None),
    ("KPI Latest Year Sub", 'IF ( ISBLANK ( [Latest YoY %] ), "no prior year in selection", ( IF ( [Latest YoY %] >= 0, "▲ ", "▼ " ) & FORMAT ( ABS ( [Latest YoY %] ), "0.0%" ) & " vs FY" & ( [Latest Year] - 1 ) ) )', None),
    ("KPI Avg Quarter Sub", '"across " & [Quarters] & " quarters"', None),
    ("KPI Best Quarter Sub", '"peak quarter: " & [Best Quarter]', None),
    ("KPI Top County Sub", 'FORMAT ( [Top County Share], "0.0%" ) & " share of sales"', None),
    ("KPI Counties Sub", '[Years] & " years · " & [Quarters] & " quarters"', None),
    # colours for conditional formatting
    ("Delta Color CAGR", f'IF ( [CAGR] >= 0, "{GOOD}", "{BAD}" )', None),
    ("Delta Color YoY", f'IF ( [Latest YoY %] >= 0, "{GOOD}", "{BAD}" )', None),
    ("YoY Bar Color", f'IF ( [YoY %] >= 0, "{GOOD}", "{BAD}" )', None),
    ("QoQ Bar Color", f'IF ( [QoQ %] >= 0, "{GOOD}", "{BAD}" )', None),
    ("CAGR Bar Color", f'IF ( [CAGR] >= 0, "{GOOD}", "{BAD}" )', None),
]

sales_table = {
    "name": "Sales",
    "columns": [
        col("Year", "int64", "0"),
        col("Quarter No", "int64", "0", hidden=True),
        col("County", "string"),
        col("Sales", "int64", "$#,0", summarize="sum"),
        col("Date", "dateTime", "yyyy-mm-dd"),
        col("Quarter", "string"),
        col("Period", "string", sort_by="Period Sort"),
        col("Period Sort", "int64", "0", hidden=True),
        col("Region", "string"),
    ],
    "partitions": [{"name": "Sales", "mode": "import", "source": {"type": "m", "expression": sales_m}}],
    "measures": [
        {k: v for k, v in {"name": n, "expression": e, "formatString": f}.items() if v is not None}
        for n, e, f in measures
    ],
}

calendar_table = {
    "name": "Calendar",
    "dataCategory": "Time",
    "columns": [
        col("Date", "dateTime", "yyyy-mm-dd", key=True),
        col("Year", "int64", "0"),
        col("Quarter No", "int64", "0", hidden=True),
        col("Quarter", "string"),
        col("Year Quarter", "string", sort_by="Year Quarter Sort"),
        col("Year Quarter Sort", "int64", "0", hidden=True),
        col("Fiscal Year", "string"),
        col("Quarter Start", "dateTime", "mmm yyyy"),
    ],
    "partitions": [{"name": "Calendar", "mode": "import", "source": {"type": "m", "expression": calendar_m}}],
}

model = {
    "name": NAME,
    "compatibilityLevel": 1606,
    "model": {
        "culture": "en-US",
        "sourceQueryCulture": "en-US",
        "defaultPowerBIDataSourceVersion": "powerBI_V3",
        "dataAccessOptions": {"legacyRedirects": True, "returnErrorValuesAsNull": True},
        "expressions": [{
            "name": "DataFolder", "kind": "m",
            "expression": [f'"{DATA}" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]'],
            "annotations": [{"name": "PBI_ResultType", "value": "Text"}],
        }],
        "tables": [sales_table, calendar_table],
        "relationships": [{"name": "SalesToCalendar", "fromTable": "Sales", "fromColumn": "Date", "toTable": "Calendar", "toColumn": "Date"}],
        "annotations": [{"name": "PBI_QueryOrder", "value": '["DataFolder","Sales","Calendar"]'}],
    },
}
# ---- emit the model as TMDL (one file per table) instead of a single model.bim ----
import uuid
NS = uuid.UUID("6f1d3c2e-7b44-4a6e-9c1a-colorado0001".replace("colorado0001", "0000c0104ad0"))


def tag(*parts):
    """stable lineageTag so regenerating the project does not churn git diffs"""
    return str(uuid.uuid5(NS, "|".join(parts)))


def q(name):
    """TMDL identifier quoting: quote anything that is not a plain word"""
    import re
    return name if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", name) else "'" + name.replace("'", "''") + "'"


def tmdl_table(t):
    L = [f"table {q(t['name'])}", f"	lineageTag: {tag('table', t['name'])}"]
    if t.get("dataCategory"):
        L.append(f"	dataCategory: {t['dataCategory']}")
    L.append("")
    for m in t.get("measures", []):
        L += [f"	measure {q(m['name'])} = ```", f"			{m['expression']}", "			```"]
        if m.get("formatString"):
            L.append(f"		formatString: {m['formatString']}")
        L += [f"		lineageTag: {tag('measure', t['name'], m['name'])}", ""]
    for c in t["columns"]:
        L.append(f"	column {q(c['name'])}")
        L.append(f"		dataType: {c['dataType']}")
        if c.get("isKey"):
            L.append("		isKey")
        if c.get("isHidden"):
            L.append("		isHidden")
        if c.get("formatString"):
            L.append(f"		formatString: {c['formatString']}")
        L.append(f"		lineageTag: {tag('column', t['name'], c['name'])}")
        L.append(f"		summarizeBy: {c.get('summarizeBy', 'none')}")
        L.append(f"		sourceColumn: {c['sourceColumn']}")
        if c.get("sortByColumn"):
            L.append(f"		sortByColumn: {q(c['sortByColumn'])}")
        L += ["", "		annotation SummarizationSetBy = User", ""]
    for part in t["partitions"]:
        L += [f"	partition {q(part['name'])} = m", f"		mode: {part['mode']}", "		source ="]
        L += ["				" + line for line in part["source"]["expression"]]
        L.append("")
    L += ["	annotation PBI_ResultType = Table", ""]
    return "\n".join(L)


def emit_tmdl(model):
    for stale in ["model.bim"]:
        sp = os.path.join(SM, stale)
        if os.path.exists(sp):
            os.remove(sp)
    ddir = os.path.join(SM, "definition")
    shutil.rmtree(ddir, ignore_errors=True)
    os.makedirs(os.path.join(ddir, "tables"), exist_ok=True)
    m = model["model"]
    w = lambda name, text: open(os.path.join(ddir, name), "w", encoding="utf-8", newline="\n").write(text)
    w("database.tmdl", f"database\n\tcompatibilityLevel: {model['compatibilityLevel']}\n")
    lines = ["model Model", f"	culture: {m['culture']}", f"	defaultPowerBIDataSourceVersion: {m['defaultPowerBIDataSourceVersion']}",
             f"	sourceQueryCulture: {m['sourceQueryCulture']}", "	dataAccessOptions", "		legacyRedirects", "		returnErrorValuesAsNull", ""]
    for a in m.get("annotations", []):
        lines += [f"annotation {a['name']} = {a['value']}", ""]
    lines += ["annotation __PBI_TimeIntelligenceEnabled = 0", ""]
    lines += [f"ref table {q(t['name'])}" for t in m["tables"]] + [""]
    w("model.tmdl", "\n".join(lines))
    ex = []
    for e in m.get("expressions", []):
        ex += [f"expression {q(e['name'])} = " + " ".join(e["expression"]), f"	lineageTag: {tag('expression', e['name'])}", ""]
        for a in e.get("annotations", []):
            ex += [f"	annotation {a['name']} = {a['value']}", ""]
    w("expressions.tmdl", "\n".join(ex))
    rel = []
    for r in m["relationships"]:
        rel += [f"relationship {r['name']}", f"	fromColumn: {q(r['fromTable'])}.{q(r['fromColumn'])}", f"	toColumn: {q(r['toTable'])}.{q(r['toColumn'])}", ""]
    w("relationships.tmdl", "\n".join(rel))
    for t in m["tables"]:
        w(os.path.join("tables", f"{t['name']}.tmdl"), tmdl_table(t))


emit_tmdl(model)
dump(os.path.join(SM, "definition.pbism"), {
    "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
    "version": "4.2", "settings": {}})

# =====================================================================================
# 2. PROJECT + REPORT SHELL + THEME
# =====================================================================================
dump(os.path.join(OUT, f"{NAME}.pbip"), {"version": "1.0", "artifacts": [{"report": {"path": f"{NAME}.Report"}}], "settings": {"enableAutoRecovery": True}})
dump(os.path.join(REP, "definition.pbir"), {
    "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/1.0.0/schema.json",
    "version": "4.0", "datasetReference": {"byPath": {"path": f"../{NAME}.SemanticModel"}}})
dump(os.path.join(DEF, "version.json"), {"$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/versionMetadata/1.0.0/schema.json", "version": "2.0.0"})

THEME_FILE = "ColoradoDark.json"
table_style = {
    "grid": [{"gridVertical": False, "gridHorizontal": True, "gridHorizontalColor": {"solid": {"color": GRID}}, "outlineColor": {"solid": {"color": GRID}}, "textSize": 9, "rowPadding": 3}],
    "columnHeaders": [{"fontColor": {"solid": {"color": INK}}, "backColor": {"solid": {"color": PANEL_HEAD}}, "fontSize": 9, "bold": True, "outline": "None"}],
    "rowHeaders": [{"fontColor": {"solid": {"color": INK2}}, "backColor": {"solid": {"color": PANEL}}, "fontSize": 9, "outline": "None"}],
    "values": [{"fontColorPrimary": {"solid": {"color": INK2}}, "backColorPrimary": {"solid": {"color": PANEL}}, "fontColorSecondary": {"solid": {"color": INK2}}, "backColorSecondary": {"solid": {"color": PANEL_ALT}}, "fontSize": 9, "outline": "None"}],
    "total": [{"fontColor": {"solid": {"color": INK}}, "backColor": {"solid": {"color": PANEL_HEAD}}, "fontSize": 9, "bold": True, "outline": "None"}],
    "subTotals": [{"fontColor": {"solid": {"color": INK}}, "backColor": {"solid": {"color": PANEL_HEAD}}, "fontSize": 9, "outline": "None"}],
    "columnTotal": [{"fontColor": {"solid": {"color": INK}}, "backColor": {"solid": {"color": PANEL_HEAD}}, "fontSize": 9, "bold": True}],
    "rowTotal": [{"fontColor": {"solid": {"color": INK}}, "backColor": {"solid": {"color": PANEL_HEAD}}, "fontSize": 9, "bold": True}],
}
theme = {
    "name": "ColoradoDark", "dataColors": SERIES, "background": PANEL, "foreground": INK, "tableAccent": ACCENT,
    "good": GOOD, "neutral": GOLD, "bad": BAD,
    "textClasses": {
        "label": {"color": INK2, "fontFace": "Segoe UI", "fontSize": 9},
        "callout": {"color": INK, "fontFace": "Segoe UI Semibold", "fontSize": 20},
        "title": {"color": INK, "fontFace": "Segoe UI Semibold", "fontSize": 11},
        "header": {"color": INK, "fontFace": "Segoe UI Semibold", "fontSize": 11},
    },
    "visualStyles": {
        "*": {"*": {
            "background": [{"show": False}], "border": [{"show": False}], "visualHeader": [{"show": False}],
            "title": [{"show": True, "fontColor": {"solid": {"color": INK}}, "fontSize": 11, "fontFamily": "Segoe UI Semibold", "alignment": "left"}],
            "subTitle": [{"show": False}],
            "categoryAxis": [{"labelColor": {"solid": {"color": INK2}}, "gridlineShow": False, "showAxisTitle": False, "fontSize": 9}],
            "valueAxis": [{"labelColor": {"solid": {"color": INK2}}, "gridlineColor": {"solid": {"color": GRID}}, "gridlineShow": True, "showAxisTitle": False, "fontSize": 9}],
            "legend": [{"labelColor": {"solid": {"color": INK2}}, "fontSize": 9, "showTitle": False}],
            "labels": [{"color": {"solid": {"color": INK}}, "fontSize": 9}],
            "categoryLabels": [{"color": {"solid": {"color": INK2}}}],
            "outspace": [{"color": {"solid": {"color": PAGE_BG}}}],
        }},
        "page": {"*": {
            "background": [{"color": {"solid": {"color": PAGE_BG}}, "transparency": 0}],
            "outspace": [{"color": {"solid": {"color": PAGE_BG}}}],
            # Filters pane + filter cards live under page.* in the theme schema; dark so text is not white-on-white
            "outspacePane": [{"backgroundColor": {"solid": {"color": PANEL}}, "foregroundColor": {"solid": {"color": INK2}},
                              "transparency": 0, "border": True, "borderColor": {"solid": {"color": PANEL_BORDER}},
                              "titleSize": 12, "headerSize": 10, "fontFamily": "Segoe UI", "checkboxAndApplyColor": {"solid": {"color": ACCENT}},
                              "inputBoxColor": {"solid": {"color": PANEL_ALT}}, "width": 220}],
            "filterCard": [
                {"$id": "Available", "backgroundColor": {"solid": {"color": PANEL_ALT}}, "foregroundColor": {"solid": {"color": INK}},
                 "transparency": 0, "border": True, "borderColor": {"solid": {"color": PANEL_BORDER}}, "textSize": 9, "fontFamily": "Segoe UI",
                 "inputBoxColor": {"solid": {"color": PANEL}}},
                {"$id": "Applied", "backgroundColor": {"solid": {"color": "#3A1E28"}}, "foregroundColor": {"solid": {"color": INK}},
                 "transparency": 0, "border": True, "borderColor": {"solid": {"color": ACCENT}}, "textSize": 9, "fontFamily": "Segoe UI",
                 "inputBoxColor": {"solid": {"color": PANEL}}},
            ]
        }},
        "slicer": {"*": {
            "general": [{"outlineColor": {"solid": {"color": PANEL_BORDER}}, "outlineWeight": 1}],
            "items": [{"fontColor": {"solid": {"color": INK}}, "background": {"solid": {"color": PANEL}}, "fontSize": 9}],
            "header": [{"show": False}],
        }},
        "tableEx": {"*": table_style},
        "pivotTable": {"*": table_style},
    },
}
dump(os.path.join(RES, THEME_FILE), theme)

# =====================================================================================
# 3. JSON HELPERS
# =====================================================================================
def lit(v): return {"expr": {"Literal": {"Value": v}}}
def s(v): return lit("'" + str(v).replace("'", "''") + "'")
def d(v): return lit(f"{v}D")
def b(v): return lit("true" if v else "false")
def color(hexv): return {"solid": {"color": s(hexv)}}
def color_by_measure(entity, prop):
    return {"solid": {"color": {"expr": {"Measure": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": prop}}}}}


def fmeasure(entity, prop, display=None):
    p = {"field": {"Measure": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": prop}}, "queryRef": f"{entity}.{prop}", "nativeQueryRef": prop}
    if display:
        p["displayName"] = display
    return p


def fcolumn(entity, prop, active=None, display=None):
    p = {"field": {"Column": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": prop}}, "queryRef": f"{entity}.{prop}", "nativeQueryRef": prop}
    if active is not None:
        p["active"] = active
    if display:
        p["displayName"] = display
    return p


def sort_by(field_proj, direction):
    return {"sort": [{"field": field_proj["field"], "direction": direction}], "isDefaultSort": True}


def title(text, size=11):
    return {"title": [{"properties": {"show": b(True), "text": s(text), "fontColor": color(INK), "fontSize": d(size), "fontFamily": s("Segoe UI Semibold"), "alignment": s("left")}}],
            "subTitle": [{"properties": {"show": b(False)}}],
            "background": [{"properties": {"show": b(False)}}], "border": [{"properties": {"show": b(False)}}]}


def no_title():
    return {"title": [{"properties": {"show": b(False)}}], "subTitle": [{"properties": {"show": b(False)}}],
            "background": [{"properties": {"show": b(False)}}], "border": [{"properties": {"show": b(False)}}]}


AX_CAT = {"showAxisTitle": b(False), "labelColor": color(INK2), "gridlineShow": b(False), "fontSize": d(9)}
AX_VAL_HIDDEN = {"show": b(False), "showAxisTitle": b(False), "gridlineShow": b(False), "labelDisplayUnits": d(1)}
AX_VAL = {"show": b(True), "showAxisTitle": b(False), "labelColor": color(INK2), "gridlineColor": color(GRID), "fontSize": d(8), "labelDisplayUnits": d(1)}
LABELS = {"show": b(True), "color": color(INK), "labelDisplayUnits": d(1), "fontSize": d(8)}
LEGEND_TOP = {"show": b(True), "position": s("Top"), "showTitle": b(False), "labelColor": color(INK2), "fontSize": d(8)}


class Page:
    def __init__(self, key, display, nav_label, nav_glyph, title_text, subtitle, slicers=True):
        self.key, self.display, self.nav_label, self.nav_glyph = key, display, nav_label, nav_glyph
        self.title_text, self.subtitle, self.slicers = title_text, subtitle, slicers
        self.visuals, self.panels, self.kpi_boxes, self.zc = [], [], [], 0

    def vis(self, name, vtype, x, y, w, h, query=None, objects=None, vc=None, extra=None):
        self.zc += 1
        v = {"$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/1.0.0/schema.json",
             "name": f"{self.key}_{name}",
             "position": {"x": round(x, 2), "y": round(y, 2), "z": self.zc * 100, "width": round(w, 2), "height": round(h, 2), "tabOrder": self.zc * 100},
             "visual": {"visualType": vtype, "drillFilterOtherVisuals": True}}
        if query: v["visual"]["query"] = query
        if objects: v["visual"]["objects"] = objects
        v["visual"]["visualContainerObjects"] = vc or {}
        if extra: v["visual"].update(extra)
        self.visuals.append(v)
        return v

    def panel(self, x, y, w, h):
        self.panels.append((x, y, w, h))

    # ---- reusable chart builders (each draws its own panel) ----
    def bar(self, name, x, y, w, h, cat, meas, ttl, fill=SERIES[0], fill_measure=None, sort_desc=True, labels=True, axis_font=8):
        self.panel(x, y, w, h)
        dp = color_by_measure("Sales", fill_measure) if fill_measure else color(fill)
        self.vis(name, "clusteredBarChart", x + 6, y + 4, w - 12, h - 8,
                 query={"queryState": {"Category": {"projections": [cat]}, "Y": {"projections": [meas]}},
                        "sortDefinition": sort_by(meas if sort_desc else cat, "Descending" if sort_desc else "Ascending")},
                 objects={"categoryAxis": [{"properties": dict(AX_CAT, fontSize=d(axis_font), maxMarginFactor=d(40))}],
                          "valueAxis": [{"properties": AX_VAL_HIDDEN}],
                          "labels": [{"properties": dict(LABELS, show=b(labels))}],
                          "dataPoint": [{"properties": {"fill": dp}}],
                          "general": [{"properties": {"responsive": b(False)}}]},
                 vc=title(ttl))

    def column(self, name, x, y, w, h, cat, meas, ttl, fill=SERIES[0], fill_measure=None, scalar=False, labels=True, value_axis=False):
        self.panel(x, y, w, h)
        dp = color_by_measure("Sales", fill_measure) if fill_measure else color(fill)
        cax = dict(AX_CAT, axisType=s("Scalar" if scalar else "Categorical"), fontSize=d(8))
        self.vis(name, "clusteredColumnChart", x + 6, y + 4, w - 12, h - 8,
                 query={"queryState": {"Category": {"projections": [cat]}, "Y": {"projections": [meas]}}, "sortDefinition": sort_by(cat, "Ascending")},
                 objects={"categoryAxis": [{"properties": cax}],
                          "valueAxis": [{"properties": AX_VAL if value_axis else AX_VAL_HIDDEN}],
                          "labels": [{"properties": dict(LABELS, show=b(labels))}],
                          "dataPoint": [{"properties": {"fill": dp}}],
                          "general": [{"properties": {"responsive": b(False)}}]},
                 vc=title(ttl))

    def series_chart(self, name, vtype, x, y, w, h, cat, series, meas, ttl, scalar=False, labels=False, colors=None):
        self.panel(x, y, w, h)
        cax = dict(AX_CAT, axisType=s("Scalar" if scalar else "Categorical"), fontSize=d(8))
        objs = {"categoryAxis": [{"properties": cax}], "valueAxis": [{"properties": AX_VAL}],
                "legend": [{"properties": LEGEND_TOP}], "labels": [{"properties": dict(LABELS, show=b(labels))}],
                "general": [{"properties": {"responsive": b(False)}}]}
        if colors:
            objs["dataPoint"] = [{"properties": {"fill": color(c)}, "selector": {"data": [{"scopeId": {"Comparison": {"ComparisonKind": 0, "Left": series["field"], "Right": {"Literal": {"Value": f"'{k}'"}}}}}]}} for k, c in colors]
        if vtype == "lineChart":
            objs["lineStyles"] = [{"properties": {"strokeWidth": d(2), "showMarker": b(False)}}]
        q = {"queryState": {"Category": {"projections": [cat]}, "Series": {"projections": [series]}, "Y": {"projections": [meas]}}, "sortDefinition": sort_by(cat, "Ascending")}
        self.vis(name, vtype, x + 6, y + 4, w - 12, h - 8, query=q, objects=objs, vc=title(ttl))

    def area(self, name, x, y, w, h, cat, meas, ttl, fill=SERIES[1]):
        self.panel(x, y, w, h)
        self.vis(name, "areaChart", x + 6, y + 4, w - 12, h - 8,
                 query={"queryState": {"Category": {"projections": [cat]}, "Y": {"projections": [meas]}}, "sortDefinition": sort_by(cat, "Ascending")},
                 objects={"categoryAxis": [{"properties": dict(AX_CAT, axisType=s("Scalar"), fontSize=d(8))}],
                          "valueAxis": [{"properties": AX_VAL}],
                          "dataPoint": [{"properties": {"fill": color(fill)}}],
                          "lineStyles": [{"properties": {"strokeWidth": d(2), "showMarker": b(False)}}],
                          "general": [{"properties": {"responsive": b(False)}}]},
                 vc=title(ttl))

    def line(self, name, x, y, w, h, cat, measures_, ttl, colors=None):
        self.panel(x, y, w, h)
        objs = {"categoryAxis": [{"properties": dict(AX_CAT, axisType=s("Scalar"), fontSize=d(8))}],
                "valueAxis": [{"properties": AX_VAL}], "legend": [{"properties": LEGEND_TOP}],
                "lineStyles": [{"properties": {"strokeWidth": d(2), "showMarker": b(False)}}],
                "general": [{"properties": {"responsive": b(False)}}]}
        if colors:
            objs["dataPoint"] = [{"properties": {"fill": color(c)}, "selector": {"metadata": m["queryRef"]}} for m, c in zip(measures_, colors)]
        self.vis(name, "lineChart", x + 6, y + 4, w - 12, h - 8,
                 query={"queryState": {"Category": {"projections": [cat]}, "Y": {"projections": measures_}}, "sortDefinition": sort_by(cat, "Ascending")},
                 objects=objs, vc=title(ttl))

    def donut(self, name, x, y, w, h, cat, meas, ttl, legend_pos="Bottom"):
        self.panel(x, y, w, h)
        self.vis(name, "donutChart", x + 6, y + 4, w - 12, h - 8,
                 query={"queryState": {"Category": {"projections": [cat]}, "Y": {"projections": [meas]}}, "sortDefinition": sort_by(cat, "Ascending")},
                 objects={"legend": [{"properties": dict(LEGEND_TOP, position=s(legend_pos))}],
                          "labels": [{"properties": {"show": b(True), "labelStyle": s("Percent of total"), "color": color(INK), "fontSize": d(8)}}],
                          "calloutLabel": [{"properties": {"show": b(False)}}]},
                 vc=title(ttl))

    def table(self, name, x, y, w, h, projections, ttl, sort=None, sort_dir="Ascending", font=9):
        self.panel(x, y, w, h)
        for pr in projections: pr.pop("active", None)
        q = {"queryState": {"Values": {"projections": projections}}}
        if sort is not None:
            q["sortDefinition"] = sort_by(sort, sort_dir)
        self.vis(name, "tableEx", x + 6, y + 4, w - 12, h - 8, query=q,
                 objects={"grid": [{"properties": {"textSize": d(font), "rowPadding": d(3)}}],
                          "columnHeaders": [{"properties": {"fontSize": d(font), "bold": b(True), "wordWrap": b(True)}}],
                          "values": [{"properties": {"fontSize": d(font)}}]},
                 vc=title(ttl) if ttl else no_title())

    def matrix(self, name, x, y, w, h, rows, cols, values, ttl, font=9):
        self.panel(x, y, w, h)
        for pr in [rows, cols] + values: pr.pop("active", None)
        self.vis(name, "pivotTable", x + 6, y + 4, w - 12, h - 8,
                 query={"queryState": {"Rows": {"projections": [rows]}, "Columns": {"projections": [cols]}, "Values": {"projections": values}}},
                 objects={"grid": [{"properties": {"textSize": d(font), "rowPadding": d(3)}}],
                          "columnHeaders": [{"properties": {"fontSize": d(font), "bold": b(True)}}],
                          "rowHeaders": [{"properties": {"fontSize": d(font)}}],
                          "values": [{"properties": {"fontSize": d(font)}}]},
                 vc=title(ttl))

    def textbox(self, name, x, y, w, h, paragraphs):
        """paragraphs: list of (text, style) where style in {'h','b','p'}"""
        styles = {"h": {"fontWeight": "bold", "fontSize": "12pt", "color": INK, "fontFamily": "Segoe UI Semibold"},
                  "b": {"fontWeight": "bold", "fontSize": "9pt", "color": INK, "fontFamily": "Segoe UI Semibold"},
                  "p": {"fontSize": "8pt", "color": INK2, "fontFamily": "Segoe UI"}}
        self.vis(name, "textbox", x, y, w, h,
                 objects={"general": [{"properties": {"paragraphs": [{"textRuns": [{"value": t, "textStyle": styles[st]}]} for t, st in paragraphs]}}]},
                 vc=no_title())

    def card(self, name, x, y, w, h, meas, size=16, col=None, col_measure=None, bold=True, font="Segoe UI Semibold"):
        labels = {"color": (color_by_measure("Sales", col_measure) if col_measure else color(col or INK)), "fontSize": d(size), "fontFamily": s(font), "bold": b(bold), "labelDisplayUnits": d(1)}
        self.vis(name, "card", x, y, w, h, query={"queryState": {"Values": {"projections": [meas]}}},
                 objects={"categoryLabels": [{"properties": {"show": b(False)}}], "labels": [{"properties": labels}], "wordWrap": [{"properties": {"show": b(False)}}]},
                 vc=no_title())


# =====================================================================================
# 4. LAYOUT CONSTANTS
# =====================================================================================
SIDEBAR_W = 150
X0, X1 = 162, 1268
CW = X1 - X0                 # 1106 content width
TOP = 66                     # first content row
BOTTOM = 690
FOOT_Y = 696
NAV_Y0, NAV_STEP, NAV_H = 76, 36, 28
GAP = 12

PAGES = [
    Page("dashboard", "Executive Dashboard", "Dashboard", "\uE80F", "Executive Dashboard", "Sales  |  Growth  |  Counties"),
    Page("trend", "Sales Trend", "Sales Trend", "\uE9D9", "Sales Trend", "Quarterly  |  Annual  |  Rolling 4-quarter"),
    Page("counties", "Counties", "Counties", "\uE707", "Counties", "Ranking  |  County × Year  |  Growth  |  Share"),
    Page("regions", "Regions", "Regions", "\uE804", "Regions", "Denver Metro  |  Front Range  |  Western Slope  |  Rest of State"),
    Page("quarters", "Quarters", "Quarters", "\uE787", "Quarters", "Seasonality  |  Quarter × Year  |  Same-quarter growth"),
    Page("data", "Data Table", "Data", "\uE9F9", "Data Table", "Every county-quarter record  |  filter with the slicers"),
    Page("about", "About", "About", "\uE946", "About this Dashboard", "Source  |  Definitions  |  Method  |  Caveats", slicers=False),
]
PAGE_BY_KEY = {p.key: p for p in PAGES}

# ---- shared chrome: slicers + sidebar navigation buttons ------------------------------
slicer_defs = [("Year", "Calendar", "Fiscal Year"), ("Quarter", "Calendar", "Quarter"), ("County", "Sales", "County")]
SL_W, SL_H, SL_GAP = 150, 30, 12
sl_x0 = X1 - (SL_W * 3 + SL_GAP * 2)
SLICER_LABELS = [(lab, sl_x0 + i * (SL_W + SL_GAP)) for i, (lab, _, _) in enumerate(slicer_defs)]


def add_chrome(pg):
    if pg.slicers:
        for i, (lab, ent, prop) in enumerate(slicer_defs):
            x = sl_x0 + i * (SL_W + SL_GAP)
            proj = fcolumn(ent, prop, active=True)
            pg.vis(f"slicer_{prop.replace(' ', '_').lower()}", "slicer", x, 22, SL_W, SL_H,
                   query={"queryState": {"Values": {"projections": [proj]}}, "sortDefinition": sort_by(proj, "Ascending")},
                   objects={"data": [{"properties": {"mode": s("Dropdown")}}],
                            "header": [{"properties": {"show": b(False)}}],
                            "selection": [{"properties": {"selectAllCheckboxEnabled": b(True), "singleSelect": b(False)}}],
                            "general": [{"properties": {"outlineColor": color(PANEL_BORDER), "outlineWeight": d(1)}}],
                            "items": [{"properties": {"fontColor": color(INK), "background": color(PANEL), "fontSize": d(9)}}]},
                   vc=no_title())
    # transparent navigation buttons over the painted sidebar items
    for i, target in enumerate(PAGES):
        ny = NAV_Y0 + i * NAV_STEP
        pg.vis(f"nav_{target.key}", "actionButton", 8, ny - 6, SIDEBAR_W - 16, NAV_H,
               objects={"icon": [{"properties": {"shapeType": s("blank")}, "selector": {"id": "default"}}, {"properties": {"show": b(False)}}],
                        "outline": [{"properties": {"show": b(False)}}],
                        "text": [{"properties": {"show": b(False)}}],
                        "fill": [{"properties": {"show": b(False)}}, {"properties": {"show": b(True), "fillColor": color(ACCENT), "transparency": d(80)}, "selector": {"id": "hover"}}]},
               vc={"visualLink": [{"properties": {"show": b(True), "type": s("PageNavigation"), "navigationSection": s(target.key),
                                                   "showDefaultTooltip": b(False), "tooltip": s(f"Go to {target.display}")}}],
                   "background": [{"properties": {"show": b(False)}}], "border": [{"properties": {"show": b(False)}}]})


for p in PAGES:
    add_chrome(p)

# field shortcuts
C_YEAR = lambda: fcolumn("Calendar", "Year", active=True)
C_FY = lambda: fcolumn("Calendar", "Fiscal Year", active=True)
C_Q = lambda: fcolumn("Calendar", "Quarter", active=True)
S_DATE = lambda: fcolumn("Calendar", "Quarter Start", active=True)   # quarter-start date on the Calendar so time intelligence works
S_COUNTY = lambda: fcolumn("Sales", "County", active=True)
S_REGION = lambda: fcolumn("Sales", "Region", active=True)
S_Q = lambda: fcolumn("Sales", "Quarter", active=True)
M = lambda n, disp=None: fmeasure("Sales", n, disp)

# =====================================================================================
# 5. PAGE: DASHBOARD
# =====================================================================================
pg = PAGE_BY_KEY["dashboard"]
KPI_Y, KPI_H, KPI_N = 66, 70, 6
KPI_W = (CW - GAP * (KPI_N - 1)) / KPI_N
kpis = [("Total Sales", "Total Sales (B)", "KPI Total Sales Sub", "Delta Color CAGR", "\uE825"),
        ("Latest Year Sales", "Latest Year Sales", "KPI Latest Year Sub", "Delta Color YoY", "\uE9D9"),
        ("Avg Quarterly Sales", "Avg Quarterly Sales", "KPI Avg Quarter Sub", None, "\uE787"),
        ("Best Quarter", "Best Quarter Sales", "KPI Best Quarter Sub", None, "\uE945"),
        ("Top County", "Top County", "KPI Top County Sub", None, "\uE707"),
        ("Counties Covered", "Counties", "KPI Counties Sub", None, "\uE804")]
for i, (lab, val, sub, subcol, glyph) in enumerate(kpis):
    x = X0 + i * (KPI_W + GAP)
    pg.panel(x, KPI_Y, KPI_W, KPI_H)
    pg.kpi_boxes.append((x, KPI_Y, KPI_W, KPI_H, lab, glyph))
    vx, vw = x + 46, KPI_W - 50
    pg.card(f"kpi_{i+1}_value", vx, KPI_Y + 15, vw, 32, M(val))
    pg.card(f"kpi_{i+1}_sub", vx, KPI_Y + 44, vw, 24, M(sub), size=8, col=INK2, col_measure=subcol, bold=False, font="Segoe UI")

ROW2_Y, ROW2_H = 146, 168
ROW3_Y, ROW3_H = 324, 176
ROW4_Y, ROW4_H = 510, 180
RIGHT_X = 824
LEFT_W = RIGHT_X - GAP - X0
RIGHT_W = X1 - RIGHT_X
RIGHT_H = 402
INS_Y = ROW2_Y + RIGHT_H + 10
INS_H = ROW4_Y + ROW4_H - INS_Y

pg.column("annual_sales", X0, ROW2_Y, LEFT_W, ROW2_H, C_YEAR(), M("Sales $B"), "Annual Sales ($B) · FY2008–FY2015")
pg.bar("top_counties", RIGHT_X, ROW2_Y, RIGHT_W, RIGHT_H, S_COUNTY(), M("Sales $B"), "Top Counties by Sales ($B)")
TREND_W = 372
pg.area("quarterly_trend", X0, ROW3_Y, TREND_W, ROW3_H, S_DATE(), M("Sales $B"), "Quarterly Sales Trend ($B)", fill=SERIES[0])
RB_X = X0 + TREND_W + GAP
pg.bar("region_bars", RB_X, ROW3_Y, RIGHT_X - GAP - RB_X, ROW3_H, S_REGION(), M("Sales $B"), "Sales by Region ($B)", fill=SERIES[1])
SEAS_W = 258
pg.donut("seasonality_donut", X0, ROW4_Y, SEAS_W, ROW4_H, S_Q(), M("Total Sales"), "Seasonality · Share by Quarter")
YOY_X = X0 + SEAS_W + GAP
pg.column("yoy_growth", YOY_X, ROW4_Y, RIGHT_X - GAP - YOY_X, ROW4_H, C_YEAR(), M("YoY %"), "Year-over-Year Growth %", fill_measure="YoY Bar Color")
pg.panel(RIGHT_X, INS_Y, RIGHT_W, INS_H)
pg.textbox("key_insights", RIGHT_X + 8, INS_Y + 4, RIGHT_W - 16, INS_H - 8, [
    ("Key Insights", "h"),
    ("●  Sales rose from $8.97B (FY2008) to $14.51B (FY2015): +61.8%, a 7.1% CAGR despite the -14.6% dip in FY2009.", "p"),
    ("●  FY2011 jumped +42.2%, FY2012 corrected -9.9%, then three straight growth years (+9.5%, +11.6%, +8.3%).", "p"),
    ("●  Arapahoe alone is 22.8% of sales and Denver Metro ≈ 62%; Q3 is the strongest quarter (27%), Q1 the weakest (23%).", "p"),
])

# =====================================================================================
# 6. PAGE: SALES TREND
# =====================================================================================
pg = PAGE_BY_KEY["trend"]
H1, H2 = 190, 200
Y1 = TOP; Y2 = Y1 + H1 + 10; Y3 = Y2 + H2 + 10; H3 = BOTTOM - Y3
pg.area("quarterly", X0, Y1, CW, H1, S_DATE(), M("Sales $B"), "Quarterly Sales ($B) · 32 quarters, FY2008 Q1 – FY2015 Q4", fill=SERIES[0])
LW = 540
pg.panel(X0, Y2, LW, H2)
pg.vis("sales_vs_py", "clusteredColumnChart", X0 + 6, Y2 + 4, LW - 12, H2 - 8,
       query={"queryState": {"Category": {"projections": [C_YEAR()]}, "Y": {"projections": [M("Sales $B", "Sales"), M("Sales PY $B", "Prior year")]}}, "sortDefinition": sort_by(C_YEAR(), "Ascending")},
       objects={"categoryAxis": [{"properties": dict(AX_CAT, axisType=s("Categorical"), fontSize=d(8))}],
                "valueAxis": [{"properties": AX_VAL_HIDDEN}],
                "labels": [{"properties": dict(LABELS, fontSize=d(7))}],
                "legend": [{"properties": LEGEND_TOP}],
                "dataPoint": [{"properties": {"fill": color(SERIES[0])}, "selector": {"metadata": "Sales.Sales $B"}},
                              {"properties": {"fill": color(SERIES[1])}, "selector": {"metadata": "Sales.Sales PY $B"}}],
                "general": [{"properties": {"responsive": b(False)}}]},
       vc=title("Annual Sales vs Prior Year ($B)"))
RX = X0 + LW + GAP
pg.column("qoq", RX, Y2, X1 - RX, H2, S_DATE(), M("QoQ %"), "Quarter-over-Quarter Growth %", fill=SERIES[1], scalar=True, labels=False, value_axis=True)
pg.line("rolling", X0, Y3, LW, H3, S_DATE(), [M("Rolling 4Q Sales $B", "Rolling 4 quarters"), M("Sales $B", "Quarter")], "Rolling 4-Quarter Sales vs Quarterly ($B)", colors=[SERIES[2], SERIES[1]])
pg.table("year_table", RX, Y3, X1 - RX, H3,
         [C_YEAR(), M("Sales $B", "Sales"), M("YoY %", "YoY"), M("Avg Quarterly Sales", "Avg quarter"), M("Best Quarter", "Best quarter"), M("Best Quarter Sales", "Best qtr sales")],
         "Year Summary", sort=C_YEAR(), sort_dir="Ascending", font=8)

# =====================================================================================
# 7. PAGE: COUNTIES
# =====================================================================================
pg = PAGE_BY_KEY["counties"]
LW = 400
pg.bar("ranking", X0, TOP, LW, BOTTOM - TOP, S_COUNTY(), M("Sales $B"), "Sales by County ($B) · FY2008–FY2015")
RX = X0 + LW + GAP; RW = X1 - RX
MH = 310
pg.matrix("county_year", RX, TOP, RW, MH, S_COUNTY(), C_YEAR(), [M("Sales $B", "$B")], "County × Year Sales ($B)", font=8)
Y3 = TOP + MH + 10; H3 = BOTTOM - Y3
HW = (RW - GAP) / 2
pg.bar("cagr", RX, Y3, HW, H3, S_COUNTY(), M("CAGR", "CAGR"), "Growth: CAGR first→last year in data", fill_measure="CAGR Bar Color", axis_font=7)
pg.bar("share", RX + HW + GAP, Y3, HW, H3, S_COUNTY(), M("Share of State %", "Share"), "Share of State Sales %", fill=SERIES[1], axis_font=7)

# =====================================================================================
# 8. PAGE: REGIONS
# =====================================================================================
pg = PAGE_BY_KEY["regions"]
LW = 650; RX = X0 + LW + GAP; RW = X1 - RX
H1 = 300; Y2 = TOP + H1 + 10; H2 = BOTTOM - Y2
REGION_COLORS = [("Denver Metro", SERIES[0]), ("S. Front Range", SERIES[1]), ("N. Front Range", SERIES[2]), ("Western Slope", SERIES[3]), ("Rest of State", SERIES[4])]
pg.series_chart("stacked_year", "columnChart", X0, TOP, LW, H1, C_YEAR(), S_REGION(), M("Sales $B"), "Sales by Region and Year ($B) · stacked")
pg.bar("share", RX, TOP, RW, H1, S_REGION(), M("Share of State %", "Share"), "Region Share of State Sales %", fill=SERIES[1])
pg.series_chart("trend", "lineChart", X0, Y2, LW, H2, S_DATE(), S_REGION(), M("Sales $B"), "Quarterly Sales by Region ($B)", scalar=True)
pg.table("region_table", RX, Y2, RW, H2,
         [S_REGION(), M("Counties", "Counties"), M("Sales $B", "Sales"), M("Share of State %", "Share"), M("CAGR", "CAGR")],
         "Region Summary", sort=M("Sales $B"), sort_dir="Descending", font=8)

# =====================================================================================
# 9. PAGE: QUARTERS
# =====================================================================================
pg = PAGE_BY_KEY["quarters"]
LW = 650; RX = X0 + LW + GAP; RW = X1 - RX
H1 = 300; Y2 = TOP + H1 + 10; H2 = BOTTOM - Y2
Q_COLORS = [("Q1", SERIES[0]), ("Q2", SERIES[1]), ("Q3", SERIES[2]), ("Q4", SERIES[3])]
pg.series_chart("by_quarter_year", "clusteredColumnChart", X0, TOP, LW, H1, C_YEAR(), C_Q(), M("Sales $B"), "Sales by Quarter and Year ($B)")
pg.donut("share_donut", RX, TOP, RW, H1, S_Q(), M("Total Sales"), "Share of Sales by Quarter", legend_pos="Bottom")
pg.matrix("year_quarter", X0, Y2, LW, H2, C_YEAR(), C_Q(), [M("Sales $B", "$B")], "Year × Quarter Sales ($B)", font=8)
pg.column("yoy_by_quarter", RX, Y2, RW, H2, S_DATE(), M("YoY %"), "Growth vs Same Quarter Last Year %", fill=SERIES[1], scalar=True, labels=False, value_axis=True)

# =====================================================================================
# 10. PAGE: DATA TABLE
# =====================================================================================
pg = PAGE_BY_KEY["data"]
TH = 560
pg.table("records", X0, TOP, CW, TH,
         [fcolumn("Calendar", "Year"), fcolumn("Calendar", "Quarter"), fcolumn("Sales", "County"), fcolumn("Sales", "Region"),
          M("Sales $M", "Sales ($M)"), M("Sales $B", "Sales ($B)"), M("YoY %", "YoY vs same quarter"), M("Share of State %", "Share of state")],
         None, sort=fcolumn("Calendar", "Year"), sort_dir="Ascending", font=9)
CY = TOP + TH + 10; CH = BOTTOM - CY
CWID = (CW - GAP * 2) / 3
for i, (lab, meas) in enumerate([("Records in selection", "Rows"), ("Total sales in selection", "Total Sales (B)"), ("Quarters in selection", "Quarters")]):
    x = X0 + i * (CWID + GAP)
    pg.panel(x, CY, CWID, CH)
    pg.kpi_boxes.append((x, CY, CWID, CH, lab, None))
    pg.card(f"sum_{i+1}", x + 10, CY + 12, CWID - 20, CH - 14, M(meas), size=13)

# =====================================================================================
# 11. PAGE: ABOUT
# =====================================================================================
pg = PAGE_BY_KEY["about"]
LW = 540; RX = X0 + LW + GAP; RW = X1 - RX
pg.panel(X0, TOP, LW, BOTTOM - TOP)
pg.textbox("about_left", X0 + 10, TOP + 6, LW - 20, BOTTOM - TOP - 12, [
    ("What this dashboard shows", "h"),
    ("Quarterly motor vehicle sales for 17 Colorado counties (plus a Rest of State bucket) across eight fiscal years, FY2008–FY2015. It is an executive-style view: headline KPIs, the growth path, county and regional concentration, and seasonality.", "p"),
    (" ", "p"),
    ("Data source", "b"),
    ("Colorado Department of Revenue, Motor Vehicle Sales by County (published on data.colorado.gov). 501 county-quarter records, four fields: year, quarter, county, sales in US dollars.", "p"),
    (" ", "p"),
    ("Pages", "b"),
    ("Dashboard – one-screen summary.  Sales Trend – quarterly series, prior-year comparison, quarter-over-quarter growth, rolling 4-quarter total.  Counties – ranking, county × year matrix, CAGR and share.  Regions – five county groups.  Quarters – seasonality and same-quarter growth.  Data – every record, filterable.", "p"),
    (" ", "p"),
    ("Region definitions", "b"),
    ("Denver Metro: Adams, Arapahoe, Denver, Douglas, Jefferson, Boulder, Broomfield.  N. Front Range: Larimer, Weld.  S. Front Range: El Paso, Pueblo, Fremont.  Western Slope: Mesa, Garfield, La Plata.  Rest of State: all other counties, as reported by the source.", "p"),
    (" ", "p"),
    ("Author", "b"),
    ("Vivek Yadav · MBA Finance · Financial Analyst portfolio project. Companion Excel analysis: github.com/vivek-yadav-02/Colorado-MV-Financial-Analysis", "p"),
])
pg.panel(RX, TOP, RW, BOTTOM - TOP)
pg.textbox("about_right", RX + 10, TOP + 6, RW - 20, BOTTOM - TOP - 12, [
    ("Definitions and method", "h"),
    ("Sales", "b"), ("Sum of reported sales in nominal US dollars. $B = billions.", "p"),
    ("YoY %", "b"), ("Change versus the same period one year earlier (same quarter, or same year), using the Calendar date table.", "p"),
    ("QoQ %", "b"), ("Change versus the immediately preceding quarter.", "p"),
    ("CAGR", "b"), ("(Last-year sales ÷ first-year sales)^(1 ÷ years) − 1, computed over the years present in the current selection.", "p"),
    ("Rolling 4Q", "b"), ("Sum of the latest four quarters; shown only once four quarters are available.", "p"),
    ("Share of state", "b"), ("County or region sales ÷ statewide sales for the same period.", "p"),
    (" ", "p"),
    ("Caveats", "h"),
    ("●  Boulder and Broomfield are reported as one combined line in FY2008 and separately from FY2009.", "p"),
    ("●  The Rest of State bucket starts in FY2009 Q4; earlier statewide totals therefore cover only the named counties.", "p"),
    ("●  Figures are nominal; no inflation adjustment. Fiscal-year labels follow the source's year field.", "p"),
    ("●  The sidebar items Vehicles/Reports/Settings from the original design were replaced with real pages; there is no vehicle-type data in this source.", "p"),
    (" ", "p"),
    ("Build", "b"),
    ("Power BI project (PBIP, PBIR report format) generated from a Python script: TMDL model with 38 DAX measures, one JSON file per visual, a custom dark theme and painted page backgrounds. Colour palette validated for colour-vision deficiency and contrast.", "p"),
])

# =====================================================================================
# 12. WRITE REPORT FILES
# =====================================================================================
report = {
    "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/1.0.0/schema.json",
    "themeCollection": {"baseTheme": {"name": "CY24SU06", "type": "SharedResources", "reportVersionAtImport": "5.55"},
                        "customTheme": {"name": THEME_FILE, "type": "RegisteredResources", "reportVersionAtImport": "5.55"}},
    "layoutOptimization": "None",
    "resourcePackages": [
        {"name": "SharedResources", "type": "SharedResources", "items": [{"name": "CY24SU06", "path": "BaseThemes/CY24SU06.json", "type": "BaseTheme"}]},
        {"name": "RegisteredResources", "type": "RegisteredResources", "items": [{"name": THEME_FILE, "path": THEME_FILE, "type": "CustomTheme"}] +
         [{"name": f"bg_{p.key}.png", "path": f"bg_{p.key}.png", "type": "Image"} for p in PAGES]},
    ],
    "settings": {"useStylableVisualContainerHeader": True, "defaultDrillFilterOtherVisuals": True, "useEnhancedTooltips": True},
}
dump(os.path.join(DEF, "report.json"), report)
dump(os.path.join(PAGES_DIR, "pages.json"), {"$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.0.0/schema.json",
                                              "pageOrder": [p.key for p in PAGES], "activePageName": "dashboard"})
for p in PAGES:
    pdir = os.path.join(PAGES_DIR, p.key)
    os.makedirs(os.path.join(pdir, "visuals"), exist_ok=True)
    dump(os.path.join(pdir, "page.json"), {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/1.0.0/schema.json",
        "name": p.key, "displayName": p.display, "displayOption": "FitToPage", "height": 720, "width": 1280,
        "objects": {"background": [{"properties": {
            "image": {"image": {"name": s(f"bg_{p.key}.png"),
                                "url": {"expr": {"ResourcePackageItem": {"PackageName": "RegisteredResources", "PackageType": 1, "ItemName": f"bg_{p.key}.png"}}},
                                "scaling": s("Fit")}},
            "transparency": d(0)}}],
            "outspace": [{"properties": {"color": color(PAGE_BG)}}]}})
    for v in p.visuals:
        vdir = os.path.join(pdir, "visuals", v["name"])
        os.makedirs(vdir, exist_ok=True)
        dump(os.path.join(vdir, "visual.json"), v)

# =====================================================================================
# 13. BACKGROUND IMAGES (2x, one per page)
# =====================================================================================
S = 2
W, H = 1280 * S, 720 * S


def font(name, size): return ImageFont.truetype(f"C:/Windows/Fonts/{name}", int(size * S))
F_REG = lambda sz: font("segoeui.ttf", sz)
F_BOLD = lambda sz: font("segoeuib.ttf", sz)
F_LIGHT = lambda sz: font("segoeuil.ttf", sz)
F_ICON = lambda sz: font("segmdl2.ttf", sz)


def hexa(h, a):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4)) + (a,)


def paint(pg, active_idx):
    img = Image.new("RGB", (W, H), PAGE_BG)
    dr = ImageDraw.Draw(img, "RGBA")

    def rr(x, y, w, h, r, fill, outline=None, width=1):
        dr.rounded_rectangle([x * S, y * S, (x + w) * S, (y + h) * S], radius=r * S, fill=fill, outline=outline, width=width * S)

    def text(x, y, t, f, fill, anchor="la"):
        dr.text((x * S, y * S), t, font=f, fill=fill, anchor=anchor)

    for yy in range(H):
        t = yy / H
        dr.line([(0, yy), (W, yy)], fill=tuple(int(a + (bb - a) * t) for a, bb in zip((11, 18, 32), (9, 14, 26))))
    # sidebar
    dr.rectangle([0, 0, SIDEBAR_W * S, H], fill=SIDEBAR)
    dr.line([(SIDEBAR_W * S, 0), (SIDEBAR_W * S, H)], fill=PANEL_BORDER, width=S)
    text(16, 18, "COLORADO", F_BOLD(17), ACCENT)
    text(16, 40, "Motor Vehicle Sales", F_REG(9), INK2)
    dr.line([(16 * S, 58 * S), ((SIDEBAR_W - 16) * S, 58 * S)], fill=PANEL_BORDER, width=S)
    for i, p in enumerate(PAGES):
        ny = NAV_Y0 + i * NAV_STEP
        active = i == active_idx
        if active:
            rr(8, ny - 6, SIDEBAR_W - 16, NAV_H, 6, hexa(ACCENT, 38))
            dr.rectangle([8 * S, (ny - 6) * S, 11 * S, (ny + NAV_H - 6) * S], fill=ACCENT)
        text(22, ny + 1, p.nav_glyph, F_ICON(12), INK if active else INK2)
        text(44, ny, p.nav_label, F_BOLD(10) if active else F_REG(10), INK if active else INK2)
    # mountains
    mt = 470
    pts = [(0, H), (0, (mt + 90) * S), (20 * S, (mt + 60) * S), (45 * S, (mt + 95) * S), (70 * S, (mt + 30) * S), (90 * S, (mt + 70) * S),
           (110 * S, (mt + 45) * S), (130 * S, (mt + 85) * S), (SIDEBAR_W * S, (mt + 65) * S), (SIDEBAR_W * S, H)]
    dr.polygon(pts, fill=hexa(ACCENT, 60))
    pts2 = [(0, H), (0, (mt + 120) * S), (30 * S, (mt + 95) * S), (60 * S, (mt + 125) * S), (85 * S, (mt + 85) * S), (115 * S, (mt + 115) * S),
            (SIDEBAR_W * S, (mt + 100) * S), (SIDEBAR_W * S, H)]
    dr.polygon(pts2, fill=hexa(ACCENT, 110))
    text(16, 598, "Drive the", F_LIGHT(13), INK)
    text(16, 614, "Front Range", F_BOLD(13), INK)
    text(16, 650, "Data: Colorado Dept. of", F_REG(7), INK2)
    text(16, 661, "Revenue · FY2008–FY2015", F_REG(7), INK2)
    text(16, 680, "\uE787", F_ICON(9), INK2)
    text(30, 679, "Quarterly · 17 counties", F_REG(7), INK2)
    # header
    text(X0, 12, pg.title_text, F_BOLD(20), INK)
    text(X0, 40, pg.subtitle, F_REG(9), INK2)
    if pg.slicers:
        for lab, x in SLICER_LABELS:
            text(x + 2, 8, lab.upper(), F_BOLD(7), MUTED)
    cx, cy = X0 + 300, 30
    if pg.key == "dashboard":
        dr.ellipse([(cx - 14) * S, (cy - 14) * S, (cx + 14) * S, (cy + 14) * S], outline=ACCENT, width=5 * S)
        dr.rectangle([(cx + 2) * S, (cy - 6) * S, (cx + 18) * S, (cy + 6) * S], fill=(14, 23, 41))
        dr.ellipse([(cx - 6) * S, (cy - 6) * S, (cx + 6) * S, (cy + 6) * S], fill=GOLD)
    # panels
    for (x, y, w, h) in pg.panels:
        rr(x, y, w, h, 8, PANEL, outline=PANEL_BORDER, width=1)
    for (x, y, w, h, lab, glyph) in pg.kpi_boxes:
        if glyph:
            rr(x + 10, y + 19, 30, 30, 7, hexa(ACCENT, 70))
            text(x + 25, y + 34, glyph, F_ICON(13), ACCENT, anchor="mm")
            text(x + 46 + (w - 52) / 2, y + 12, lab, F_REG(8), INK2, anchor="mm")
        else:
            text(x + w / 2, y + 9, lab, F_REG(8), INK2, anchor="mm")
    # footer
    dr.line([(X0 * S, (FOOT_Y - 2) * S), (X1 * S, (FOOT_Y - 2) * S)], fill=PANEL_BORDER, width=S)
    text(X0, FOOT_Y + 6, "COLORADO", F_BOLD(9), ACCENT)
    text(X0 + 62, FOOT_Y + 7, f"Motor Vehicle Sales · {pg.display}", F_REG(8), INK2)
    text(X1, FOOT_Y + 7, "Dashboard by Vivek Yadav   |   Source: Colorado Department of Revenue   |   Built with Power BI", F_REG(8), INK2, anchor="ra")
    img.save(os.path.join(RES, f"bg_{pg.key}.png"), optimize=True)
    if pg.key == "dashboard":
        img.resize((1280, 720), Image.LANCZOS).save(os.path.join(OUT, "background_preview.png"))


for i, p in enumerate(PAGES):
    paint(p, i)
print("pages:", len(PAGES), "visuals:", sum(len(p.visuals) for p in PAGES), "measures:", len(measures))
print("written to", OUT)
