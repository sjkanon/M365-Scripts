"""
Combined Licensing & Azure Cost Report
=======================================
Reads a Pax8 CSV and an Ingram Excel, combines all customers,
and writes one Excel file with a dedicated tab per customer.

Usage:
    python genereer_licentie_overzicht.py --ingram <ingram.xlsx> --pax8 <pax8.csv>
    python genereer_licentie_overzicht.py --ingram <ingram.xlsx> --pax8 <pax8.csv> --output report.xlsx

Output (default):  Licensing_Report_YYYY-MM.xlsx

Requirements:  pip install pandas openpyxl
"""

import sys
import re

def _pause_if_interactive(prompt: str = "Press Enter to exit..."):
    try:
        if sys.stdin and sys.stdin.isatty():
            input(prompt)
    except Exception:
        pass

MONTHS_EN = {
    1:"January", 2:"February", 3:"March",    4:"April",
    5:"May",     6:"June",     7:"July",      8:"August",
    9:"September",10:"October",11:"November", 12:"December"
}

def _verbose_period(period_str: str) -> str:
    """'2025-12' -> 'December 2025'"""
    try:
        y, m = period_str[:7].split("-")
        return f"{MONTHS_EN[int(m)]} {y}"
    except Exception:
        return period_str


# Mapping: Acronis end-customer name in Pax8 -> customer name in the report
# Add entries here if the name in Pax8 differs from the name used elsewhere.
# Example: "Customer Name in Pax8": "Customer Name in Report"
ACRONIS_ENDCUSTOMER_ALIASES = {
}

import pandas as pd
import warnings
warnings.filterwarnings("ignore")
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Colours ───────────────────────────────────────────────────────────────────
DARK_BLUE   = "1F3864"
MED_BLUE    = "2E75B6"
AZURE_BLUE  = "4472C4"   # Azure section header
LIC_GREEN   = "375623"   # Licenses section header (dark green)
LIC_MED     = "538135"   # Licenses sub-header
ROW_LIGHT   = "EAF2FB"
ROW_WHITE   = "FFFFFF"
CAT_GREY    = "F2F2F2"
INGRAM_TEAL = "1F5C6B"   # Ingram section

# ── Style helpers ─────────────────────────────────────────────────────────────
def _side():
    return Side(style="thin", color="C0C0C0")

def _border():
    s = _side()
    return Border(left=s, right=s, top=s, bottom=s)

def _fill(hex_color):
    return PatternFill("solid", start_color=hex_color, end_color=hex_color)

def _money(cell):
    cell.number_format = "€#,##0.00"

def _int_fmt(cell):
    cell.number_format = "#,##0"

def _styled(ws, row, col, value, bg, fg="FFFFFF", bold=True, size=10,
            halign="center", valign="center", wrap=False, merge_to=None,
            number_format=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font      = Font(name="Arial", bold=bold, color=fg, size=size)
    c.fill      = _fill(bg)
    c.alignment = Alignment(horizontal=halign, vertical=valign, wrap_text=wrap)
    c.border    = _border()
    if merge_to:
        ws.merge_cells(start_row=row, start_column=col,
                       end_row=row,   end_column=merge_to)
    if number_format:
        c.number_format = number_format
    return c

def _period_label(start, end):
    """Format a billing period as 'dd/mm – dd/mm/yyyy', or '' on failure."""
    try:
        s = pd.to_datetime(start)
        e = pd.to_datetime(end)
        return f"{s.strftime('%d/%m')} \u2013 {e.strftime('%d/%m/%Y')}"
    except Exception:
        return ""


def _empty_row(ws, row, ncols=7):
    for col in range(1, ncols + 1):
        ws.cell(row=row, column=col).border = Border()
    ws.row_dimensions[row].height = 8


# ── Pax8 parser ───────────────────────────────────────────────────────────────

PAX8_AZURE_CATEGORIES = {
    "Virtual Machines":          "Virtual Machines",
    "Virtual Machines Licenses": "Virtual Machines",
    "Storage":                   "Storage",
    "Backup":                    "Backup",
    "Bandwidth":                 "Network & Traffic",
    "Virtual Network":           "Network & Traffic",
    "Load Balancer":             "Network & Traffic",
    "Azure DNS":                 "Network & Traffic",
    "Log Analytics":             "Monitoring & Management",
    "Azure Monitor":             "Monitoring & Management",
    "Key Vault":                 "Monitoring & Management",
    "Logic Apps":                "Monitoring & Management",
    "Azure App Service":         "App Services",
    "SQL Database":              "App Services",
    "Azure DevOps":              "App Services",
}

AZURE_CATEGORY_ORDER = [
    "Virtual Machines", "Storage", "Backup",
    "Network & Traffic", "Monitoring & Management", "App Services", "Other",
]

# Optional: map raw Pax8 subscription IDs to human-readable labels.
# Add entries here to give subscriptions a friendly display name.
# Example: "raw-sub-id-001": "Friendly Subscription Name"
PAX8_SUB_LABELS = {
}

# Optional: define the display order of subscriptions within an Azure section.
# Subscriptions not listed here appear after the listed ones, in their original order.
PAX8_SUB_ORDER = [
]


def _parse_pax8_desc(desc):
    desc = re.sub(r'\s*\[options:.*?\]\s*$', '', desc).strip()
    parts = desc.split(" - ")
    sub     = parts[2].strip() if len(parts) > 2 else "Unknown"
    cat_raw = parts[4].strip() if len(parts) > 4 else "Other"
    return sub, PAX8_AZURE_CATEGORIES.get(cat_raw, "Other")


def load_pax8(csv_path: str):
    df = pd.read_csv(csv_path, sep=None, engine="python")
    df["company_name"] = df["company_name"].str.strip()

    # Split Azure vs licenties op basis van description
    is_azure = df["description"].str.contains(
        r"Microsoft Azure Plan - Arrears Charge", case=False, na=False
    )

    azure = df[is_azure].copy()
    azure[["subscription", "az_category"]] = azure["description"].apply(
        lambda x: pd.Series(_parse_pax8_desc(x))
    )

    licenses = df[~is_azure].copy()
    # Product name from description (everything before " - " or the full string)
    licenses["product"] = licenses["description"].str.split(" - ").str[0].str.strip()
    licenses["product"] = licenses["product"].str.replace(r'\s*\[options:.*?\]', '', regex=True).str.strip()
    licenses["product"] = licenses["product"].str.replace(r'^\[Deprecated\]\s*', '', regex=True).str.strip()

    # Acronis: extract end-customer from description
    # Format: "Product - Product - EndCustomer - qty [options:]"
    def _acronis_endcustomer(row):
        if 'acronis' not in str(row['description']).lower():
            return ''
        desc = re.sub(r'\s*\[options:.*?\]\s*$', '', str(row['description'])).strip()
        parts = [p.strip() for p in desc.split(' - ')]
        # Last part is qty (number), one before that is the end-customer
        for i in range(len(parts)-1, -1, -1):
            try:
                float(parts[i])
                continue
            except ValueError:
                if i > 0 and parts[i] not in parts[:i]:
                    return parts[i]
                break
        return ''
    licenses['acronis_endcustomer'] = licenses.apply(_acronis_endcustomer, axis=1)
    # Apply end-customer name aliases (see ACRONIS_ENDCUSTOMER_ALIASES at the top)
    licenses['acronis_endcustomer'] = licenses['acronis_endcustomer'].replace(ACRONIS_ENDCUSTOMER_ALIASES)

    return azure, licenses


# ── Ingram parser ─────────────────────────────────────────────────────────────

INGRAM_CATEGORY_MAP = {
    "Virtual Machines":             "Virtual Machines",
    "Virtual Machines Licenses":    "Virtual Machines",
    "Storage":                      "Storage",
    "Backup":                       "Backup",
    "Bandwidth":                    "Network & Traffic",
    "Virtual Network":              "Network & Traffic",
    "Load Balancer":                "Network & Traffic",
    "Azure DNS":                    "Network & Traffic",
    "VPN Gateway":                  "Network & Traffic",
    "Network Watcher":              "Network & Traffic",
    "Log Analytics":                "Monitoring & Management",
    "Azure Monitor":                "Monitoring & Management",
    "Microsoft Defender for Cloud": "Monitoring & Management",
    "Key Vault":                    "Monitoring & Management",
    "Logic Apps":                   "Monitoring & Management",
    "Azure App Service":            "App Services",
    "SQL Database":                 "App Services",
    "Azure Databricks":             "App Services",
    "Azure DevOps":                 "App Services",
    "Reserved VM Instance":         "Reserved Instances (RI)",
}


def _parse_ingram_azure_desc(desc):
    """
    Two formats in Ingram:
      1) '<UUID> <Service Category> from ...'   -> subscription=UUID, cat=service
      2) '# <UUID> Reserved VM Instance, ...'  -> cat=Reserved Instances (RI)
    """
    desc = str(desc).strip()
    if desc.startswith("#"):
        m = re.match(r'#\s*([\w-]+)\s+Reserved VM Instance,\s*([^,]+)', desc)
        if m:
            return "Reserved Instances (RI)", "Reserved Instances (RI)"
        return "RI", "Reserved Instances (RI)"
    m = re.match(r'([\w-]+)\s+(.+?)\s+from\s+', desc)
    if m:
        sub     = m.group(1)
        cat_raw = m.group(2).strip()
        return sub, INGRAM_CATEGORY_MAP.get(cat_raw, "Other")
    return "Unknown", "Other"


def load_ingram(xlsx_path: str):
    import warnings
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        df = pd.read_excel(xlsx_path)

    df["product"] = df["CUSTOMER_RESOURCE_NAME"].str.extract(r'en ([^\t]+)')
    df["product"] = df["product"].fillna(df["CUSTOMER_RESOURCE_NAME"].str.strip())

    is_azure = df["product"].str.contains("Azure", case=False, na=False)

    azure = df[is_azure].copy()
    azure[["az_subscription", "az_category"]] = azure["CUSTOMER_DETAIL_DESCRIPTION"].apply(
        lambda x: pd.Series(_parse_ingram_azure_desc(x))
    )

    licenses = df[~is_azure].copy()
    return azure, licenses


# ── Collect all customers ─────────────────────────────────────────────────────

def gather_customers(pax8_az, pax8_lic, ingram_az, ingram_lic):
    """Returns a sorted list of all unique customer names across both sources."""
    customers = set()
    for frame, col in [
        (pax8_az,    "company_name"),
        (pax8_lic,   "company_name"),
        (ingram_az,  "CUSTOMER_NAME"),
        (ingram_lic, "CUSTOMER_NAME"),
    ]:
        if col in frame.columns:
            customers.update(frame[col].dropna().unique())
    return sorted(customers)


# ── Write customer sheet ──────────────────────────────────────────────────────

# Column layout:
# A: Description / Subscription / Product
# B: Category / Detail
# C: Qty
# D: Unit purchase price
# E: Unit sales price
# F: Total purchase
# G: Total sales

NCOLS = 7

def _col_headers(ws, row):
    headers = [
        ("Description",         "left"),
        ("Category / Detail",   "left"),
        ("Qty",                 "center"),
        ("Purchase Price",      "center"),
        ("Sales Price",         "center"),
        ("Total Purchase (€)",  "center"),
        ("Total Sales (€)",     "center"),
    ]
    for col, (txt, align) in enumerate(headers, 1):
        c = _styled(ws, row, col, txt, bg=MED_BLUE, size=9)
        c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 28


def _section_header(ws, row, label, bg, fg="FFFFFF"):
    _styled(ws, row, 1, label, bg=bg, size=10, halign="left", merge_to=NCOLS)
    for col in range(2, NCOLS + 1):
        ws.cell(row=row, column=col).fill   = _fill(bg)
        ws.cell(row=row, column=col).border = _border()
    ws.row_dimensions[row].height = 20


def _sub_header(ws, row, label, inkoop_total, verkoop_total, bg, fg="FFFFFF"):
    _styled(ws, row, 1, label, bg=bg, fg=fg, size=9, halign="left", merge_to=5)
    for col in [2, 3, 4, 5]:
        ws.cell(row=row, column=col).fill   = _fill(bg)
        ws.cell(row=row, column=col).border = _border()
    for col_idx, val in [(6, inkoop_total), (7, verkoop_total)]:
        c = ws.cell(row=row, column=col_idx, value=val)
        c.font      = Font(name="Arial", bold=True, color=fg, size=9)
        c.fill      = _fill(bg)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = _border()
        _money(c)
    ws.row_dimensions[row].height = 18


def _data_row(ws, row, description, detail, qty, unit_purchase, unit_sales,
              tot_purchase, tot_sales, alt=False):
    bg = ROW_LIGHT if alt else ROW_WHITE
    fg = "333333"

    def _dc(col, val, align="left", fmt=None):
        c = ws.cell(row=row, column=col, value=val)
        c.font      = Font(name="Arial", size=9, color=fg)
        c.fill      = _fill(bg)
        c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
        c.border    = _border()
        if fmt:
            c.number_format = fmt
        return c

    _dc(1, description,  "left")
    _dc(2, detail,       "left")
    _dc(3, qty,          "center", "#,##0")
    _dc(4, unit_purchase,"right",  "€#,##0.00")
    _dc(5, unit_sales,   "right",  "€#,##0.00")
    _dc(6, tot_purchase, "right",  "€#,##0.00")
    _dc(7, tot_sales,    "right",  "€#,##0.00")
    ws.row_dimensions[row].height = 16


def _totaal_row(ws, row, label, inkoop, verkoop):
    _styled(ws, row, 1, label, bg=DARK_BLUE, size=10,
            halign="right", merge_to=5)
    for col in [2, 3, 4, 5]:
        ws.cell(row=row, column=col).fill   = _fill(DARK_BLUE)
        ws.cell(row=row, column=col).border = _border()
    for col_idx, val in [(6, inkoop), (7, verkoop)]:
        c = ws.cell(row=row, column=col_idx, value=val)
        c.font      = Font(name="Arial", bold=True, color="FFFFFF", size=11)
        c.fill      = _fill(DARK_BLUE)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = _border()
        _money(c)
    ws.row_dimensions[row].height = 26


def write_customer_sheet(wb, customer, period, ingram_lic_period,
                         pax8_az, pax8_lic, ingram_az, ingram_lic,
                         pax8_lic_all=None):

    # Sheet name max 31 chars, no special characters
    sheet_name = re.sub(r'[\\/*?:\[\]]', '', customer)[:31]
    ws = wb.create_sheet(title=sheet_name)

    # Column widths
    widths = [32, 26, 8, 13, 13, 15, 15]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    row = 1

    # ── Title row ──
    _styled(ws, row, 1,
            f"{customer}  —  Licensing & Services Overview  |  {_verbose_period(period)}",
            bg=DARK_BLUE, size=13, merge_to=NCOLS)
    ws.row_dimensions[row].height = 32
    row += 1

    # ── Column headers ──
    _col_headers(ws, row)
    row += 1

    grand_purchase = 0.0
    grand_sales    = 0.0

    # ════════════════════════════════════════════════════════
    # SECTION 1: PAX8 AZURE
    # ════════════════════════════════════════════════════════
    cust_az = pax8_az[pax8_az["company_name"] == customer]
    if not cust_az.empty:
        _section_header(ws, row, "☁  Azure (via Pax8)", AZURE_BLUE)
        row += 1

        subs = cust_az["subscription"].unique().tolist()
        ordered_subs = [s for s in PAX8_SUB_ORDER if s in subs]
        ordered_subs += [s for s in subs if s not in ordered_subs]

        for sub in ordered_subs:
            sub_data     = cust_az[cust_az["subscription"] == sub]
            sub_purchase = sub_data["cost_total"].sum()
            sub_sales    = sub_data["subtotal"].sum()
            sub_label    = PAX8_SUB_LABELS.get(sub, sub)

            _sub_header(ws, row, sub_label, sub_purchase, sub_sales,
                        bg="2E5F8E", fg="FFFFFF")
            row += 1

            alt = False
            for cat in AZURE_CATEGORY_ORDER:
                cat_data = sub_data[sub_data["az_category"] == cat]
                if cat_data.empty:
                    continue
                _data_row(ws, row, "", cat, "",
                          "", "",
                          cat_data["cost_total"].sum(),
                          cat_data["subtotal"].sum(),
                          alt=alt)
                alt = not alt
                row += 1

        grand_purchase += cust_az["cost_total"].sum()
        grand_sales    += cust_az["subtotal"].sum()

        _empty_row(ws, row)
        row += 1

    # ════════════════════════════════════════════════════════
    # SECTION 2: INGRAM AZURE
    # ════════════════════════════════════════════════════════
    cust_iaz = ingram_az[ingram_az["CUSTOMER_NAME"] == customer]
    if not cust_iaz.empty:
        start = pd.to_datetime(cust_iaz["RESELLER_DETAIL_START_DATE"]).min()
        end   = pd.to_datetime(cust_iaz["RESELLER_DETAIL_END_DATE"]).max()
        az_period = f"{start.strftime('%d/%m/%Y')} – {end.strftime('%d/%m/%Y')}"
        _section_header(ws, row, f"☁  Azure (via Ingram)  —  usage period {az_period}", INGRAM_TEAL)
        row += 1

        subs = cust_iaz["az_subscription"].unique().tolist()
        # Reserved Instances shown last
        ri_data     = cust_iaz[cust_iaz["az_category"] == "Reserved Instances (RI)"]
        normal_subs = [s for s in subs if s != "Reserved Instances (RI)"]

        for sub in normal_subs:
            sub_data     = cust_iaz[cust_iaz["az_subscription"] == sub]
            sub_purchase = sub_data["RESELLER_DETAIL_TOTAL"].sum()
            sub_sales    = sub_data["CUSTOMER_DETAIL_TOTAL"].sum()

            _sub_header(ws, row, sub, sub_purchase, sub_sales,
                        bg="2E5F8E", fg="FFFFFF")
            row += 1

            alt = False
            for cat in AZURE_CATEGORY_ORDER:
                cat_data = sub_data[sub_data["az_category"] == cat]
                if cat_data.empty:
                    continue
                _data_row(ws, row, "", cat, "",
                          "", "",
                          cat_data["RESELLER_DETAIL_TOTAL"].sum(),
                          cat_data["CUSTOMER_DETAIL_TOTAL"].sum(),
                          alt=alt)
                alt = not alt
                row += 1

        # Reserved Instances as a separate section
        if not ri_data.empty:
            _sub_header(ws, row, "Reserved Instances (RI)",
                        ri_data["RESELLER_DETAIL_TOTAL"].sum(),
                        ri_data["CUSTOMER_DETAIL_TOTAL"].sum(),
                        bg="2E5F8E", fg="FFFFFF")
            row += 1
            alt = False
            for _, r in ri_data.iterrows():
                desc = str(r["CUSTOMER_DETAIL_DESCRIPTION"])
                m = __import__('re').match(r'#\s*[\w-]+\s+(Reserved VM Instance,\s*[^-]+)', desc)
                ri_label = m.group(1).strip() if m else desc[:60]
                _data_row(ws, row, "", ri_label, r["CUSTOMER_DETAIL_QTY"],
                          r["RESELLER_DETAIL_UNIT_PRICE"],
                          r["CUSTOMER_DETAIL_UNIT_PRICE"],
                          r["RESELLER_DETAIL_TOTAL"],
                          r["CUSTOMER_DETAIL_TOTAL"],
                          alt=alt)
                alt = not alt
                row += 1

        grand_purchase += cust_iaz["RESELLER_DETAIL_TOTAL"].sum()
        grand_sales    += cust_iaz["CUSTOMER_DETAIL_TOTAL"].sum()

        _empty_row(ws, row)
        row += 1

    # ════════════════════════════════════════════════════════
    # SECTION 3: PAX8 LICENSES
    # ════════════════════════════════════════════════════════
    cust_pl = pax8_lic[pax8_lic["company_name"] == customer]
    if not cust_pl.empty:
        _section_header(ws, row, "📋  Licenses (via Pax8)", LIC_GREEN)
        row += 1

        # Group by product; split Acronis per end-customer, group the rest normally
        non_acronis = cust_pl[cust_pl["acronis_endcustomer"] == ""]
        acronis     = cust_pl[cust_pl["acronis_endcustomer"] != ""]

        grp = non_acronis.groupby("product").agg(
            qty=("quantity", "sum"),
            purchase=("cost_total", "sum"),
            sales=("subtotal", "sum"),
            u_pur=("cost", "mean"),
            u_sal=("price", "mean"),
        ).reset_index()

        alt = False
        for _, r in grp.iterrows():
            _data_row(ws, row, r["product"], "", r["qty"],
                      r["u_pur"], r["u_sal"],
                      r["purchase"], r["sales"], alt=alt)
            alt = not alt
            row += 1

        # Acronis: sub-group per end-customer
        if not acronis.empty:
            for endcustomer, ec_data in acronis.groupby("acronis_endcustomer"):
                ec_purchase = ec_data["cost_total"].sum()
                ec_sales    = ec_data["subtotal"].sum()
                _sub_header(ws, row, f"Acronis  —  {endcustomer}",
                            ec_purchase, ec_sales, bg="4A4A8A", fg="FFFFFF")
                row += 1
                alt = False
                for _, r in ec_data.iterrows():
                    _data_row(ws, row, r["product"], "", r["quantity"],
                              r["cost"], r["price"],
                              r["cost_total"], r["subtotal"], alt=alt)
                    alt = not alt
                    row += 1

        grand_purchase += cust_pl["cost_total"].sum()
        grand_sales    += cust_pl["subtotal"].sum()

        _empty_row(ws, row)
        row += 1

    # ════════════════════════════════════════════════════════
    # SECTION 4: INGRAM LICENSES
    # ════════════════════════════════════════════════════════
    cust_il = ingram_lic[ingram_lic["CUSTOMER_NAME"] == customer]
    if not cust_il.empty:
        _section_header(ws, row, "📋  Licenses (via Ingram)", LIC_MED)
        row += 1

        grp = cust_il.groupby("product").agg(
            qty=("CUSTOMER_DETAIL_QTY", "sum"),
            purchase=("RESELLER_DETAIL_TOTAL", "sum"),
            sales=("CUSTOMER_DETAIL_TOTAL", "sum"),
            u_pur=("RESELLER_DETAIL_UNIT_PRICE", "mean"),
            u_sal=("CUSTOMER_DETAIL_UNIT_PRICE", "mean"),
        ).reset_index()

        alt = False
        for _, r in grp.iterrows():
            _data_row(ws, row, r["product"], "", r["qty"],
                      r["u_pur"], r["u_sal"],
                      r["purchase"], r["sales"], alt=alt)
            alt = not alt
            row += 1

        grand_purchase += cust_il["RESELLER_DETAIL_TOTAL"].sum()
        grand_sales    += cust_il["CUSTOMER_DETAIL_TOTAL"].sum()

        _empty_row(ws, row)
        row += 1

    # ════════════════════════════════════════════════════════
    # SECTION 5: ACRONIS (customer is end-customer billed via reseller)
    # ════════════════════════════════════════════════════════
    if pax8_lic_all is not None:
        acronis_for_cust = pax8_lic_all[
            (pax8_lic_all["acronis_endcustomer"] == customer) &
            (pax8_lic_all["acronis_endcustomer"] != "")
        ]
        if not acronis_for_cust.empty:
            _section_header(ws, row, "📋  Acronis Backup (via Reseller)", "4A4A8A")
            row += 1
            alt = False
            for _, r in acronis_for_cust.iterrows():
                _data_row(ws, row, r["product"], "", r["quantity"],
                          r["cost"], r["price"],
                          r["cost_total"], r["subtotal"], alt=alt)
                alt = not alt
                row += 1
            grand_purchase += acronis_for_cust["cost_total"].sum()
            grand_sales    += acronis_for_cust["subtotal"].sum()
            _empty_row(ws, row)
            row += 1

    # ════════════════════════════════════════════════════════
    # GRAND TOTAL
    # ════════════════════════════════════════════════════════
    _totaal_row(ws, row, "TOTAL", grand_purchase, grand_sales)
    row += 1

    return grand_purchase, grand_sales


# ── Summary sheet ─────────────────────────────────────────────────────────────

def write_summary_sheet(wb, summary_rows, period):
    ws = wb.active
    ws.title = "Overview"

    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 14

    row = 1
    _styled(ws, row, 1,
            f"Licensing & Services Overview — All Customers  |  {_verbose_period(period)}",
            bg=DARK_BLUE, size=13, merge_to=4)
    ws.row_dimensions[row].height = 32
    row += 1

    for col, txt in enumerate(["Customer", "Total Purchase (€)", "Total Sales (€)", "Margin (€)"], 1):
        _styled(ws, row, col, txt, bg=MED_BLUE, size=10)
    ws.row_dimensions[row].height = 22
    row += 1

    total_purchase = total_sales = 0.0
    for i, (name, purchase, sales) in enumerate(summary_rows):
        bg = ROW_LIGHT if i % 2 else ROW_WHITE
        fg = "333333"
        margin = sales - purchase

        def _sc(col, val, fmt=None, bold=False):
            c = ws.cell(row=row, column=col, value=val)
            c.font      = Font(name="Arial", size=10, color=fg, bold=bold)
            c.fill      = _fill(bg)
            c.alignment = Alignment(horizontal="left" if col == 1 else "right",
                                    vertical="center")
            c.border    = _border()
            if fmt:
                c.number_format = fmt
            return c

        _sc(1, name)
        _sc(2, purchase, "€#,##0.00")
        _sc(3, sales,    "€#,##0.00")
        _sc(4, margin,   "€#,##0.00")
        ws.row_dimensions[row].height = 18
        total_purchase += purchase
        total_sales    += sales
        row += 1

    # Total row — summary has 4 columns, no merge needed
    for col, (val, align) in enumerate([
        ("TOTAL ALL CUSTOMERS", "right"),
        (total_purchase, "right"),
        (total_sales,    "right"),
        (total_sales - total_purchase, "right"),
    ], 1):
        c = ws.cell(row=row, column=col, value=val)
        c.font      = Font(name="Arial", bold=True, color="FFFFFF", size=11)
        c.fill      = _fill(DARK_BLUE)
        c.alignment = Alignment(horizontal=align, vertical="center")
        c.border    = _border()
        if isinstance(val, float):
            c.number_format = "€#,##0.00"
    ws.row_dimensions[row].height = 26


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    """
    Usage:
        python genereer_licentie_overzicht.py --ingram <ingram.xlsx> --pax8 <pax8.csv>
        python genereer_licentie_overzicht.py --ingram <ingram.xlsx> --pax8 <pax8.csv> --output report.xlsx

    If --ingram / --pax8 are omitted, files are auto-detected from the INGRAM_DIR / PAX8_DIR folders.
    Output defaults to: Licensing_Report_YYYY-MM.xlsx in EXPORT_DIR.
    """
    import argparse
    import logging
    import shutil
    from pathlib import Path

    # ── Paths ────────────────────────────────────────────────────────────────
    # Update EXPORT_DIR to the folder where input files are placed and output is written.
    # The PowerShell launcher (genereer_rapport.ps1) passes --ingram/--pax8/--output
    # explicitly, so EXPORT_DIR is only used when running this script directly.
    SCRIPT_DIR  = Path(__file__).parent.resolve()
    EXPORT_DIR  = Path(r"C:\OneDrive\BraveHub\BraveHub - Finance - Licenses_facturatie_upload")
    IMPORT_DIR  = EXPORT_DIR / "Import"
    INGRAM_DIR  = IMPORT_DIR / "Ingram"
    PAX8_DIR    = IMPORT_DIR / "Pax8"
    ARCHIVE_DIR = EXPORT_DIR / "Archive"
    LOG_DIR     = SCRIPT_DIR / "Log"
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    LOG_FILE    = LOG_DIR / "licensing_report.log"

    # ── Logging ───────────────────────────────────────────────────────────────
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[
            logging.FileHandler(LOG_FILE, encoding="utf-8"),
            logging.StreamHandler(),
        ]
    )
    log = logging.getLogger(__name__)

    # ── Arguments ─────────────────────────────────────────────────────────────
    parser = argparse.ArgumentParser()
    parser.add_argument("--ingram", required=False, help="Ingram billing Excel (.xlsx) — auto-detected if omitted")
    parser.add_argument("--pax8",   required=False, help="Pax8 billing CSV — auto-detected if omitted")
    parser.add_argument("--output", required=False, help="Output file path (default: Licensing_Report_YYYY-MM.xlsx)")
    args = parser.parse_args()

    log.info("========================================")
    log.info("Starting licensing report generation")

    # ── Validate export directory ─────────────────────────────────────────────
    if not EXPORT_DIR.exists():
        log.error(f"Export directory not found: {EXPORT_DIR}")
        log.error("Update the EXPORT_DIR variable in this script.")
        _pause_if_interactive("Press Enter to exit...")
        sys.exit(2)
    log.info("Export directory: OK")

    # ── Create subfolders if missing ──────────────────────────────────────────
    for d in [INGRAM_DIR, PAX8_DIR, ARCHIVE_DIR]:
        d.mkdir(parents=True, exist_ok=True)

    # ── Locate Ingram file ────────────────────────────────────────────────────
    ingram_path = None
    if args.ingram:
        ingram_path = Path(args.ingram)
        if not ingram_path.exists():
            log.error(f"Ingram file not found: {ingram_path}")
            _pause_if_interactive("Press Enter to exit...")
            sys.exit(2)
    else:
        log.info(f"Looking for Ingram file in {INGRAM_DIR} ...")
        ingram_files = list(INGRAM_DIR.glob("*.xlsx"))
        if len(ingram_files) > 1:
            log.error(f"Multiple .xlsx files found in {INGRAM_DIR}:")
            for f in ingram_files:
                log.error(f"  {f.name}")
            log.error("Ensure exactly 1 Ingram file is present.")
            _pause_if_interactive("Press Enter to exit...")
            sys.exit(2)
        if len(ingram_files) == 1:
            ingram_path = ingram_files[0]

    if ingram_path:
        log.info(f"Ingram file: {ingram_path.name}")
    else:
        log.warning("No Ingram input found - continuing without Ingram data.")

    # ── Locate Pax8 file ──────────────────────────────────────────────────────
    pax8_path = None
    if args.pax8:
        pax8_path = Path(args.pax8)
        if not pax8_path.exists():
            log.error(f"Pax8 file not found: {pax8_path}")
            _pause_if_interactive("Press Enter to exit...")
            sys.exit(2)
    else:
        log.info(f"Looking for Pax8 file in {PAX8_DIR} ...")
        pax8_files = list(PAX8_DIR.glob("*.csv"))
        if len(pax8_files) > 1:
            log.error(f"Multiple .csv files found in {PAX8_DIR}:")
            for f in pax8_files:
                log.error(f"  {f.name}")
            log.error("Ensure exactly 1 Pax8 file is present.")
            _pause_if_interactive("Press Enter to exit...")
            sys.exit(2)
        if len(pax8_files) == 1:
            pax8_path = pax8_files[0]

    if pax8_path:
        log.info(f"Pax8 file:    {pax8_path.name}")
    else:
        log.warning("No Pax8 input found - continuing without Pax8 data.")

    if not ingram_path and not pax8_path:
        log.error("No input files found (Ingram/Pax8). Provide at least one source file.")
        _pause_if_interactive("Press Enter to exit...")
        sys.exit(2)

    empty_pax8_az = pd.DataFrame(columns=["company_name", "subscription", "az_category", "cost_total", "subtotal"])
    empty_pax8_lic = pd.DataFrame(columns=["company_name", "acronis_endcustomer", "product", "quantity", "cost", "price", "cost_total", "subtotal"])
    empty_ingram_az = pd.DataFrame(columns=["CUSTOMER_NAME", "az_subscription", "az_category", "RESELLER_DETAIL_TOTAL", "CUSTOMER_DETAIL_TOTAL", "RESELLER_DETAIL_START_DATE", "RESELLER_DETAIL_END_DATE", "CUSTOMER_DETAIL_DESCRIPTION", "CUSTOMER_DETAIL_QTY", "RESELLER_DETAIL_UNIT_PRICE", "CUSTOMER_DETAIL_UNIT_PRICE"])
    empty_ingram_lic = pd.DataFrame(columns=["CUSTOMER_NAME", "product", "CUSTOMER_DETAIL_QTY", "RESELLER_DETAIL_TOTAL", "CUSTOMER_DETAIL_TOTAL", "RESELLER_DETAIL_UNIT_PRICE", "CUSTOMER_DETAIL_UNIT_PRICE"])

    # ── Load Ingram ───────────────────────────────────────────────────────────
    ingram_period = None
    if ingram_path:
        log.info("Loading Ingram data...")
        ingram_az, ingram_lic = load_ingram(str(ingram_path))
        ingram_df = pd.read_excel(ingram_path)
        if "RESELLER_INVOICE_DATE" in ingram_df.columns and not ingram_df.empty:
            ingram_period = str(ingram_df["RESELLER_INVOICE_DATE"].iloc[0])[:7]
            log.info(f"  Ingram period : {ingram_period}")
    else:
        ingram_az, ingram_lic = empty_ingram_az.copy(), empty_ingram_lic.copy()

    # ── Load Pax8 ─────────────────────────────────────────────────────────────
    pax8_period = None
    if pax8_path:
        log.info("Loading Pax8 data...")
        pax8_az, pax8_lic = load_pax8(str(pax8_path))
        pax8_df = pd.read_csv(str(pax8_path), encoding="utf-8-sig")
        if "invoice_date" in pax8_df.columns and not pax8_df.empty:
            pax8_period = str(pax8_df["invoice_date"].iloc[0])[:7]
            log.info(f"  Pax8 period   : {pax8_period}")
    else:
        pax8_az, pax8_lic = empty_pax8_az.copy(), empty_pax8_lic.copy()

    period = pax8_period or ingram_period or pd.Timestamp.today().strftime("%Y-%m")

    # ── Archive subfolder for this period ─────────────────────────────────────
    archive_period = ARCHIVE_DIR / period
    archive_period.mkdir(parents=True, exist_ok=True)

    # ── Generate report ───────────────────────────────────────────────────────
    customers = gather_customers(pax8_az, pax8_lic, ingram_az, ingram_lic)
    log.info(f"Customers: {len(customers)}")

    source_label = "Ingram-Pax8" if (ingram_path and pax8_path) else ("IngramOnly" if ingram_path else "Pax8Only")
    is_preliminary = not (ingram_path and pax8_path)
    name_suffix = "_Voorlopig" if is_preliminary else ""
    output_path = Path(args.output) if args.output else EXPORT_DIR / f"Licensing_Report_{period}_{source_label}{name_suffix}.xlsx"

    wb      = Workbook()
    summary = []

    for customer in customers:
        purchase, sales = write_customer_sheet(
            wb, customer, period, ingram_period,
            pax8_az, pax8_lic, ingram_az, ingram_lic,
            pax8_lic_all=pax8_lic
        )
        summary.append((customer, purchase, sales))
        log.info(f"  {customer:<40} purchase \u20ac{purchase:>9,.2f}  /  sales \u20ac{sales:>9,.2f}")

    write_summary_sheet(wb, summary, period)
    wb.save(str(output_path))
    log.info(f"Report saved: {output_path}")

    # ── Archive input files ───────────────────────────────────────────────────
    if ingram_path is not None and not args.ingram:
        shutil.move(str(ingram_path), str(archive_period / ingram_path.name))
        log.info(f"Ingram archived to: {archive_period}")
    if pax8_path is not None and not args.pax8:
        shutil.move(str(pax8_path), str(archive_period / pax8_path.name))
        log.info(f"Pax8 archived to: {archive_period}")

    log.info("Report generated successfully.")
    log.info("========================================")
    print(f"\nDone \u2192 {output_path}")
    _pause_if_interactive("\nPress Enter to exit...")


if __name__ == "__main__":
    main()
