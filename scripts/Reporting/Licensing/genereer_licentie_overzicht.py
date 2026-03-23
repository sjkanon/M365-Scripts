"""
Gecombineerd Licentie & Azure Kostenrapport
============================================
Leest een Pax8 CSV én een Ingram Excel, combineert alle klanten
en schrijft één Excel met per klant een eigen tabblad.

Gebruik:
    python genereer_licentie_overzicht.py <pax8.csv> <ingram.xlsx> [output.xlsx]

Output (standaard):  Licentie_Overzicht_YYYY-MM.xlsx

Vereisten:  pip install pandas openpyxl
"""

import sys
import re

MONTHS_NL = {
    1:"Januari", 2:"Februari", 3:"Maart",    4:"April",
    5:"Mei",     6:"Juni",     7:"Juli",      8:"Augustus",
    9:"September",10:"Oktober",11:"November", 12:"December"
}

def _verbose_period(period_str: str) -> str:
    """'2025-12' -> 'December 2025'"""
    try:
        y, m = period_str[:7].split("-")
        return f"{MONTHS_NL[int(m)]} {y}"
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

# ── Kleuren ───────────────────────────────────────────────────────────────────
DARK_BLUE   = "1F3864"
MED_BLUE    = "2E75B6"
AZURE_BLUE  = "4472C4"   # Azure-sectieheader
LIC_GREEN   = "375623"   # Licenties-sectieheader (donker groen)
LIC_MED     = "538135"   # Licenties sub-header
ROW_LIGHT   = "EAF2FB"
ROW_WHITE   = "FFFFFF"
CAT_GREY    = "F2F2F2"
INGRAM_TEAL = "1F5C6B"   # Ingram sectie

# ── Stijlhulpfuncties ─────────────────────────────────────────────────────────
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

def _empty_row(ws, row, ncols=7):
    for col in range(1, ncols + 1):
        ws.cell(row=row, column=col).border = Border()
    ws.row_dimensions[row].height = 8


# ── Pax8 inlezen ─────────────────────────────────────────────────────────────

PAX8_AZURE_CATEGORIES = {
    "Virtual Machines":          "Virtual Machines",
    "Virtual Machines Licenses": "Virtual Machines",
    "Storage":                   "Storage",
    "Backup":                    "Backup",
    "Bandwidth":                 "Netwerk & Trafiek",
    "Virtual Network":           "Netwerk & Trafiek",
    "Load Balancer":             "Netwerk & Trafiek",
    "Azure DNS":                 "Netwerk & Trafiek",
    "Log Analytics":             "Monitoring & Beheer",
    "Azure Monitor":             "Monitoring & Beheer",
    "Key Vault":                 "Monitoring & Beheer",
    "Logic Apps":                "Monitoring & Beheer",
    "Azure App Service":         "App Services",
    "SQL Database":              "App Services",
    "Azure DevOps":              "App Services",
}

AZURE_CATEGORY_ORDER = [
    "Virtual Machines", "Storage", "Backup",
    "Netwerk & Trafiek", "Monitoring & Beheer", "App Services", "Overig",
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
    cat_raw = parts[4].strip() if len(parts) > 4 else "Overig"
    return sub, PAX8_AZURE_CATEGORIES.get(cat_raw, "Overig")


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

    licenties = df[~is_azure].copy()
    # Productnaam uit description (alles voor " - " of de hele string)
    licenties["product"] = licenties["description"].str.split(" - ").str[0].str.strip()
    licenties["product"] = licenties["product"].str.replace(r'\s*\[options:.*?\]', '', regex=True).str.strip()
    licenties["product"] = licenties["product"].str.replace(r'^\[Deprecated\]\s*', '', regex=True).str.strip()

    # Acronis: extraheer eindklant uit description
    # Formaat: "Product - Product - Eindklant - qty [options:]"
    def _acronis_endcustomer(row):
        if 'acronis' not in str(row['description']).lower():
            return ''
        desc = re.sub(r'\s*\[options:.*?\]\s*$', '', str(row['description'])).strip()
        parts = [p.strip() for p in desc.split(' - ')]
        # Laatste deel is qty (getal), één daarvoor is eindklant
        for i in range(len(parts)-1, -1, -1):
            try:
                float(parts[i])
                continue
            except ValueError:
                if i > 0 and parts[i] not in parts[:i]:
                    return parts[i]
                break
        return ''
    licenties['acronis_endcustomer'] = licenties.apply(_acronis_endcustomer, axis=1)
    # Vertaal Acronis eindklant-namen naar klantnamen in het rapport
    licenties['acronis_endcustomer'] = licenties['acronis_endcustomer'].replace(ACRONIS_ENDCUSTOMER_ALIASES)

    return azure, licenties


# ── Ingram inlezen ────────────────────────────────────────────────────────────

INGRAM_CATEGORY_MAP = {
    "Virtual Machines":             "Virtual Machines",
    "Virtual Machines Licenses":    "Virtual Machines",
    "Storage":                      "Storage",
    "Backup":                       "Backup",
    "Bandwidth":                    "Netwerk & Trafiek",
    "Virtual Network":              "Netwerk & Trafiek",
    "Load Balancer":                "Netwerk & Trafiek",
    "Azure DNS":                    "Netwerk & Trafiek",
    "VPN Gateway":                  "Netwerk & Trafiek",
    "Network Watcher":              "Netwerk & Trafiek",
    "Log Analytics":                "Monitoring & Beheer",
    "Azure Monitor":                "Monitoring & Beheer",
    "Microsoft Defender for Cloud": "Monitoring & Beheer",
    "Key Vault":                    "Monitoring & Beheer",
    "Logic Apps":                   "Monitoring & Beheer",
    "Azure App Service":            "App Services",
    "SQL Database":                 "App Services",
    "Azure Databricks":             "App Services",
    "Azure DevOps":                 "App Services",
    "Reserved VM Instance":         "Reserved Instances (RI)",
}


def _parse_ingram_azure_desc(desc):
    """
    Twee formaten in Ingram:
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
        return sub, INGRAM_CATEGORY_MAP.get(cat_raw, "Overig")
    return "Unknown", "Overig"


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

    licenties = df[~is_azure].copy()
    return azure, licenties


# ── Alle klanten verzamelen ───────────────────────────────────────────────────

def gather_customers(pax8_az, pax8_lic, ingram_az, ingram_lic):
    """Geeft een set van alle klantnamen terug."""
    customers = set()
    for frame, col in [
        (pax8_az,    "company_name"),
        (pax8_lic,   "company_name"),
        (ingram_az,  "CUSTOMER_NAME"),
        (ingram_lic, "CUSTOMER_NAME"),
    ]:
        customers.update(frame[col].dropna().unique())
    return sorted(customers)


# ── Tabblad schrijven ─────────────────────────────────────────────────────────

# Kolomindeling:
# A: Bron / Subscription / Product
# B: Categorie / Detail
# C: Aantal (#)
# D: Eenheidsprijs inkoop
# E: Eenheidsprijs verkoop
# F: Totaal inkoop
# G: Totaal verkoop

NCOLS = 7

def _col_headers(ws, row):
    headers = [
        ("Omschrijving",        "left"),
        ("Categorie / Detail",  "left"),
        ("Aantal",              "center"),
        ("Prijs Inkoop",        "center"),
        ("Prijs Verkoop",       "center"),
        ("Totaal Inkoop (€)",   "center"),
        ("Totaal Verkoop (€)",  "center"),
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


def _data_row(ws, row, omschrijving, detail, qty, unit_inkoop, unit_verkoop,
              tot_inkoop, tot_verkoop, alt=False):
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

    _dc(1, omschrijving, "left")
    _dc(2, detail, "left")
    _dc(3, qty,          "center", "#,##0")
    _dc(4, unit_inkoop,  "right",  "€#,##0.00")
    _dc(5, unit_verkoop, "right",  "€#,##0.00")
    _dc(6, tot_inkoop,   "right",  "€#,##0.00")
    _dc(7, tot_verkoop,  "right",  "€#,##0.00")
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

    # Tabblad-naam max 31 tekens, geen speciale tekens
    sheet_name = re.sub(r'[\\/*?:\[\]]', '', customer)[:31]
    ws = wb.create_sheet(title=sheet_name)

    # Kolombreedte
    widths = [32, 26, 8, 13, 13, 15, 15]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    row = 1

    # ── Titelrij ──
    _styled(ws, row, 1,
            f"{customer}  —  Licentie & Diensten Overzicht  |  {_verbose_period(period)}",
            bg=DARK_BLUE, size=13, merge_to=NCOLS)
    ws.row_dimensions[row].height = 32
    row += 1

    # ── Kolomkoppen ──
    _col_headers(ws, row)
    row += 1

    grand_inkoop  = 0.0
    grand_verkoop = 0.0

    # ════════════════════════════════════════════════════════
    # SECTIE 1: PAX8 AZURE
    # ════════════════════════════════════════════════════════
    cust_az = pax8_az[pax8_az["company_name"] == customer]
    if not cust_az.empty:
        _section_header(ws, row, "☁  Azure (via Pax8)", AZURE_BLUE)
        row += 1

        subs = cust_az["subscription"].unique().tolist()
        ordered_subs = [s for s in PAX8_SUB_ORDER if s in subs]
        ordered_subs += [s for s in subs if s not in ordered_subs]

        for sub in ordered_subs:
            sub_data    = cust_az[cust_az["subscription"] == sub]
            sub_inkoop  = sub_data["cost_total"].sum()
            sub_verkoop = sub_data["subtotal"].sum()
            sub_label   = PAX8_SUB_LABELS.get(sub, sub)

            _sub_header(ws, row, sub_label, sub_inkoop, sub_verkoop,
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

        grand_inkoop  += cust_az["cost_total"].sum()
        grand_verkoop += cust_az["subtotal"].sum()

        _empty_row(ws, row)
        row += 1

    # ════════════════════════════════════════════════════════
    # SECTIE 2: INGRAM AZURE
    # ════════════════════════════════════════════════════════
    cust_iaz = ingram_az[ingram_az["CUSTOMER_NAME"] == customer]
    if not cust_iaz.empty:
        start = pd.to_datetime(cust_iaz["RESELLER_DETAIL_START_DATE"]).min()
        end   = pd.to_datetime(cust_iaz["RESELLER_DETAIL_END_DATE"]).max()
        az_period = f"{start.strftime('%d/%m/%Y')} – {end.strftime('%d/%m/%Y')}"
        _section_header(ws, row, f"☁  Azure (via Ingram)  —  verbruiksperiode {az_period}", INGRAM_TEAL)
        row += 1

        subs = cust_iaz["az_subscription"].unique().tolist()
        # RI apart als laatste
        ri_data = cust_iaz[cust_iaz["az_category"] == "Reserved Instances (RI)"]
        normal_subs = [s for s in subs if s != "Reserved Instances (RI)"]

        for sub in normal_subs:
            sub_data    = cust_iaz[cust_iaz["az_subscription"] == sub]
            sub_inkoop  = sub_data["RESELLER_DETAIL_TOTAL"].sum()
            sub_verkoop = sub_data["CUSTOMER_DETAIL_TOTAL"].sum()
            sub_label   = sub  # geen mapping beschikbaar voor Ingram UUIDs

            _sub_header(ws, row, sub_label, sub_inkoop, sub_verkoop,
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

        # Reserved Instances als aparte sectie
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

        grand_inkoop  += cust_iaz["RESELLER_DETAIL_TOTAL"].sum()
        grand_verkoop += cust_iaz["CUSTOMER_DETAIL_TOTAL"].sum()

        _empty_row(ws, row)
        row += 1

    # ════════════════════════════════════════════════════════
    # SECTIE 3: PAX8 LICENTIES
    # ════════════════════════════════════════════════════════
    cust_pl = pax8_lic[pax8_lic["company_name"] == customer]
    if not cust_pl.empty:
        _section_header(ws, row, "📋  Licenties (via Pax8)", LIC_GREEN)
        row += 1

        # Groepeer per product
        # Splits Acronis-producten per eindklant, groepeer de rest normaal
        non_acronis = cust_pl[cust_pl["acronis_endcustomer"] == ""]
        acronis     = cust_pl[cust_pl["acronis_endcustomer"] != ""]

        grp = non_acronis.groupby("product").agg(
            qty=("quantity", "sum"),
            inkoop=("cost_total", "sum"),
            verkoop=("subtotal", "sum"),
            u_ink=("cost", "mean"),
            u_vk=("price", "mean"),
        ).reset_index()

        alt = False
        for _, r in grp.iterrows():
            _data_row(ws, row, r["product"], "", r["qty"],
                      r["u_ink"], r["u_vk"],
                      r["inkoop"], r["verkoop"], alt=alt)
            alt = not alt
            row += 1

        # Acronis: per eindklant een subgroep
        if not acronis.empty:
            for endcustomer, ec_data in acronis.groupby("acronis_endcustomer"):
                # Subheader voor eindklant
                ec_inkoop  = ec_data["cost_total"].sum()
                ec_verkoop = ec_data["subtotal"].sum()
                _sub_header(ws, row, f"Acronis  —  {endcustomer}",
                            ec_inkoop, ec_verkoop, bg="4A4A8A", fg="FFFFFF")
                row += 1
                alt = False
                for _, r in ec_data.iterrows():
                    _data_row(ws, row, r["product"], "", r["quantity"],
                              r["cost"], r["price"],
                              r["cost_total"], r["subtotal"], alt=alt)
                    alt = not alt
                    row += 1

        grand_inkoop  += cust_pl["cost_total"].sum()
        grand_verkoop += cust_pl["subtotal"].sum()

        _empty_row(ws, row)
        row += 1

    # ════════════════════════════════════════════════════════
    # SECTIE 4: INGRAM LICENTIES
    # ════════════════════════════════════════════════════════
    cust_il = ingram_lic[ingram_lic["CUSTOMER_NAME"] == customer]
    if not cust_il.empty:
        _section_header(ws, row, "📋  Licenties (via Ingram)", LIC_MED)
        row += 1

        grp = cust_il.groupby("product").agg(
            qty=("CUSTOMER_DETAIL_QTY", "sum"),
            inkoop=("RESELLER_DETAIL_TOTAL", "sum"),
            verkoop=("CUSTOMER_DETAIL_TOTAL", "sum"),
            u_ink=("RESELLER_DETAIL_UNIT_PRICE", "mean"),
            u_vk=("CUSTOMER_DETAIL_UNIT_PRICE", "mean"),
        ).reset_index()

        alt = False
        for _, r in grp.iterrows():
            _data_row(ws, row, r["product"], "", r["qty"],
                      r["u_ink"], r["u_vk"],
                      r["inkoop"], r["verkoop"], alt=alt)
            alt = not alt
            row += 1

        grand_inkoop  += cust_il["RESELLER_DETAIL_TOTAL"].sum()
        grand_verkoop += cust_il["CUSTOMER_DETAIL_TOTAL"].sum()

        _empty_row(ws, row)
        row += 1

    # ════════════════════════════════════════════════════════
    # SECTIE 5: ACRONIS (customer is end-customer billed via reseller)
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
            grand_inkoop  += acronis_for_cust["cost_total"].sum()
            grand_verkoop += acronis_for_cust["subtotal"].sum()
            _empty_row(ws, row)
            row += 1

    # ════════════════════════════════════════════════════════
    # GRAND TOTAL
    # ════════════════════════════════════════════════════════
    _totaal_row(ws, row, "TOTAAL", grand_inkoop, grand_verkoop)
    row += 1

    return grand_inkoop, grand_verkoop


# ── Samenvattingstabblad ──────────────────────────────────────────────────────

def write_summary_sheet(wb, summary_rows, period):
    ws = wb.active
    ws.title = "Overzicht"

    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 14

    row = 1
    _styled(ws, row, 1,
            f"Licentie & Diensten Overzicht — Alle Klanten  |  {_verbose_period(period)}",
            bg=DARK_BLUE, size=13, merge_to=4)
    ws.row_dimensions[row].height = 32
    row += 1

    for col, txt in enumerate(["Klant", "Totaal Inkoop (€)", "Totaal Verkoop (€)", "Marge (€)"], 1):
        _styled(ws, row, col, txt, bg=MED_BLUE, size=10)
    ws.row_dimensions[row].height = 22
    row += 1

    total_ink = total_vk = 0.0
    for i, (name, ink, vk) in enumerate(summary_rows):
        bg = ROW_LIGHT if i % 2 else ROW_WHITE
        fg = "333333"
        marge = vk - ink

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
        _sc(2, ink,   "€#,##0.00")
        _sc(3, vk,    "€#,##0.00")
        _sc(4, marge, "€#,##0.00")
        ws.row_dimensions[row].height = 18
        total_ink += ink
        total_vk  += vk
        row += 1

    # Totaalrij — summary heeft maar 4 kolommen, geen merge nodig
    for col, (val, align) in enumerate([
        ("TOTAAL ALLE KLANTEN", "right"),
        (total_ink,   "right"),
        (total_vk,    "right"),
        (total_vk - total_ink, "right"),
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
    Gebruik:
        Alleen Ingram (dag 1-3):
            python genereer_licentie_overzicht.py --ingram <ingram.xlsx>

        Ingram + Pax8 (dag 4-5, compleet rapport):
            python genereer_licentie_overzicht.py --ingram <ingram.xlsx> --pax8 <pax8.csv>

    Output: Licentie_Overzicht_YYYY-MM.xlsx (of opgegeven bestandsnaam)
    """
    import argparse
    import logging
    import shutil
    from datetime import datetime
    from pathlib import Path

    # ── Paths — update EXPORT_DIR to match your OneDrive folder ──────────────
    # Example: Path(r"C:\OneDrive\CompanyName\CompanyName - Finance - Licenses")
    SCRIPT_DIR  = Path(__file__).parent.resolve()
    EXPORT_DIR  = Path(r"C:\OneDrive\CompanyName\CompanyName - Finance - Licenses")
    IMPORT_DIR  = EXPORT_DIR / "Import"
    INGRAM_DIR  = IMPORT_DIR / "Ingram"
    PAX8_DIR    = IMPORT_DIR / "Pax8"
    ARCHIVE_DIR = EXPORT_DIR / "Archive"
    LOG_DIR     = SCRIPT_DIR / "Log"
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    LOG_FILE    = LOG_DIR / "licentie_rapport.log"

    # ── Logging instellen ──
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

    # ── Argparse — voor CLI gebruik, anders auto-detect ──
    parser = argparse.ArgumentParser()
    parser.add_argument("--ingram", required=False, help="Ingram billing Excel (optioneel, anders auto-detect)")
    parser.add_argument("--pax8",   required=False, help="Pax8 billing CSV (optioneel, anders auto-detect)")
    parser.add_argument("--output", required=False, help="Output bestandsnaam")
    args = parser.parse_args()

    log.info("========================================")
    log.info(f"Start licentie rapport generatie")

    # ── Controleer OneDrive bereikbaarheid ──
    if not EXPORT_DIR.exists():
        log.error(f"OneDrive map niet bereikbaar: {EXPORT_DIR}")
        log.error("Controleer of je ingelogd bent en OneDrive gesynchroniseerd is.")
        input("Druk op Enter om af te sluiten...")
        sys.exit(2)
    log.info("OneDrive map bereikbaar: OK")

    # ── Maak mappen aan ──
    for d in [INGRAM_DIR, PAX8_DIR, ARCHIVE_DIR]:
        d.mkdir(parents=True, exist_ok=True)

    # ── Ingram bestand bepalen ──
    if args.ingram:
        ingram_path = Path(args.ingram)
    else:
        log.info(f"Zoeken naar Ingram bestand in {INGRAM_DIR} ...")
        ingram_files = list(INGRAM_DIR.glob("*.xlsx"))
        if len(ingram_files) == 0:
            log.error(f"Geen Excel bestand gevonden in {INGRAM_DIR}")
            log.error("Plaats het Ingram billing bestand in de Ingram map en probeer opnieuw.")
            input("Druk op Enter om af te sluiten...")
            sys.exit(2)
        if len(ingram_files) > 1:
            log.error(f"Meerdere Excel bestanden gevonden in {INGRAM_DIR}:")
            for f in ingram_files:
                log.error(f"  {f.name}")
            log.error("Zorg dat er precies 1 Ingram bestand aanwezig is.")
            input("Druk op Enter om af te sluiten...")
            sys.exit(2)
        ingram_path = ingram_files[0]
    log.info(f"Ingram bestand: {ingram_path.name}")

    # ── Pax8 bestand bepalen ──
    if args.pax8:
        pax8_path = Path(args.pax8)
    else:
        log.info(f"Zoeken naar Pax8 bestand in {PAX8_DIR} ...")
        pax8_files = list(PAX8_DIR.glob("*.csv"))
        if len(pax8_files) == 0:
            log.error(f"Geen CSV bestand gevonden in {PAX8_DIR}")
            log.error("Plaats het Pax8 factuur CSV bestand in de Pax8 map en probeer opnieuw.")
            input("Druk op Enter om af te sluiten...")
            sys.exit(2)
        if len(pax8_files) > 1:
            log.error(f"Meerdere CSV bestanden gevonden in {PAX8_DIR}:")
            for f in pax8_files:
                log.error(f"  {f.name}")
            log.error("Zorg dat er precies 1 Pax8 bestand aanwezig is.")
            input("Druk op Enter om af te sluiten...")
            sys.exit(2)
        pax8_path = pax8_files[0]
    log.info(f"Pax8 bestand:   {pax8_path.name}")

    # ── Ingram laden ──
    log.info(f"Inlezen Ingram...")
    ingram_az, ingram_lic = load_ingram(str(ingram_path))
    ingram_df = pd.read_excel(ingram_path)
    ingram_period = str(ingram_df["RESELLER_INVOICE_DATE"].iloc[0])[:7]
    log.info(f"  Ingram periode (licenties) : {ingram_period}")

    # ── Pax8 laden ──
    log.info(f"Inlezen Pax8...")
    pax8_az, pax8_lic = load_pax8(str(pax8_path))
    pax8_period = pd.read_csv(str(pax8_path), encoding="utf-8-sig")["invoice_date"].iloc[0][:7]
    log.info(f"  Pax8 periode               : {pax8_period}")
    period = pax8_period

    # ── Maand archief map ──
    archive_maand = ARCHIVE_DIR / period
    archive_maand.mkdir(parents=True, exist_ok=True)

    # ── Rapport aanmaken ──
    customers = gather_customers(pax8_az, pax8_lic, ingram_az, ingram_lic)
    log.info(f"Klanten totaal: {len(customers)}")

    output_path = Path(args.output) if args.output else EXPORT_DIR / f"Licentie_Overzicht_{period}.xlsx"

    wb = Workbook()
    summary = []

    for customer in customers:
        ink, vk = write_customer_sheet(
            wb, customer, period, ingram_period,
            pax8_az, pax8_lic, ingram_az, ingram_lic,
            pax8_lic_all=pax8_lic
        )
        summary.append((customer, ink, vk))
        log.info(f"  [Ingram+Pax8]  {customer:<40} inkoop \u20ac{ink:>9,.2f}  /  verkoop \u20ac{vk:>9,.2f}")

    write_summary_sheet(wb, summary, period)
    wb.save(str(output_path))
    log.info(f"Rapport opgeslagen: {output_path}")

    # ── Archiveren ──
    if not args.ingram:
        shutil.move(str(ingram_path), str(archive_maand / ingram_path.name))
        log.info(f"Ingram gearchiveerd naar: {archive_maand}")
    if not args.pax8:
        shutil.move(str(pax8_path), str(archive_maand / pax8_path.name))
        log.info(f"Pax8 gearchiveerd naar: {archive_maand}")

    log.info("Rapport succesvol aangemaakt.")
    log.info("========================================")
    print(f"\nKlaar \u2192 {output_path}")
    input("\nDruk op Enter om af te sluiten...")


if __name__ == "__main__":
    main()
