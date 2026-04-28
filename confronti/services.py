import io
import re
from datetime import date, datetime
from pathlib import Path

import openpyxl
from django.conf import settings


BASE_DIR = Path(settings.BASE_DIR)
TEMPLATE_XLSX = BASE_DIR / "esempio_confronto_corretto.xlsx"
TARIFFE_BASE = BASE_DIR / "tariffe"
INDICI_XLSX = BASE_DIR / "indici_pun_psv_2025_2026.xlsx"

KEYS = [
    "vendita_consumo",
    "vendita_fissa",
    "rete_consumi",
    "rete_fissa",
    "quota_potenza",
    "sconti",
    "ricalcoli",
    "bonus_sociale",
    "arrotondamenti",
    "accise_iva",
]

MONTH_NAMES_IT = {
    1: "Gennaio",
    2: "Febbraio",
    3: "Marzo",
    4: "Aprile",
    5: "Maggio",
    6: "Giugno",
    7: "Luglio",
    8: "Agosto",
    9: "Settembre",
    10: "Ottobre",
    11: "Novembre",
    12: "Dicembre",
}


def clean_text(v) -> str:
    return "" if v is None else str(v).strip()


def safe_download_filename(name: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", str(name).strip())
    return cleaned or "confronto_illumia.xlsx"


def parse_number(x) -> float:
    if x is None:
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip().replace("€", "").replace("\u20ac", "").replace(" ", "")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


def parse_date_any(x):
    if x is None:
        return None
    if isinstance(x, datetime):
        return x.date()
    if isinstance(x, date):
        return x
    s = str(x).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            pass
    return None


def month_key_from_date(d: date) -> str:
    return f"{d.year:04d}-{d.month:02d}"


def billing_months_from_dates(start: date, end: date) -> int:
    if end < start:
        start, end = end, start
    months = (end.year - start.year) * 12 + (end.month - start.month) + 1
    return max(1, months)


def billing_divisor_from_months(months) -> float:
    months = max(1, int(months))
    return 12.0 / float(months)


def billing_label_from_months(months) -> str:
    labels = {1: "MENSILE", 2: "BIMESTRALE", 3: "TRIMESTRALE"}
    return labels.get(int(months), f"{int(months)} MESI")


def month_year_label(d: date) -> str:
    return f"{MONTH_NAMES_IT[d.month]} {d.year}"


def bill_period_label(start: date, end: date) -> str:
    if end < start:
        start, end = end, start
    if start.year == end.year and start.month == end.month:
        return month_year_label(start)
    return f"{month_year_label(start)} - {month_year_label(end)}"


def parse_index_month(x):
    if x is None:
        return ""
    if isinstance(x, datetime):
        return month_key_from_date(x.date())
    if isinstance(x, date):
        return month_key_from_date(x)
    s = str(x).strip()
    m = re.search(r"(\d{4})[-/](\d{1,2})", s)
    if m:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}"
    parsed = parse_date_any(s)
    return month_key_from_date(parsed) if parsed else ""


def load_indici_rows(path: Path = INDICI_XLSX):
    if not path.exists():
        return []
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    if "indici" not in wb.sheetnames:
        return []
    ws = wb["indici"]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header = [clean_text(c).lower() for c in rows[0]]
    if not {"mese", "pun", "psv"}.issubset(set(header)):
        return []
    i_mese = header.index("mese")
    i_pun = header.index("pun")
    i_psv = header.index("psv")
    out = []
    for r in rows[1:]:
        if not r or not any(r):
            continue
        mese = parse_index_month(r[i_mese] if i_mese < len(r) else None)
        if not mese:
            continue
        out.append(
            {
                "mese": mese,
                "pun": parse_number(r[i_pun] if i_pun < len(r) else None),
                "psv": parse_number(r[i_psv] if i_psv < len(r) else None),
            }
        )
    return sorted(out, key=lambda row: row["mese"])


def select_indice_for_bill_period(rows, bill_start: date, bill_end: date):
    if not rows:
        return None, "missing"
    if bill_end < bill_start:
        bill_start, bill_end = bill_end, bill_start
    start_month = month_key_from_date(bill_start)
    end_month = month_key_from_date(bill_end)
    in_period = [r for r in rows if start_month <= r["mese"] <= end_month]
    if in_period:
        return in_period[-1], "period"
    before_end = [r for r in rows if r["mese"] <= end_month]
    if before_end:
        return before_end[-1], "before_end"
    return rows[-1], "latest_available"


def tariff_month_key(path: Path) -> str:
    try:
        parts = path.relative_to(TARIFFE_BASE).parts
    except ValueError:
        parts = path.parts
    for i, part in enumerate(parts):
        if re.match(r"^\d{4}-\d{2}$", part):
            return part
        if re.match(r"^\d{4}$", part) and i + 1 < len(parts) and re.match(r"^\d{2}$", parts[i + 1]):
            return f"{part}-{parts[i + 1]}"
    return ""


def tariff_matches_segment(path: Path, segmento: str) -> bool:
    name = path.name.lower()
    parts = {p.lower() for p in path.parts}
    if path.suffix.lower() != ".xlsx" or name.startswith("~$"):
        return False
    if segmento == "BUSINESS":
        return "business" in parts or "business" in name
    return ("residenziale" in parts or "template" in name) and "business" not in name


def load_tariffe_file_for_segment(segmento: str):
    if not TARIFFE_BASE.exists():
        return None
    candidates = [p for p in TARIFFE_BASE.rglob("*.xlsx") if tariff_matches_segment(p, segmento)]
    if not candidates:
        return None
    candidates.sort(key=lambda p: (tariff_month_key(p), p.stat().st_mtime, p.name))
    return candidates[-1]


def get_file_valid_range(xlsx_path: Path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
    if "tariffe" not in wb.sheetnames:
        return None, None
    ws = wb["tariffe"]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return None, None
    header = [clean_text(c).lower() for c in rows[0]]
    if "valid_from" not in header or "valid_to" not in header:
        return None, None
    i_from = header.index("valid_from")
    i_to = header.index("valid_to")
    dfs, dts = [], []
    for r in rows[1:]:
        if not r:
            continue
        vf = parse_date_any(r[i_from] if i_from < len(r) else None)
        vt = parse_date_any(r[i_to] if i_to < len(r) else None)
        if vf:
            dfs.append(vf)
        if vt:
            dts.append(vt)
    if not dfs or not dts:
        return None, None
    return min(dfs), max(dts)


def load_tariffe_from_path(path: Path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    if "tariffe" not in wb.sheetnames:
        return []
    ws = wb["tariffe"]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header = [clean_text(c).lower() for c in rows[0]]
    out = []
    for r in rows[1:]:
        if not any(r):
            continue
        d = {key: (r[i] if i < len(r) else None) for i, key in enumerate(header)}
        d["commodity"] = clean_text(d.get("commodity")).upper()
        d["offer_type"] = clean_text(d.get("offer_type")).upper()
        d["offer_name"] = clean_text(d.get("offer_name"))
        d["component"] = clean_text(d.get("component")).lower()
        d["value_num"] = parse_number(d.get("value"))
        out.append(d)
    return out


def filter_rows_by_offer(rows, commodity, offer_type, offer_name):
    c = commodity.upper()
    o = offer_type.upper()
    return [r for r in rows if r.get("commodity") == c and r.get("offer_type") == o and clean_text(r.get("offer_name")) == offer_name]


def select_offer_name(rows, commodity, offer_type, segmento):
    c = commodity.upper()
    o = offer_type.upper()
    subset = [r for r in rows if r.get("commodity") == c and r.get("offer_type") == o]
    if segmento == "BUSINESS" and o == "FISSA":
        return ""
    counts = {}
    for r in subset:
        name = clean_text(r.get("offer_name"))
        if name:
            counts[name] = counts.get(name, 0) + 1
    if not counts:
        return ""
    return sorted(counts.items(), key=lambda x: (-x[1], x[0]))[0][0]


def comp_sum(rows, commodity, offer, component):
    c = commodity.upper()
    o = offer.upper()
    comp = component.lower()
    return float(sum(r.get("value_num", 0.0) for r in rows if r.get("commodity") == c and r.get("offer_type") == o and r.get("component") == comp))


def comp_first(rows, commodity, offer, component):
    c = commodity.upper()
    o = offer.upper()
    comp = component.lower()
    for r in rows:
        if r.get("commodity") == c and r.get("offer_type") == o and r.get("component") == comp:
            return float(r.get("value_num", 0.0))
    return 0.0


def calc_illumia_vendite(rows_offer, commodity, offer, consumo, pun, psv, billing_divisor, include_dispbt=True):
    comm = commodity.upper()
    off = offer.upper()

    def fixed_annua(c, o):
        base = comp_sum(rows_offer, c, o, "ccv_quota_fissa")
        disp = comp_sum(rows_offer, c, o, "dispbt") if include_dispbt else 0.0
        return base + disp

    if comm == "EE":
        if off == "VARIABILE":
            fee = comp_sum(rows_offer, "EE", "VARIABILE", "fee_energia")
            ccv_var = comp_sum(rows_offer, "EE", "VARIABILE", "ccv_quota_variabile")
            sbil = comp_sum(rows_offer, "EE", "VARIABILE", "sbilanciamento")
            prezzo = (pun * 1.10) + fee + ccv_var + sbil
            return float(consumo * prezzo), float(fixed_annua("EE", "VARIABILE") / float(billing_divisor))
        prezzo = comp_first(rows_offer, "EE", "FISSA", "prezzo_mono")
        if prezzo == 0.0:
            prezzo = max(comp_first(rows_offer, "EE", "FISSA", "prezzo_f1"), comp_first(rows_offer, "EE", "FISSA", "prezzo_f23"))
        return float(consumo * prezzo), float(fixed_annua("EE", "FISSA") / float(billing_divisor))

    if off == "VARIABILE":
        fee = comp_sum(rows_offer, "GAS", "VARIABILE", "fee_energia")
        ccv_var = comp_sum(rows_offer, "GAS", "VARIABILE", "ccv_quota_variabile")
        bil = comp_sum(rows_offer, "GAS", "VARIABILE", "bilanciamento")
        prezzo = psv + fee + ccv_var + bil
        return float(consumo * prezzo), float(fixed_annua("GAS", "VARIABILE") / float(billing_divisor))
    prezzo = comp_first(rows_offer, "GAS", "FISSA", "prezzo_fisso")
    return float(consumo * prezzo), float(fixed_annua("GAS", "FISSA") / float(billing_divisor))


def format_eur(value):
    if isinstance(value, str):
        return value
    amount = float(value)
    sign = "-" if amount < 0 else ""
    formatted = f"{abs(amount):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{sign}€ {formatted}"


def comparison_subtotal(vals, commodity):
    subtotal = (
        float(vals["vendita_consumo"])
        + float(vals["rete_consumi"])
        + float(vals["vendita_fissa"])
        + float(vals["rete_fissa"])
        + float(vals["sconti"])
        + float(vals["ricalcoli"])
        + float(vals.get("bonus_sociale", 0.0))
        + float(vals["arrotondamenti"])
    )
    if commodity == "EE":
        subtotal += float(vals["quota_potenza"])
    return subtotal


def comparison_total(vals, commodity):
    return comparison_subtotal(vals, commodity) + float(vals["accise_iva"])


def comparison_value(vals, key, commodity):
    if key == "vendita_fissa_luce":
        return float(vals["vendita_fissa"]) if commodity == "EE" else "N.A."
    if key == "vendita_fissa_gas":
        return float(vals["vendita_fissa"]) if commodity == "GAS" else "N.A."
    if key == "quota_potenza":
        return float(vals["quota_potenza"]) if commodity == "EE" else "N.A."
    if key == "totale":
        return comparison_total(vals, commodity)
    return float(vals[key])


def build_comparison_table_rows(values):
    comm = values["commodity"]
    rows_config = [
        ("vendita_consumo", "Vendita Consumo"),
        ("rete_consumi", "Rete e oneri di sistema Consumi"),
        ("vendita_fissa_luce", "Vendita Fissa Luce"),
        ("vendita_fissa_gas", "Vendita Fissa Gas"),
        ("rete_fissa", "Rete e oneri di sistema Fissa"),
        ("quota_potenza", "Quota Potenza"),
        ("sconti", "Sconti"),
        ("ricalcoli", "Ricalcoli/Partite pregresse"),
        ("bonus_sociale", "Bonus Sociale"),
        ("arrotondamenti", "Arrotondamenti"),
        ("accise_iva", "Accise e Iva"),
        ("totale", "Totale"),
    ]
    out = []
    for key, label in rows_config:
        var_value = comparison_value(values["variabile"], key, comm) if values["has_offer_var"] else "N.D."
        fix_value = comparison_value(values["fissa"], key, comm) if values["has_offer_fix"] else "N.D."
        out.append(
            {
                "voce": label,
                "bolletta": format_eur(comparison_value(values["bolletta"], key, comm)),
                "variabile": format_eur(var_value),
                "fissa": format_eur(fix_value),
            }
        )
    return out


def build_comparison_values(data, calc):
    comm = data["commodity"]
    months = calc["billing_months"]
    b_vals = {k: float(data.get(f"b_{k}", 0.0)) for k in KEYS}
    b_vals["bonus_sociale"] = -abs(float(b_vals.get("bonus_sociale", 0.0)))
    b_vals["vendita_fissa"] *= months
    b_vals["rete_fissa"] *= months

    c_vals = b_vals.copy()
    d_vals = b_vals.copy()
    c_vals["vendita_consumo"] = calc["v_cons"]
    c_vals["vendita_fissa"] = calc["v_fix"]
    c_vals["sconti"] = float(data.get("ill_sconto_var", -3.0))
    c_vals["ricalcoli"] = 0.0
    c_vals["arrotondamenti"] = 0.0
    d_vals["vendita_consumo"] = calc["f_cons"]
    d_vals["vendita_fissa"] = calc["f_fix"]
    d_vals["sconti"] = float(data.get("ill_sconto_fix", -3.0))
    d_vals["ricalcoli"] = 0.0
    d_vals["arrotondamenti"] = 0.0

    base_subtotal = comparison_subtotal(b_vals, comm)
    tax_rate = float(b_vals["accise_iva"]) / base_subtotal if base_subtotal else 0.0
    c_vals["accise_iva"] = tax_rate * comparison_subtotal(c_vals, comm)
    d_vals["accise_iva"] = tax_rate * comparison_subtotal(d_vals, comm)
    return {
        "commodity": comm,
        "bolletta": b_vals,
        "variabile": c_vals,
        "fissa": d_vals,
        "has_offer_var": bool(calc["offer_var"]),
        "has_offer_fix": bool(calc["offer_fix"]),
    }


def prepare_comparison(data):
    bill_start = data["bill_start"]
    bill_end = data["bill_end"]
    segmento = data["segmento"]
    commodity = data["commodity"]
    consumo = float(data["consumo"])
    billing_months = billing_months_from_dates(bill_start, bill_end)
    billing_divisor = billing_divisor_from_months(billing_months)

    offer_file = load_tariffe_file_for_segment(segmento)
    tariffe_rows = load_tariffe_from_path(offer_file) if offer_file else []
    offer_valid_from, offer_valid_to = get_file_valid_range(offer_file) if offer_file else (None, None)
    offer_var = select_offer_name(tariffe_rows, commodity, "VARIABILE", segmento)
    offer_fix = select_offer_name(tariffe_rows, commodity, "FISSA", segmento)

    indici_rows = load_indici_rows(INDICI_XLSX)
    indice, indice_reason = select_indice_for_bill_period(indici_rows, bill_start, bill_end)
    pun = float(indice["pun"]) if indice else 0.0
    psv = float(indice["psv"]) if indice else 0.0

    if offer_var:
        rows_var = filter_rows_by_offer(tariffe_rows, commodity, "VARIABILE", offer_var)
        v_cons, v_fix = calc_illumia_vendite(rows_var, commodity, "VARIABILE", consumo, pun, psv, billing_divisor)
    else:
        v_cons, v_fix = 0.0, 0.0

    if offer_fix:
        rows_fix = filter_rows_by_offer(tariffe_rows, commodity, "FISSA", offer_fix)
        f_cons, f_fix = calc_illumia_vendite(rows_fix, commodity, "FISSA", consumo, pun, psv, billing_divisor)
    else:
        f_cons, f_fix = 0.0, 0.0

    calc = {
        "billing_months": billing_months,
        "billing_divisor": billing_divisor,
        "billing_label": billing_label_from_months(billing_months),
        "period_label": bill_period_label(bill_start, bill_end),
        "offer_file": str(offer_file) if offer_file else "",
        "offer_valid_from": offer_valid_from,
        "offer_valid_to": offer_valid_to,
        "offer_var": offer_var,
        "offer_fix": offer_fix,
        "indice": indice,
        "indice_reason": indice_reason,
        "v_cons": v_cons,
        "v_fix": v_fix,
        "f_cons": f_cons,
        "f_fix": f_fix,
    }
    values = build_comparison_values(data, calc)
    return {"calc": calc, "values": values, "rows": build_comparison_table_rows(values)}


def find_row_map(ws):
    rm = {}
    for r in range(1, ws.max_row + 1):
        t = clean_text(ws[f"A{r}"].value).lower()
        if not t:
            continue
        if "vendita consumo" in t:
            rm["vendita_consumo"] = r
        elif "vendita fissa luce" in t:
            rm["vendita_fissa_luce"] = r
        elif "vendita fissa" in t and "gas" in t:
            rm["vendita_fissa_gas"] = r
        elif "rete" in t and "consumi" in t:
            rm["rete_consumi"] = r
        elif "rete" in t and "fissa" in t:
            rm["rete_fissa"] = r
        elif "quota potenza" in t:
            rm["quota_potenza"] = r
        elif "sconti" in t:
            rm["sconti"] = r
        elif "ricalcoli" in t or "partite" in t:
            rm["ricalcoli"] = r
        elif "bonus social" in t:
            rm["bonus_sociale"] = r
        elif "arrotondamenti" in t:
            rm["arrotondamenti"] = r
        elif "accise" in t or "iva" in t:
            rm["accise_iva"] = r
        elif t == "totale":
            rm["totale"] = r
    if "arrotondamenti" not in rm and "bonus_sociale" not in rm and "accise_iva" in rm:
        candidate = rm["accise_iva"] - 1
        if candidate > 0:
            rm["arrotondamenti"] = candidate
    return rm


def validate_row_map(rm):
    required = [
        "vendita_consumo",
        "vendita_fissa_luce",
        "vendita_fissa_gas",
        "rete_consumi",
        "rete_fissa",
        "quota_potenza",
        "sconti",
        "ricalcoli",
        "bonus_sociale",
        "accise_iva",
        "totale",
    ]
    missing = [key for key in required if key not in rm]
    if missing:
        raise ValueError("Template Excel incompleto: mancano le righe " + ", ".join(missing))


def apply_export_labels(ws, nome_cliente: str):
    labels = {
        1: nome_cliente or "Cliente",
        3: "VOCE",
        4: "Vendita Consumo",
        5: "Rete e oneri di sistema Consumi",
        6: "Vendita Fissa Luce",
        7: "Vendita Fissa Gas",
        8: "Rete e oneri di sistema Fissa",
        9: "Quota Potenza",
        10: "Sconti",
        11: "Ricalcoli/Partite pregresse",
        12: "Bonus Sociale",
        13: "Accise e Iva",
        14: "Totale",
    }
    for row, value in labels.items():
        ws[f"A{row}"] = value


def write_export_metadata(ws, prepared):
    calc = prepared["calc"]
    ws["F1"] = f"Offerta Illumia VARIABILE: {calc['offer_var'] or 'N.D.'}"
    ws["F2"] = f"Offerta Illumia FISSA: {calc['offer_fix'] or 'N.D.'}"
    if calc["offer_valid_from"] and calc["offer_valid_to"]:
        ws["F3"] = f"Validità offerta: {calc['offer_valid_from']} → {calc['offer_valid_to']}"
    else:
        ws["F3"] = "Validità offerta: N.D."
    ws["F4"] = f"File tariffe: {calc['offer_file']}"
    indice = calc["indice"] or {}
    ws["F5"] = f"Indice PUN/PSV: {indice.get('mese', 'N.D.')} ({INDICI_XLSX.name})"


def apply_accise_formula_conforme(ws, rm, col_letter):
    acc = rm["accise_iva"]
    start = rm["vendita_consumo"]
    end = acc - 1
    ws[f"{col_letter}{acc}"] = f"=SUM({col_letter}{start}:{col_letter}{end})*B{acc}/SUM(B{start}:B{end})"


def apply_total_formula(ws, rm, col_letter):
    acc = rm["accise_iva"]
    start = rm["vendita_consumo"]
    end = acc - 1
    ws[f"{col_letter}{rm['totale']}"] = f"=SUM({col_letter}{start}:{col_letter}{end})+{col_letter}{acc}"


def write_column(ws, rm, col, vals, commodity):
    ws[f"{col}{rm['vendita_consumo']}"] = float(vals["vendita_consumo"])
    if commodity == "GAS":
        ws[f"{col}{rm['vendita_fissa_gas']}"] = float(vals["vendita_fissa"])
        ws[f"{col}{rm['vendita_fissa_luce']}"] = "N.A."
        ws[f"{col}{rm['quota_potenza']}"] = "N.A."
    else:
        ws[f"{col}{rm['vendita_fissa_luce']}"] = float(vals["vendita_fissa"])
        ws[f"{col}{rm['vendita_fissa_gas']}"] = "N.A."
        ws[f"{col}{rm['quota_potenza']}"] = float(vals["quota_potenza"])
    ws[f"{col}{rm['rete_consumi']}"] = float(vals["rete_consumi"])
    ws[f"{col}{rm['rete_fissa']}"] = float(vals["rete_fissa"])
    ws[f"{col}{rm['sconti']}"] = float(vals["sconti"])
    ws[f"{col}{rm['ricalcoli']}"] = float(vals["ricalcoli"])
    if "bonus_sociale" in rm:
        ws[f"{col}{rm['bonus_sociale']}"] = float(vals.get("bonus_sociale", 0.0))
    if "arrotondamenti" in rm:
        ws[f"{col}{rm['arrotondamenti']}"] = float(vals["arrotondamenti"])


def fill_column_text(ws, rm, col, text):
    for key, r in rm.items():
        if key == "totale" and text == "N.A.":
            continue
        ws[f"{col}{r}"] = text


def build_excel_bytes(data, prepared=None):
    prepared = prepared or prepare_comparison(data)
    wb = openpyxl.load_workbook(TEMPLATE_XLSX)
    ws = wb["Confronto"]
    apply_export_labels(ws, data.get("nome_cliente", "Cliente"))
    rm = find_row_map(ws)
    validate_row_map(rm)
    ws["B1"] = None
    ws["C1"] = float(data["consumo"])
    write_export_metadata(ws, prepared)

    values = prepared["values"]
    comm = values["commodity"]
    b_vals = values["bolletta"]
    c_vals = values["variabile"]
    d_vals = values["fissa"]

    write_column(ws, rm, "B", b_vals, comm)
    ws[f"B{rm['accise_iva']}"] = b_vals["accise_iva"]
    apply_total_formula(ws, rm, "B")
    if values["has_offer_var"]:
        write_column(ws, rm, "C", c_vals, comm)
        apply_accise_formula_conforme(ws, rm, "C")
        apply_total_formula(ws, rm, "C")
    else:
        fill_column_text(ws, rm, "C", "N.D.")
    if values["has_offer_fix"]:
        write_column(ws, rm, "D", d_vals, comm)
        apply_accise_formula_conforme(ws, rm, "D")
        apply_total_formula(ws, rm, "D")
    else:
        fill_column_text(ws, rm, "D", "N.D.")

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
