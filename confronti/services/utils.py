import re
from datetime import date, datetime
from pathlib import Path

from django.conf import settings
from django.utils import timezone


BASE_DIR = Path(settings.BASE_DIR)
TEMPLATE_XLSX = BASE_DIR / "esempio_confronto_corretto.xlsx"
TARIFFE_BASE = BASE_DIR / "tariffe"
EON_TARIFFE_DIR = BASE_DIR / "estrazioni_tariffe"
INDICI_XLSX = BASE_DIR / "indici_pun_psv_2025_2026.xlsx"
PROVIDERS = {
    "ILLUMIA": "Illumia",
    "EON": "E.ON",
    "CVE": "CVE",
}
SEGMENTS = ("RESIDENZIALE", "MICROBUSINESS", "BUSINESS")
EE_EXCISE_RATE = 0.0227
EE_RESIDENTIAL_EXEMPT_KWH_PER_MONTH = 150.0
GAS_VAT_REDUCED_ANNUAL_THRESHOLD = 480.0
GAS_EXCISE_BRACKETS = (
    (120.0, 0.044),
    (480.0, 0.175),
    (1560.0, 0.170),
    (None, 0.186),
)
GAS_REGIONAL_ADDITIONAL_MIN_RATES = {
    "Abruzzo": 0.02,
    "Basilicata": 0.02,
    "Calabria": 0.02,
    "Campania": 0.02,
    "Emilia-Romagna": 0.02,
    "Friuli-Venezia Giulia": 0.0,
    "Lazio": 0.02,
    "Liguria": 0.02,
    "Lombardia": 0.0,
    "Marche": 0.02,
    "Molise": 0.02,
    "Piemonte": 0.022,
    "Puglia": 0.02,
    "Sardegna": 0.0,
    "Sicilia": 0.0,
    "Toscana": 0.02,
    "Trentino-Alto Adige": 0.0,
    "Umbria": 0.02,
    "Valle d'Aosta": 0.0,
    "Veneto": 0.007747,
}
BILL_TARIFF_TYPE_LABELS = {
    "VARIABILE": "Variabile",
    "FISSA": "Fissa",
}
TARIFF_SELECTION_MODE_LABELS = {
    "LATEST": "Ultime tariffe disponibili",
    "PERIOD": "Tariffe del periodo bolletta",
}
ACCESSORY_SERVICES_VAT_RATES = {
    "22%": 0.22,
    "10%": 0.10,
    "Esente": 0.0,
}
ACCESSORY_SERVICES_VAT_OPTIONS = tuple(ACCESSORY_SERVICES_VAT_RATES)

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
    "servizi_accessori",
    "accise_iva",
]

MONTH_NAMES_IT = {
    1: "Gennaio", 2: "Febbraio", 3: "Marzo", 4: "Aprile",
    5: "Maggio", 6: "Giugno", 7: "Luglio", 8: "Agosto",
    9: "Settembre", 10: "Ottobre", 11: "Novembre", 12: "Dicembre",
}


def clean_text(v):
    return "" if v is None else str(v).strip()


def normalize_provider(value):
    cleaned = clean_text(value or "ILLUMIA").upper().replace(".", "").replace("-", "").replace(" ", "")
    if cleaned == "EON":
        return "EON"
    if cleaned in {"CVE", "CENTROVENETOENERGIE", "CENTROVENETO"}:
        return "CVE"
    return "ILLUMIA"


def normalize_providers(value):
    if value is None:
        return []
    if isinstance(value, str):
        parts = re.split(r"[,;|]+", value.replace("[", "").replace("]", "").replace("'", "").replace('"', ""))
    elif isinstance(value, (list, tuple, set)):
        parts = value
    else:
        parts = [value]
    out = []
    for part in parts:
        if not clean_text(part):
            continue
        provider = normalize_provider(part)
        if provider not in out:
            out.append(provider)
    return out


def provider_list_label(providers):
    normalized = normalize_providers(providers)
    return " + ".join(provider_label(provider) for provider in normalized) if normalized else "N.D."


def provider_label(value):
    return PROVIDERS.get(normalize_provider(value), "Illumia")


def bool_from_data(value):
    if isinstance(value, bool):
        return value
    return clean_text(value).upper() in {"1", "TRUE", "SI", "SÌ", "YES", "ON"}


def normalize_bill_tariff_type(value):
    cleaned = clean_text(value).upper()
    return cleaned if cleaned in BILL_TARIFF_TYPE_LABELS else "VARIABILE"


def bill_tariff_type_label(value):
    return BILL_TARIFF_TYPE_LABELS.get(normalize_bill_tariff_type(value), "Variabile")


def normalize_tariff_selection_mode(value):
    cleaned = clean_text(value).upper()
    return cleaned if cleaned in TARIFF_SELECTION_MODE_LABELS else "LATEST"


def tariff_selection_mode_label(value):
    return TARIFF_SELECTION_MODE_LABELS[normalize_tariff_selection_mode(value)]


def normalize_accessory_services_vat_label(value):
    label = clean_text(value) or "22%"
    return label if label in ACCESSORY_SERVICES_VAT_RATES else "22%"


def normalize_primary_home(value):
    cleaned = clean_text(value).upper()
    return "NO" if cleaned in {"NO", "N", "0", "FALSE"} else "SI"


def primary_home_label(value):
    return "Sì" if normalize_primary_home(value) == "SI" else "No"


def normalize_region(value):
    cleaned = clean_text(value)
    return cleaned if cleaned in GAS_REGIONAL_ADDITIONAL_MIN_RATES else "Veneto"


def accessory_services_vat_rate(value):
    return float(ACCESSORY_SERVICES_VAT_RATES[normalize_accessory_services_vat_label(value)])


def progressive_tax_amount(consumption, brackets):
    consumption = max(0.0, float(consumption))
    total = 0.0
    lower = 0.0
    for upper, rate in brackets:
        if upper is None:
            taxable = max(0.0, consumption - lower)
        else:
            taxable = max(0.0, min(consumption, upper) - lower)
        total += taxable * float(rate)
        if upper is None or consumption <= upper:
            break
        lower = float(upper)
    return total


def annual_progressive_tax_for_period(period_consumption, annual_consumption, brackets):
    period_consumption = max(0.0, float(period_consumption))
    annual_consumption = max(0.0, float(annual_consumption))
    if period_consumption <= 0 or annual_consumption <= 0:
        return 0.0
    annual_tax = progressive_tax_amount(annual_consumption, brackets)
    return annual_tax * (period_consumption / annual_consumption)


def comparison_datetime_from_data(value=None):
    if isinstance(value, datetime):
        dt = value
    elif value:
        try:
            dt = datetime.fromisoformat(str(value))
        except ValueError:
            dt = timezone.localtime()
    else:
        dt = timezone.localtime()
    if timezone.is_naive(dt):
        dt = timezone.make_aware(dt, timezone.get_current_timezone())
    return timezone.localtime(dt)


def comparison_datetime_label(value=None):
    return comparison_datetime_from_data(value).strftime("%d/%m/%Y %H:%M:%S")


def date_label_it(value):
    return value.strftime("%d/%m/%Y") if isinstance(value, date) else "N.D."


def safe_download_filename(name):
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", str(name).strip())
    return cleaned or "confronto_bollette.xlsx"


def parse_number(x):
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


def month_key_from_date(d):
    return f"{d.year:04d}-{d.month:02d}"


def billing_months_from_dates(start, end):
    if end < start:
        start, end = end, start
    months = (end.year - start.year) * 12 + (end.month - start.month) + 1
    return max(1, months)


def billing_divisor_from_months(months):
    months = max(1, int(months))
    return 12.0 / float(months)


def billing_label_from_months(months):
    labels = {1: "MENSILE", 2: "BIMESTRALE", 3: "TRIMESTRALE"}
    return labels.get(int(months), f"{int(months)} MESI")


def month_year_label(d):
    return f"{MONTH_NAMES_IT[d.month]} {d.year}"


def bill_period_label(start, end):
    if end < start:
        start, end = end, start
    if start.year == end.year and start.month == end.month:
        return month_year_label(start)
    return f"{month_year_label(start)} - {month_year_label(end)}"


def bill_period_outside_offer_validity(bill_start, bill_end, offer_valid_from, offer_valid_to):
    if not offer_valid_from or not offer_valid_to:
        return False
    if bill_end < bill_start:
        bill_start, bill_end = bill_end, bill_start
    return bill_start < offer_valid_from or bill_end > offer_valid_to


def format_eur(value):
    if isinstance(value, str):
        return value
    amount = float(value)
    sign = "-" if amount < 0 else ""
    formatted = f"{abs(amount):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{sign}€ {formatted}"


def format_percent(value):
    if value is None:
        return "N.D."
    return f"{float(value) * 100:.2f}".replace(".", ",") + "%"


def tax_incidence_ratio(total_tax, subtotal):
    if subtotal <= 0:
        return None
    return max(0.0, float(total_tax)) / float(subtotal)


def fiscal_unit(commodity):
    return "Smc" if commodity == "GAS" else "kWh"


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
        + float(vals.get("servizi_accessori", 0.0))
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
    if key in {"accise", "iva"}:
        value = vals.get(key)
        return "N.D." if value is None else float(value)
    if key == "totale":
        return comparison_total(vals, commodity)
    return float(vals[key])


def _excel_decimal(value):
    text = f"{float(value):.6f}".rstrip("0").rstrip(".")
    return text or "0"
