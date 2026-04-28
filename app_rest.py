import io
import re
import base64
import calendar
import hashlib
import hmac
from pathlib import Path
from datetime import datetime, date

import streamlit as st
import openpyxl

st.set_page_config(page_title="Confronto bollette vs Illumia", layout="wide")

BASE_DIR = Path(__file__).parent
LOGO_JPG = BASE_DIR / "assets" / "logo_energia_solidale.jpg"
LOGO_PNG = BASE_DIR / "assets" / "logo_energia_solidale.png"
TEMPLATE_XLSX = BASE_DIR / "esempio_confronto_corretto.xlsx"
TARIFFE_BASE = BASE_DIR / "tariffe"
INDICI_XLSX = BASE_DIR / "indici_pun_psv_2025_2026.xlsx"
STATIC_DIR = BASE_DIR / "static"
STATIC_DOWNLOADS_DIR = STATIC_DIR / "downloads"
APP_STATE_VERSION = "2026-04-28-static-download-same-tab-1"

if str(st.query_params.get("debug", "")).lower() in {"1", "true", "si", "yes"}:
    st.title("Energia Solidale")
    st.success("Modalita test: Streamlit sta renderizzando la pagina.")
    st.write(
        "Se vedi questo messaggio, il browser riesce a mostrare Streamlit. "
        "Il problema e in un componente della pagina principale."
    )
    print("APP_DEBUG_RENDER", flush=True)
    st.stop()

# Legacy fallback (solo se non hai ancora tariffe/)
LEGACY_RES_FILES = [
    BASE_DIR / "tariffe_illumia_template.xlsx",
    BASE_DIR / "tariffe_illumia_templatemarzo.xlsx",
]
LEGACY_RES_FOUND = [p for p in LEGACY_RES_FILES if p.exists()]

if not TEMPLATE_XLSX.exists():
    st.error("❌ Manca esempio_confronto_corretto.xlsx nella cartella del progetto.")
    st.stop()

KEYS = [
    "vendita_consumo",
    "vendita_fissa",
    "rete_consumi",
    "rete_fissa",
    "quota_potenza",
    "sconti",
    "ricalcoli",
    "arrotondamenti",
    "accise_iva",
]

def ss(k, v):
    if k not in st.session_state:
        st.session_state[k] = v

def reset_state_on_app_update():
    if str(st.query_params.get("reset", "")).lower() in {"1", "true", "si", "yes"}:
        st.session_state.clear()
        st.query_params.clear()
        st.rerun()
    if st.session_state.get("_app_state_version") == APP_STATE_VERSION:
        return
    keep = {
        "auth_logged_in": st.session_state.get("auth_logged_in", False),
        "auth_email": st.session_state.get("auth_email", ""),
        "auth_name": st.session_state.get("auth_name", ""),
    }
    st.session_state.clear()
    for key, value in keep.items():
        if value:
            st.session_state[key] = value
    st.session_state["_app_state_version"] = APP_STATE_VERSION

reset_state_on_app_update()

# Stato base
ss("segmento", "RESIDENZIALE")  # RESIDENZIALE / BUSINESS
ss("nome_cliente", "")
ss("nome_cliente_input", st.session_state["nome_cliente"] or "Cliente")
ss("commodity", "GAS")          # GAS / EE
if st.session_state["commodity"] not in ("GAS", "EE"):
    st.session_state["commodity"] = "GAS"
ss("segmento_choice", 1 if st.session_state["segmento"] == "RESIDENZIALE" else 2)
ss("commodity_choice", 1 if st.session_state["commodity"] == "GAS" else 2)
ss("prev_segmento", st.session_state["segmento"])
ss("prev_commodity", st.session_state["commodity"])
ss("consumo", 0.0)

# Periodo bolletta (solo informativo/controllo; NON è validità offerta)
ss("bill_start", date.today())
ss("bill_end", date.today())
ss("bill_start_day", st.session_state["bill_start"].day)
ss("bill_start_month", st.session_state["bill_start"].month)
ss("bill_start_year", st.session_state["bill_start"].year)
ss("bill_end_day", st.session_state["bill_end"].day)
ss("bill_end_month", st.session_state["bill_end"].month)
ss("bill_end_year", st.session_state["bill_end"].year)

# Periodicità (per quota fissa scontrino)
ss("billing_divisor", 12)
ss("billing_months", 1)
ss("assume_fixed_is_monthly", True)

BILLING_PERIODS = {
    1: "MENSILE",
    2: "BIMESTRALE",
    3: "TRIMESTRALE",
}

# Bolletta B
for k in KEYS:
    ss(f"b_{k}", 0.0)

# Indici
ss("pun_override", 0.0)
ss("psv_override", 0.0)
ss("indice_month", "")
ss("indice_source", "")
ss("indice_file_path", "")
ss("include_dispbt", True)

# Illumia outputs
ss("c_vendita_consumo", 0.0)
ss("c_vendita_fissa", 0.0)
ss("d_vendita_consumo", 0.0)
ss("d_vendita_fissa", 0.0)
ss("illumia_calculated", False)

# Sconti Illumia
ss("ill_sconto_var", -3.0)
ss("ill_sconto_fix", -3.0)

# Tariffe: upload override
ss("tariffe_uploaded_bytes", None)
ss("offers_loaded", False)

# Offerte selezionate automaticamente dall'app
ss("offer_var_name", "")
ss("offer_fix_name", "")

# Validità offerta (presa dal file tariffe più recente)
ss("offer_valid_from", None)
ss("offer_valid_to", None)
ss("offer_file_path", "")

# UI
ss("export_mode", "ENTRAMBE")
ss("excel_ready", False)


# -----------------------------
# Utils
# -----------------------------
def norm(s: str) -> str:
    return " ".join(str(s).strip().lower().split())

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
    # accetta più formati perché nel tuo file si vedono date tipo 03/11/2026, 04/10/2026 ecc. [1](https://onedrive.live.com/personal/8d36b8086e9d2af7/_layouts/15/doc.aspx?resid=ae62e5e9-da89-4f80-8edb-c1ae21b28205&cid=8d36b8086e9d2af7)
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            pass
    return None

def date_from_parts(prefix):
    try:
        year = int(st.session_state[f"{prefix}_year"])
        month = int(st.session_state[f"{prefix}_month"])
        day = int(st.session_state[f"{prefix}_day"])
    except Exception:
        return date.today()

    year = min(max(year, 2000), 2100)
    month = min(max(month, 1), 12)
    day = min(max(day, 1), calendar.monthrange(year, month)[1])
    return date(year, month, day)

def truthy(v):
    if v is None:
        return None
    if isinstance(v, bool):
        return v
    s = str(v).strip().lower()
    if s in ("true", "1", "yes", "y", "si", "sì"):
        return True
    if s in ("false", "0", "no", "n"):
        return False
    return None

def get_auth_config():
    try:
        return st.secrets.get("auth", {})
    except Exception:
        return {}

def auth_enabled():
    return truthy(get_auth_config().get("enabled", False)) is True

def verify_password(password, stored_hash):
    try:
        algo, iterations, salt, expected = str(stored_hash).split("$", 3)
        if algo != "pbkdf2_sha256":
            return False
        iterations = int(iterations)
        digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), iterations)
        actual = base64.b64encode(digest).decode("ascii")
        return hmac.compare_digest(actual, expected)
    except Exception:
        return False

def logout():
    st.session_state["auth_logged_in"] = False
    st.session_state["auth_email"] = ""
    st.session_state["auth_name"] = ""

def require_authentication():
    if not auth_enabled():
        return

    ss("auth_logged_in", False)
    ss("auth_email", "")
    ss("auth_name", "")
    ss("auth_error", "")

    if st.session_state["auth_logged_in"]:
        with st.sidebar:
            st.caption(f"Accesso: {st.session_state['auth_name'] or st.session_state['auth_email']}")
            if st.button("Esci", key="auth_logout"):
                logout()
                st.rerun()
        return

    auth = get_auth_config()
    users = auth.get("users", {})
    names = auth.get("names", {})
    if not users:
        st.error("Accesso non configurato: aggiungi gli utenti nei Secrets di Streamlit Cloud.")
        st.stop()

    logo_path = LOGO_JPG if LOGO_JPG.exists() else LOGO_PNG
    if logo_path.exists():
        render_logo(logo_path)
    st.subheader("Accedi")
    with st.form("login_form"):
        email = st.text_input("Email").strip().lower()
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Accedi")

    if submitted:
        stored_hash = users.get(email, "")
        if stored_hash and verify_password(password, stored_hash):
            st.session_state["auth_logged_in"] = True
            st.session_state["auth_email"] = email
            st.session_state["auth_name"] = names.get(email, email) if names else email
            st.session_state["auth_error"] = ""
            st.rerun()
        else:
            st.session_state["auth_error"] = "Email o password non valide."

    if st.session_state["auth_error"]:
        st.error(st.session_state["auth_error"])
    st.stop()

def clean_text(v) -> str:
    if v is None:
        return ""
    return str(v).strip()

def billing_months_from_dates(start: date, end: date) -> int:
    if not isinstance(start, date) or not isinstance(end, date):
        return 1
    if end < start:
        start, end = end, start
    months = (end.year - start.year) * 12 + (end.month - start.month) + 1
    return max(1, months)

def billing_label_from_months(months) -> str:
    try:
        months = int(months)
    except Exception:
        months = 1
    return BILLING_PERIODS.get(months, f"{months} MESI")

def billing_divisor_from_months(months) -> float:
    try:
        months = int(months)
    except Exception:
        months = 1
    months = max(1, months)
    return 12.0 / float(months)

def bill_fixed_multiplier() -> int:
    try:
        return max(1, int(st.session_state.get("billing_months", 1)))
    except Exception:
        return 1

def bill_fixed_period_amount(value) -> float:
    return float(value) * float(bill_fixed_multiplier())

def month_key_from_date(d: date) -> str:
    return f"{d.year:04d}-{d.month:02d}"

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

def load_indici_rows(path: Path):
    if not path.exists():
        return []
    try:
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
    except Exception:
        return []

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

def commodity_label(comm: str) -> str:
    return "GAS" if comm == "GAS" else "LUCE"

def safe_nome_cognome(nome_cliente: str) -> str:
    if not nome_cliente:
        return "senza_nome"
    s = re.sub(r'[\\/:*?"<>|]', "", nome_cliente.strip())
    parts = s.split()
    if len(parts) >= 2:
        return f"{parts[0]}_{parts[1]}"
    return parts[0]

def safe_download_filename(name: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", str(name).strip())
    return cleaned or "confronto_illumia.xlsx"

def write_static_download(data: bytes, file_name: str):
    STATIC_DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)
    safe_name = safe_download_filename(file_name)
    stamp = datetime.now().strftime("%Y%m%d%H%M%S")
    stored_name = f"{stamp}_{safe_name}"
    path = STATIC_DOWNLOADS_DIR / stored_name
    path.write_bytes(data)
    return f"app/static/downloads/{stored_name}", safe_name

def badge_ok():
    st.markdown(
        "<div style='background:#1f7a1f;color:white;padding:8px 12px;border-radius:10px;display:inline-block;font-weight:700;'>Dati completi ✅</div>",
        unsafe_allow_html=True,
    )

def badge_missing(msg: str):
    st.markdown(
        "<div style='background:#b3261e;color:white;padding:8px 12px;border-radius:10px;display:inline-block;font-weight:700;'>Dati mancanti ❌</div>"
        f"<div style='margin-top:6px;color:#b3261e;font-weight:600;'>{msg}</div>",
        unsafe_allow_html=True,
    )

# -----------------------------
# Tariffe: trova cartella più recente + carica template fisso
# -----------------------------
def find_latest_offer_folder(base: Path):
    if not base.exists():
        return None

    candidates = []

    # supporta: tariffe/2026-04
    for p in base.glob("*"):
        if p.is_dir() and re.match(r"^\d{4}-\d{2}$", p.name):
            candidates.append((p.name, p))

    # supporta: tariffe/2026/04
    for y in base.glob("*"):
        if y.is_dir() and re.match(r"^\d{4}$", y.name):
            for m in y.glob("*"):
                if m.is_dir() and re.match(r"^\d{2}$", m.name):
                    candidates.append((f"{y.name}-{m.name}", m))

    if not candidates:
        return None

    candidates.sort(key=lambda x: x[0])
    return candidates[-1][1]

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
    if TARIFFE_BASE.exists():
        candidates = [
            p for p in TARIFFE_BASE.rglob("*.xlsx")
            if tariff_matches_segment(p, segmento)
        ]
        if candidates:
            candidates.sort(key=lambda p: (tariff_month_key(p), p.stat().st_mtime, p.name))
            selected = candidates[-1]
            st.session_state["offer_file_path"] = str(selected)
            return selected

    # fallback legacy SOLO residenziale
    if segmento == "RESIDENZIALE" and LEGACY_RES_FOUND:
        st.session_state["offer_file_path"] = str(LEGACY_RES_FOUND[0])
        return LEGACY_RES_FOUND[0]

    return None

def get_file_valid_range(xlsx_path: Path):
    try:
        wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
        if "tariffe" not in wb.sheetnames:
            return None, None
        ws = wb["tariffe"]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return None, None
        header = [str(c).strip().lower() if c is not None else "" for c in rows[0]]
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
    except Exception:
        return None, None

def load_tariffe_from_path(path: Path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    if "tariffe" not in wb.sheetnames:
        return []
    ws = wb["tariffe"]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header = [str(c).strip().lower() if c is not None else "" for c in rows[0]]
    out = []
    for r in rows[1:]:
        if not any(r):
            continue
        d = {}
        for i, key in enumerate(header):
            d[key] = r[i] if i < len(r) else None
        d["commodity"] = clean_text(d.get("commodity")).upper()
        d["offer_type"] = clean_text(d.get("offer_type")).upper()
        d["offer_name"] = clean_text(d.get("offer_name"))
        d["component"] = clean_text(d.get("component")).lower()
        d["value_num"] = parse_number(d.get("value"))
        d["is_sellable"] = d.get("is_sellable")
        out.append(d)
    return out

def load_tariffe_from_bytes(xlsx_bytes: bytes):
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    if "tariffe" not in wb.sheetnames:
        return []
    ws = wb["tariffe"]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header = [str(c).strip().lower() if c is not None else "" for c in rows[0]]
    out = []
    for r in rows[1:]:
        if not any(r):
            continue
        d = {}
        for i, key in enumerate(header):
            d[key] = r[i] if i < len(r) else None
        d["commodity"] = clean_text(d.get("commodity")).upper()
        d["offer_type"] = clean_text(d.get("offer_type")).upper()
        d["offer_name"] = clean_text(d.get("offer_name"))
        d["component"] = clean_text(d.get("component")).lower()
        d["value_num"] = parse_number(d.get("value"))
        d["is_sellable"] = d.get("is_sellable")
        out.append(d)
    return out


# -----------------------------
# Offerta scelta dall'app
# -----------------------------
def select_offer_name(rows, commodity, offer_type, segmento):
    c = commodity.upper()
    o = offer_type.upper()
    subset = [r for r in rows if r.get("commodity") == c and r.get("offer_type") == o]

    if any(r.get("is_sellable") is not None for r in subset):
        subset = [r for r in subset if truthy(r.get("is_sellable")) is True]

    # BUSINESS: se manca is_sellable, per prudenza non vendiamo FISSA
    if segmento == "BUSINESS" and not any(r.get("is_sellable") is not None for r in rows):
        if o == "FISSA":
            return ""

    if not subset:
        return ""

    counts = {}
    for r in subset:
        name = clean_text(r.get("offer_name"))
        if name:
            counts[name] = counts.get(name, 0) + 1
    if not counts:
        return ""

    return sorted(counts.items(), key=lambda x: (-x[1], x[0]))[0][0]

def filter_rows_by_offer(rows, commodity, offer_type, offer_name):
    c = commodity.upper()
    o = offer_type.upper()
    return [r for r in rows if r.get("commodity") == c and r.get("offer_type") == o and clean_text(r.get("offer_name")) == offer_name]


# -----------------------------
# Calcolo Illumia
# -----------------------------
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

def calc_illumia_vendite(rows_offer, commodity, offer, consumo, pun, psv, billing_divisor):
    comm = commodity.upper()
    off = offer.upper()

    def fixed_annua(c, o):
        base = comp_sum(rows_offer, c, o, "ccv_quota_fissa")
        disp = comp_sum(rows_offer, c, o, "dispbt") if st.session_state.get("include_dispbt", True) else 0.0
        return base + disp

    if comm == "EE":
        if off == "VARIABILE":
            fee = comp_sum(rows_offer, "EE", "VARIABILE", "fee_energia")
            ccv_var = comp_sum(rows_offer, "EE", "VARIABILE", "ccv_quota_variabile")
            sbil = comp_sum(rows_offer, "EE", "VARIABILE", "sbilanciamento")
            base = pun * 1.10
            prezzo = base + fee + ccv_var + sbil
            vend_cons = consumo * prezzo
            vend_fix = fixed_annua("EE", "VARIABILE") / float(billing_divisor)
            return float(vend_cons), float(vend_fix)

        prezzo = comp_first(rows_offer, "EE", "FISSA", "prezzo_mono")
        if prezzo == 0.0:
            p1 = comp_first(rows_offer, "EE", "FISSA", "prezzo_f1")
            p23 = comp_first(rows_offer, "EE", "FISSA", "prezzo_f23")
            prezzo = max(p1, p23)
        vend_cons = consumo * prezzo
        vend_fix = fixed_annua("EE", "FISSA") / float(billing_divisor)
        return float(vend_cons), float(vend_fix)

    if off == "VARIABILE":
        fee = comp_sum(rows_offer, "GAS", "VARIABILE", "fee_energia")
        ccv_var = comp_sum(rows_offer, "GAS", "VARIABILE", "ccv_quota_variabile")
        bil = comp_sum(rows_offer, "GAS", "VARIABILE", "bilanciamento")
        prezzo = psv + fee + ccv_var + bil
        vend_cons = consumo * prezzo
        vend_fix = fixed_annua("GAS", "VARIABILE") / float(billing_divisor)
        return float(vend_cons), float(vend_fix)

    prezzo = comp_first(rows_offer, "GAS", "FISSA", "prezzo_fisso")
    vend_cons = consumo * prezzo
    vend_fix = fixed_annua("GAS", "FISSA") / float(billing_divisor)
    return float(vend_cons), float(vend_fix)


# -----------------------------
# Export Excel (template + accise conformi al tuo esempio)
# -----------------------------
def find_row_map(ws):
    rm = {}
    for r in range(1, ws.max_row + 1):
        v = ws[f"A{r}"].value
        if not isinstance(v, str):
            continue
        t = norm(v)
        if t == "vendita consumo":
            rm["vendita_consumo"] = r
        elif t.startswith("vendita fissa") and "luce" in t:
            rm["vendita_fissa_luce"] = r
        elif t.startswith("vendita fissa") and "gas" in t:
            rm["vendita_fissa_gas"] = r
        elif "rete" in t and "consumi" in t:
            rm["rete_consumi"] = r
        elif "rete" in t and "fissa" in t:
            rm["rete_fissa"] = r
        elif "quota potenza" in t:
            rm["quota_potenza"] = r
        elif t == "sconti":
            rm["sconti"] = r
        elif "ricalc" in t:
            rm["ricalcoli"] = r
        elif "arrotond" in t:
            rm["arrotondamenti"] = r
        elif "accise" in t:
            rm["accise_iva"] = r
        elif t == "totale":
            rm["totale"] = r

    if "arrotondamenti" not in rm and "accise_iva" in rm:
        candidate = rm["accise_iva"] - 1
        if candidate > 0:
            rm["arrotondamenti"] = candidate

    return rm

def apply_export_labels(ws, nome_cliente: str):
    labels = {
        1: nome_cliente,
        2: None,
        3: "VOCE",
        4: "Vendita Consumo",
        5: "Rete e oneri di sistema Consumi",
        6: "Vendita Fissa Luce",
        7: "Vendita Fissa  Gas",
        8: "Rete e oneri di sistema Fissa",
        9: "Quota Potenza",
        10: "Sconti",
        11: "Ricalcoli/Partite pregresse",
        12: None,
        13: "Accise e Iva",
        14: "Totale",
    }
    for row, value in labels.items():
        ws[f"A{row}"] = value

def write_export_metadata(ws):
    ws["F1"] = f"Offerta Illumia VARIABILE: {st.session_state['offer_var_name']}"
    ws["F2"] = f"Offerta Illumia FISSA: {st.session_state['offer_fix_name'] if st.session_state['offer_fix_name'] else 'N.D.'}"
    if st.session_state.get("offer_valid_from") and st.session_state.get("offer_valid_to"):
        ws["F3"] = f"Validità offerta: {st.session_state['offer_valid_from']} → {st.session_state['offer_valid_to']}"
    else:
        ws["F3"] = "Validità offerta: N.D. (upload manuale)"
    ws["F4"] = f"File tariffe: {st.session_state.get('offer_file_path','')}"
    ws["F5"] = (
        f"Indice PUN/PSV: {st.session_state.get('indice_month','N.D.')} "
        f"({st.session_state.get('indice_file_path','')})"
    )

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
        "arrotondamenti",
        "accise_iva",
        "totale",
    ]
    missing = [key for key in required if key not in rm]
    if missing:
        raise ValueError("Template Excel incompleto: mancano le righe " + ", ".join(missing))

def apply_accise_formula_conforme(ws, rm, col_letter):
    acc = rm["accise_iva"]
    start = rm["vendita_consumo"]
    end = rm["arrotondamenti"]
    ws[f"{col_letter}{acc}"] = f"=SUM({col_letter}{start}:{col_letter}{end})*B{acc}/SUM(B{start}:B{end})"

def apply_total_formula(ws, rm, col_letter):
    acc = rm["accise_iva"]
    start = rm["vendita_consumo"]
    end = rm["arrotondamenti"]
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
    ws[f"{col}{rm['arrotondamenti']}"] = float(vals["arrotondamenti"])

def fill_column_text(ws, rm, col, text):
    for key, r in rm.items():
        if key == "totale" and text == "N.A.":
            continue
        ws[f"{col}{r}"] = text

def fill_column_na(ws, rm, col):
    fill_column_text(ws, rm, col, "N.A.")

def fill_column_nd(ws, rm, col):
    fill_column_text(ws, rm, col, "N.D.")

def reset_illumia_results():
    st.session_state["illumia_calculated"] = False
    st.session_state["c_vendita_consumo"] = 0.0
    st.session_state["c_vendita_fissa"] = 0.0
    st.session_state["d_vendita_consumo"] = 0.0
    st.session_state["d_vendita_fissa"] = 0.0

def comparison_subtotal(vals, commodity):
    subtotal = (
        float(vals["vendita_consumo"])
        + float(vals["rete_consumi"])
        + float(vals["vendita_fissa"])
        + float(vals["rete_fissa"])
        + float(vals["sconti"])
        + float(vals["ricalcoli"])
        + float(vals["arrotondamenti"])
    )
    if commodity == "EE":
        subtotal += float(vals["quota_potenza"])
    return subtotal

def comparison_total(vals, commodity):
    return comparison_subtotal(vals, commodity) + float(vals["accise_iva"])

def build_comparison_values():
    comm = st.session_state["commodity"]

    b_vals = {k: float(st.session_state[f"b_{k}"]) for k in KEYS}
    b_vals["vendita_fissa"] = bill_fixed_period_amount(b_vals["vendita_fissa"])
    b_vals["rete_fissa"] = bill_fixed_period_amount(b_vals["rete_fissa"])

    c_vals = b_vals.copy()
    d_vals = b_vals.copy()

    c_vals["vendita_consumo"] = float(st.session_state["c_vendita_consumo"])
    c_vals["vendita_fissa"] = float(st.session_state["c_vendita_fissa"])
    c_vals["sconti"] = float(st.session_state["ill_sconto_var"])
    c_vals["ricalcoli"] = 0.0
    c_vals["arrotondamenti"] = 0.0

    d_vals["vendita_consumo"] = float(st.session_state["d_vendita_consumo"])
    d_vals["vendita_fissa"] = float(st.session_state["d_vendita_fissa"])
    d_vals["sconti"] = float(st.session_state["ill_sconto_fix"])
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
        "has_offer_var": bool(st.session_state.get("offer_var_name")),
        "has_offer_fix": bool(st.session_state.get("offer_fix_name")),
    }

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

def format_eur(value):
    if isinstance(value, str):
        return value
    amount = float(value)
    sign = "-" if amount < 0 else ""
    formatted = f"{abs(amount):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{sign}€ {formatted}"

def format_index_value(value):
    try:
        return f"{float(value):.4f}".replace(".", ",")
    except Exception:
        return "0,0000"

def render_logo(path):
    suffix = path.suffix.lower()
    mime = "image/png" if suffix == ".png" else "image/jpeg"
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    st.markdown(
        f"""
        <div style="display:flex;justify-content:center;margin:0.25rem 0 1.5rem;">
            <img src="data:{mime};base64,{encoded}" style="width:300px;max-width:50%;height:auto;" alt="Energia Solidale">
        </div>
        """,
        unsafe_allow_html=True,
    )

def render_date_parts(title, prefix, disabled=False):
    st.markdown(f"**{title}**")
    d_col, m_col, y_col = st.columns([1, 1, 1.2])
    with d_col:
        st.number_input("Giorno", min_value=1, max_value=31, step=1, key=f"{prefix}_day", disabled=disabled)
    with m_col:
        st.number_input("Mese", min_value=1, max_value=12, step=1, key=f"{prefix}_month", disabled=disabled)
    with y_col:
        st.number_input("Anno", min_value=2000, max_value=2100, step=1, key=f"{prefix}_year", disabled=disabled)

def query_value(name):
    value = st.query_params.get(name, "")
    if isinstance(value, list):
        value = value[0] if value else ""
    return str(value).strip()

def query_float(name, default):
    raw = query_value(name)
    return parse_number(raw) if raw else float(default)

def build_comparison_table_rows(values=None):
    values = values or build_comparison_values()
    comm = values["commodity"]
    rows_config = [
        ("vendita_consumo", "Vendita Consumo"),
        ("rete_consumi", "Rete e oneri di sistema Consumi"),
        ("vendita_fissa_luce", "Vendita Fissa Luce"),
        ("vendita_fissa_gas", "Vendita Fissa  Gas"),
        ("rete_fissa", "Rete e oneri di sistema Fissa"),
        ("quota_potenza", "Quota Potenza"),
        ("sconti", "Sconti"),
        ("ricalcoli", "Ricalcoli/Partite pregresse"),
        ("arrotondamenti", "Arrotondamenti"),
        ("accise_iva", "Accise e Iva"),
        ("totale", "Totale"),
    ]
    rows = []
    for key, label in rows_config:
        var_value = (
            comparison_value(values["variabile"], key, comm)
            if values["has_offer_var"]
            else "N.D."
        )
        fix_value = (
            comparison_value(values["fissa"], key, comm)
            if values["has_offer_fix"]
            else "N.D."
        )
        rows.append(
            {
                "VOCE": label,
                "Bolletta": format_eur(comparison_value(values["bolletta"], key, comm)),
                "Illumia Variabile": format_eur(var_value),
                "Illumia Fissa": format_eur(fix_value),
            }
        )
    return rows

def build_markdown_table(rows):
    lines = ["| VOCE | Bolletta | Illumia Variabile | Illumia Fissa |", "|---|---:|---:|---:|"]
    for row in rows:
        lines.append(
            "| "
            + str(row["VOCE"]).replace("|", "/")
            + " | "
            + str(row["Bolletta"]).replace("|", "/")
            + " | "
            + str(row["Illumia Variabile"]).replace("|", "/")
            + " | "
            + str(row["Illumia Fissa"]).replace("|", "/")
            + " |"
        )
    return "\n".join(lines)

def render_comparison_rows(rows):
    voce_w = 32
    value_w = 12
    lines = [
        f"{'VOCE':<{voce_w}} {'BOLLETTA':>{value_w}} {'ILL. VAR.':>{value_w}} {'ILL. FISSA':>{value_w}}",
        "-" * (voce_w + (value_w * 3) + 3),
    ]
    for row in rows:
        voce = str(row["VOCE"])
        if len(voce) > voce_w:
            voce = voce[: voce_w - 1] + "."
        lines.append(
            f"{voce:<{voce_w}} "
            f"{str(row['Bolletta']):>{value_w}} "
            f"{str(row['Illumia Variabile']):>{value_w}} "
            f"{str(row['Illumia Fissa']):>{value_w}}"
        )
    st.code("\n".join(lines), language=None)

def bolletta_is_valid():
    v1 = float(st.session_state["b_vendita_consumo"])
    v2 = float(st.session_state["b_vendita_fissa"])
    v3 = float(st.session_state["b_rete_consumi"])
    v4 = float(st.session_state["b_rete_fissa"])
    return not (v1 == 0.0 and v2 == 0.0 and v3 == 0.0 and v4 == 0.0)

require_authentication()

# -----------------------------
# UI
# -----------------------------
logo_path = LOGO_JPG if LOGO_JPG.exists() else LOGO_PNG
if logo_path.exists():
    render_logo(logo_path)

left, center, right = st.columns([1, 3, 1])
with center:
    st.title("Confronto bollette vs Illumia")

    top_name, top_segment, top_supply = st.columns([2, 1, 1])
    with top_name:
        cliente_param = query_value("cliente")
        if cliente_param:
            st.session_state["nome_cliente"] = cliente_param
            st.session_state["nome_cliente_input"] = cliente_param
        elif not st.session_state["nome_cliente"].strip():
            st.session_state["nome_cliente"] = "Cliente"
            st.session_state["nome_cliente_input"] = "Cliente"
        st.caption("Cliente")
        st.text_input("Nome e Cognome", key="nome_cliente_input", disabled=st.session_state["offers_loaded"])
        st.session_state["nome_cliente"] = st.session_state["nome_cliente_input"].strip() or "Cliente"

    with top_segment:
        st.caption("Segmento: 1 RES / 2 BUS")
        st.number_input("Segmento", min_value=1, max_value=2, step=1, key="segmento_choice", disabled=st.session_state["offers_loaded"])
        st.session_state["segmento"] = "RESIDENZIALE" if int(st.session_state["segmento_choice"]) == 1 else "BUSINESS"

    with top_supply:
        st.caption("Fornitura: 1 GAS / 2 LUCE")
        st.number_input("Fornitura", min_value=1, max_value=2, step=1, key="commodity_choice", disabled=st.session_state["offers_loaded"])
        st.session_state["commodity"] = "GAS" if int(st.session_state["commodity_choice"]) == 1 else "EE"

    if (
        st.session_state["segmento"] != st.session_state["prev_segmento"]
        or st.session_state["commodity"] != st.session_state["prev_commodity"]
    ):
        reset_illumia_results()
        st.session_state["offers_loaded"] = False
        st.session_state["prev_segmento"] = st.session_state["segmento"]
        st.session_state["prev_commodity"] = st.session_state["commodity"]

    st.caption(
        f"Selezione attiva: {st.session_state['segmento']} | "
        f"{'Luce' if st.session_state['commodity'] == 'EE' else 'Gas'}"
    )
    if st.session_state["offers_loaded"]:
        st.info("Dati bolletta bloccati in modalità risultato. Per modificarli, ricarica la pagina e reinserisci i valori prima di caricare tariffe e indici.")
    print(
        "APP_RERUN "
        f"segmento={st.session_state['segmento']} "
        f"commodity={st.session_state['commodity']}",
        flush=True,
    )
    nome_ok = bool(st.session_state["nome_cliente"].strip())

    st.divider()

    # 1) Periodicità
    st.header("1️⃣ Periodicità")
    p1, p2 = st.columns(2)
    with p1:
        render_date_parts("Dal (bolletta)", "bill_start", disabled=st.session_state["offers_loaded"])
    with p2:
        render_date_parts("Al (bolletta)", "bill_end", disabled=st.session_state["offers_loaded"])

    st.session_state["bill_start"] = date_from_parts("bill_start")
    st.session_state["bill_end"] = date_from_parts("bill_end")

    billing_months = billing_months_from_dates(st.session_state["bill_start"], st.session_state["bill_end"])
    st.session_state["billing_months"] = billing_months
    st.session_state["billing_divisor"] = billing_divisor_from_months(billing_months)

    f1, f2 = st.columns([2, 3])
    with f1:
        st.markdown(f"**Periodicità:** {billing_label_from_months(billing_months)}")
        st.caption(
            f"Periodo: {st.session_state['bill_start'].strftime('%d/%m/%Y')} - "
            f"{st.session_state['bill_end'].strftime('%d/%m/%Y')}"
        )
    with f2:
        st.session_state["assume_fixed_is_monthly"] = True
        st.caption(
            "Quote fisse inserite come importo mensile: vendita fissa luce/gas "
            "e rete/oneri fissa sono sempre moltiplicate per i mesi fatturati."
        )

    st.caption(
        f"Mesi fatturati = {billing_months} | "
        f"Divisore periodo = {float(st.session_state['billing_divisor']):g} | "
        f"Moltiplicatore quote fisse bolletta = {bill_fixed_multiplier()}"
    )
    print("APP_SECTION periodicita_ok", flush=True)
    st.divider()

    # 2) Bolletta
    st.subheader("2️⃣ Bolletta (scontrino dell'energia)")

    unit = "Smc" if st.session_state["commodity"] == "GAS" else "kWh"

    if st.session_state["offers_loaded"]:
        fixed_multiplier = bill_fixed_multiplier()
        vendita_fissa_periodo = bill_fixed_period_amount(st.session_state["b_vendita_fissa"])
        rete_fissa_periodo = bill_fixed_period_amount(st.session_state["b_rete_fissa"])
        consumo_fmt = f"{float(st.session_state['consumo']):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        st.text(
            "\n".join(
                [
                    f"Consumo ({unit}): {consumo_fmt}",
                    f"Vendita consumo: {format_eur(st.session_state['b_vendita_consumo'])}",
                    f"Rete/oneri consumi: {format_eur(st.session_state['b_rete_consumi'])}",
                    f"Vendita fissa totale periodo ({fixed_multiplier} mesi): {format_eur(vendita_fissa_periodo)}",
                    f"Rete/oneri fissa totale periodo ({fixed_multiplier} mesi): {format_eur(rete_fissa_periodo)}",
                    f"Quota potenza: {format_eur(st.session_state['b_quota_potenza'])}",
                    f"Sconti: {format_eur(st.session_state['b_sconti'])}",
                    f"Ricalcoli: {format_eur(st.session_state['b_ricalcoli'])}",
                    f"Arrotondamenti: {format_eur(st.session_state['b_arrotondamenti'])}",
                    f"Accise + IVA: {format_eur(st.session_state['b_accise_iva'])}",
                ]
            )
        )

    bill_inputs_disabled = st.session_state["offers_loaded"]
    b1, b2 = st.columns(2)
    with b1:
        st.number_input(f"Consumo ({unit})", min_value=0.0, step=0.01, key="consumo", disabled=bill_inputs_disabled)
        st.number_input("Vendita consumo", step=0.01, key="b_vendita_consumo", disabled=bill_inputs_disabled)
        st.number_input("Rete/oneri consumi", step=0.01, key="b_rete_consumi", disabled=bill_inputs_disabled)
        st.number_input("Vendita fissa mensile (scontrino)", step=0.01, key="b_vendita_fissa", disabled=bill_inputs_disabled)
        st.number_input("Rete/oneri fissa mensile", step=0.01, key="b_rete_fissa", disabled=bill_inputs_disabled)
    with b2:
        st.number_input("Quota potenza (solo luce)", step=0.01, key="b_quota_potenza", disabled=(bill_inputs_disabled or st.session_state["commodity"] == "GAS"))
        st.number_input("Sconti", step=0.01, key="b_sconti", disabled=bill_inputs_disabled)
        st.number_input("Ricalcoli", step=0.01, key="b_ricalcoli", disabled=bill_inputs_disabled)
        st.number_input("Arrotondamenti", step=0.01, key="b_arrotondamenti", disabled=bill_inputs_disabled)
        st.number_input("Accise + IVA", step=0.01, key="b_accise_iva", disabled=bill_inputs_disabled)

    bol_ok = bolletta_is_valid()
    if not bol_ok:
        st.warning("⚠️ Bolletta incompleta: inserisci almeno una voce tra Vendita/Rete.")
    print("APP_SECTION bolletta_ok", flush=True)
    st.divider()

    # 3) Offerta più recente (validità bloccata) + indici + selezione offerta + calcolo
    st.header("3️⃣ Offerta Illumia (più recente) + Indici + Offerta automatica + Calcolo")
    if not st.session_state["offers_loaded"]:
        st.info("Carica tariffe e indici quando hai completato i dati della bolletta.")
        if st.button("Carica tariffe e indici", use_container_width=True):
            st.session_state["offers_loaded"] = True
            st.session_state["excel_ready"] = False
            st.rerun()
        print("APP_SECTION waiting_offer_load", flush=True)
        st.stop()

    print("APP_SECTION loading_offers_start", flush=True)

    tariffe_rows = []
    offer_vf, offer_vt = None, None

    try:
        print("APP_SECTION select_offer_file_start", flush=True)
        offer_file = load_tariffe_file_for_segment(st.session_state["segmento"])
        print(f"APP_SECTION select_offer_file_done file={offer_file}", flush=True)
        if offer_file is None:
            st.error("❌ Nessun file tariffe trovato (né tariffe/ né legacy).")
        else:
            print("APP_SECTION load_tariffe_start", flush=True)
            tariffe_rows = load_tariffe_from_path(offer_file)
            print(f"APP_SECTION load_tariffe_done rows={len(tariffe_rows)}", flush=True)
            print("APP_SECTION valid_range_start", flush=True)
            offer_vf, offer_vt = get_file_valid_range(offer_file)
            print(f"APP_SECTION valid_range_done from={offer_vf} to={offer_vt}", flush=True)
            st.session_state["offer_valid_from"] = offer_vf
            st.session_state["offer_valid_to"] = offer_vt
    except Exception as exc:
        print(f"APP_ERROR tariffe {type(exc).__name__}: {exc}", flush=True)
        st.error(f"Errore nel caricamento tariffe: {exc}")
        tariffe_rows = []
        offer_vf, offer_vt = None, None

    # MOSTRA VALIDITÀ OFFERTA SOLO COME TESTO (NON MODIFICABILE)
    if offer_vf and offer_vt:
        st.markdown(
            f"**Validità offerta Illumia (da file più recente):**  \n"
            f"🗓️ {offer_vf.strftime('%d/%m/%Y')} → {offer_vt.strftime('%d/%m/%Y')}"
        )
        st.caption(f"Fonte: {st.session_state.get('offer_file_path','')}")
        # warning se bolletta fuori validità, ma confronto comunque con offerta più recente
        if not (offer_vf <= st.session_state["bill_start"] <= offer_vt and offer_vf <= st.session_state["bill_end"] <= offer_vt):
            st.warning("⚠️ Il periodo bolletta NON rientra nella validità dell’ultima offerta. Il confronto usa comunque l’offerta più recente.")
    else:
        st.info("Validità offerta: N.D. (override upload)")

    # Indici PUN/PSV automatici dal file indici, sul mese più recente del periodo bolletta.
    try:
        print("APP_SECTION load_indici_start", flush=True)
        indici_rows = load_indici_rows(INDICI_XLSX)
        print(f"APP_SECTION load_indici_done rows={len(indici_rows)}", flush=True)
        selected_indice, indice_source = select_indice_for_bill_period(
            indici_rows,
            st.session_state["bill_start"],
            st.session_state["bill_end"],
        )
        print(f"APP_SECTION select_indice_done month={selected_indice['mese'] if selected_indice else None}", flush=True)
    except Exception as exc:
        print(f"APP_ERROR indici {type(exc).__name__}: {exc}", flush=True)
        st.error(f"Errore nel caricamento indici PUN/PSV: {exc}")
        indici_rows = []
        selected_indice, indice_source = None, "missing"

    if selected_indice:
        st.session_state["pun_override"] = float(selected_indice["pun"])
        st.session_state["psv_override"] = float(selected_indice["psv"])
        st.session_state["indice_month"] = selected_indice["mese"]
        st.session_state["indice_source"] = indice_source
        st.session_state["indice_file_path"] = str(INDICI_XLSX)

        st.markdown(
            f"**Indice PUN/PSV usato:** {selected_indice['mese']}  \n"
            f"Fonte: `{INDICI_XLSX.name}`"
        )
        if indice_source == "before_end":
            st.warning(
                "⚠️ Nessun indice disponibile dentro il periodo bolletta: uso il mese più recente disponibile "
                "non successivo alla fine periodo."
            )
        elif indice_source == "latest_available":
            st.warning("⚠️ Il periodo bolletta è precedente agli indici disponibili: uso l’ultimo mese presente nel file.")
    else:
        st.session_state["pun_override"] = 0.0
        st.session_state["psv_override"] = 0.0
        st.session_state["indice_month"] = ""
        st.session_state["indice_source"] = "missing"
        st.session_state["indice_file_path"] = str(INDICI_XLSX)
        st.error("❌ Nessun indice PUN/PSV trovato nel file indici_pun_psv_2025_2026.xlsx.")

    if st.session_state["commodity"] == "EE":
        st.caption("EE variabile: perdite rete 10% su PUN applicate automaticamente")
        st.markdown(f"**PUN (€/kWh) da file indici:** {format_index_value(st.session_state['pun_override'])}")
        indice_ok = float(st.session_state["pun_override"]) > 0
    else:
        st.markdown(f"**PSV (€/Smc) da file indici:** {format_index_value(st.session_state['psv_override'])}")
        indice_ok = float(st.session_state["psv_override"]) > 0

    st.session_state["ill_sconto_var"] = query_float("sconto_var", st.session_state["ill_sconto_var"])
    st.session_state["ill_sconto_fix"] = query_float("sconto_fissa", st.session_state["ill_sconto_fix"])
    s1, s2 = st.columns(2)
    with s1:
        st.markdown(f"**Sconto Illumia Variabile:** {format_eur(st.session_state['ill_sconto_var'])}")
    with s2:
        st.markdown(f"**Sconto Illumia Fissa:** {format_eur(st.session_state['ill_sconto_fix'])}")
    print("APP_SECTION sconti_ok", flush=True)

    # Offerta automatica dall’app
    offer_var = select_offer_name(tariffe_rows, st.session_state["commodity"], "VARIABILE", st.session_state["segmento"]) if tariffe_rows else ""
    offer_fix = select_offer_name(tariffe_rows, st.session_state["commodity"], "FISSA", st.session_state["segmento"]) if tariffe_rows else ""

    st.session_state["offer_var_name"] = offer_var
    st.session_state["offer_fix_name"] = offer_fix

    if offer_var:
        st.success(f"✅ Offerta VARIABILE selezionata dall’app: {offer_var}")
    else:
        st.info("ℹ️ Offerta VARIABILE non selezionata (assente/non vendibile).")

    if offer_fix:
        st.success(f"✅ Offerta FISSA selezionata dall’app: {offer_fix}")
    else:
        st.info("ℹ️ Offerta FISSA non selezionata (assente/non vendibile).")

    print("APP_SECTION offers_loaded_ok", flush=True)

    missing = []
    if not nome_ok:
        missing.append("Nome cliente")
    if not bol_ok:
        missing.append("Bolletta completa")
    if not indice_ok:
        missing.append("PUN/PSV")
    if not tariffe_rows:
        missing.append("Tariffe")
    if not offer_var and not offer_fix:
        missing.append("Offerta Illumia")

    if missing:
        badge_missing("Mancano: " + ", ".join(missing))
        can_calc = False
    else:
        badge_ok()
        can_calc = True

    if can_calc:
        print("APP_SECTION calc_illumia_start", flush=True)
        comm = st.session_state["commodity"]
        consumo = float(st.session_state["consumo"])
        div = float(st.session_state["billing_divisor"])
        pun = float(st.session_state["pun_override"])
        psv = float(st.session_state["psv_override"])

        if offer_var:
            rows_var = filter_rows_by_offer(tariffe_rows, comm, "VARIABILE", offer_var)
            v_cons, v_fix = calc_illumia_vendite(rows_var, comm, "VARIABILE", consumo, pun, psv, div)
        else:
            v_cons, v_fix = 0.0, 0.0

        if offer_fix:
            rows_fix = filter_rows_by_offer(tariffe_rows, comm, "FISSA", offer_fix)
            f_cons, f_fix = calc_illumia_vendite(rows_fix, comm, "FISSA", consumo, pun, psv, div)
        else:
            f_cons, f_fix = 0.0, 0.0

        st.session_state["c_vendita_consumo"] = v_cons
        st.session_state["c_vendita_fissa"] = v_fix
        st.session_state["d_vendita_consumo"] = f_cons
        st.session_state["d_vendita_fissa"] = f_fix
        st.session_state["illumia_calculated"] = True
        print("APP_SECTION calc_illumia_done", flush=True)
        st.success("✅ Calcolo Illumia aggiornato automaticamente.")
    else:
        st.session_state["illumia_calculated"] = False

    st.divider()

    # 4) Confronto in dashboard
    st.header("4️⃣ Confronto")
    if st.session_state["illumia_calculated"]:
        print("APP_SECTION comparison_render_start", flush=True)
        try:
            comparison_values = build_comparison_values()
            print("APP_SECTION comparison_values_ok", flush=True)
            comparison_rows = build_comparison_table_rows(comparison_values)
            print(f"APP_SECTION comparison_rows_ok rows={len(comparison_rows)}", flush=True)
            offer_var_label = st.session_state.get("offer_var_name") or "N.D."
            offer_fix_label = st.session_state.get("offer_fix_name") or "N.D."
            st.caption(
                f"Offerta variabile: {offer_var_label} | "
                f"Offerta fissa: {offer_fix_label}"
            )
            render_comparison_rows(comparison_rows)
            print("APP_SECTION comparison_render_done", flush=True)
        except Exception as exc:
            print(f"APP_ERROR comparison_render {type(exc).__name__}: {exc}", flush=True)
            st.error(f"Errore nella visualizzazione del confronto: {exc}")
            st.stop()
    else:
        st.info("Completa i dati richiesti per visualizzare il confronto.")

    st.divider()

    # 5) Export Excel
    st.header("5️⃣ Scarica Excel")
    st.session_state["export_mode"] = "ENTRAMBE"
    st.caption("Export: Bolletta + Illumia variabile + Illumia fissa")

    def build_excel_bytes():
        wb = openpyxl.load_workbook(TEMPLATE_XLSX)
        ws = wb["Confronto"]
        apply_export_labels(ws, st.session_state["nome_cliente"])
        rm = find_row_map(ws)
        validate_row_map(rm)

        ws["B1"] = None
        ws["C1"] = float(st.session_state["consumo"])
        write_export_metadata(ws)

        comparison_values = build_comparison_values()
        comm = comparison_values["commodity"]
        b_vals = comparison_values["bolletta"]
        c_vals = comparison_values["variabile"]
        d_vals = comparison_values["fissa"]

        # B
        write_column(ws, rm, "B", b_vals, comm)
        ws[f"B{rm['accise_iva']}"] = b_vals["accise_iva"]
        apply_total_formula(ws, rm, "B")

        mode = st.session_state["export_mode"]
        has_offer_var = comparison_values["has_offer_var"]
        has_offer_fix = comparison_values["has_offer_fix"]
        if mode == "ENTRAMBE":
            if has_offer_var:
                write_column(ws, rm, "C", c_vals, comm)
                apply_accise_formula_conforme(ws, rm, "C")
                apply_total_formula(ws, rm, "C")
            else:
                fill_column_nd(ws, rm, "C")
            if has_offer_fix:
                write_column(ws, rm, "D", d_vals, comm)
                apply_accise_formula_conforme(ws, rm, "D")
                apply_total_formula(ws, rm, "D")
            else:
                fill_column_nd(ws, rm, "D")
        elif mode == "VARIABILE":
            if has_offer_var:
                write_column(ws, rm, "C", c_vals, comm)
                apply_accise_formula_conforme(ws, rm, "C")
                apply_total_formula(ws, rm, "C")
            else:
                fill_column_nd(ws, rm, "C")
            if has_offer_fix:
                fill_column_na(ws, rm, "D")
            else:
                fill_column_nd(ws, rm, "D")
        else:
            if has_offer_var:
                fill_column_na(ws, rm, "C")
            else:
                fill_column_nd(ws, rm, "C")
            if has_offer_fix:
                write_column(ws, rm, "D", d_vals, comm)
                apply_accise_formula_conforme(ws, rm, "D")
                apply_total_formula(ws, rm, "D")
            else:
                fill_column_nd(ws, rm, "D")

        out = io.BytesIO()
        wb.save(out)
        return out.getvalue()

    if can_calc:
        if not st.session_state["excel_ready"]:
            st.info("Il confronto è pronto. Prepara l'Excel solo quando vuoi scaricarlo.")
            if st.button("Prepara Excel", use_container_width=True):
                st.session_state["excel_ready"] = True
            else:
                print("APP_SECTION excel_waiting_user_action", flush=True)
        if st.session_state["excel_ready"]:
            try:
                print("APP_SECTION excel_build_start", flush=True)
                data = build_excel_bytes()
                print("APP_SECTION excel_build_done", flush=True)
            except Exception as exc:
                print(f"APP_ERROR excel_build {type(exc).__name__}: {exc}", flush=True)
                st.error(f"❌ Errore durante la generazione Excel: {exc}")
                st.stop()
            nome_file = safe_nome_cognome(st.session_state["nome_cliente"])
            comm_lab = commodity_label(st.session_state["commodity"])
            mode = st.session_state["export_mode"]
            file_name = f"confronto_illumia_{nome_file}_{comm_lab}_{mode}.xlsx"

            href, download_name = write_static_download(data, file_name)
            st.markdown(
                f"""
                <a href="{href}" download="{download_name}"
                   style="display:block;text-align:center;padding:0.65rem 1rem;border-radius:0.5rem;
                          background:#1597d3;color:white;font-weight:700;text-decoration:none;">
                    Scarica Excel
                </a>
                """,
                unsafe_allow_html=True,
            )
            st.caption("Se Chrome non scarica subito: clic destro su Scarica Excel > Salva link con nome.")
            st.caption(f"Link diretto file: {href}")
            print(f"APP_SECTION excel_static_link_done href={href}", flush=True)
    else:
        st.info("Il file Excel sarà disponibile appena il confronto è completo.")

    st.divider()
    st.caption("Energia Solidale")
    st.caption("Associazione senza scopo di lucro - Chioggia (VE) - info@energiasolidale.org")
    print("APP_SECTION footer_render_done", flush=True)
