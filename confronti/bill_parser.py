import re
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal, InvalidOperation


ITALIAN_MONTHS = {
    "gennaio": 1,
    "febbraio": 2,
    "marzo": 3,
    "aprile": 4,
    "maggio": 5,
    "giugno": 6,
    "luglio": 7,
    "agosto": 8,
    "settembre": 9,
    "ottobre": 10,
    "novembre": 11,
    "dicembre": 12,
}
NUMBER_PATTERN = r"[-+]?(?:\d{1,3}(?:\.\d{3})+|\d+)(?:,\d+)?"
DATE_PATTERN = (
    r"\d{1,2}[/.]\d{1,2}[/.]\d{2,4}"
    r"|\d{1,2}\s+(?:gennaio|febbraio|marzo|aprile|maggio|giugno|luglio|agosto|settembre|ottobre|novembre|dicembre)\s+\d{4}"
)


@dataclass
class ParsedBill:
    values: dict = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)


def _normalize_text(text):
    return (
        str(text or "")
        .replace("©", "'")
        .replace("ʼ", "'")
        .replace("’", "'")
        .replace("–", "-")
        .replace("\xa0", " ")
    )


def _collapsed(text):
    return re.sub(r"\s+", " ", _normalize_text(text)).strip()


def _decimal(value):
    if value in (None, ""):
        return None
    cleaned = str(value).strip().replace(".", "").replace(",", ".")
    try:
        return Decimal(cleaned)
    except InvalidOperation:
        return None


def _float(value):
    parsed = _decimal(value)
    return float(parsed) if parsed is not None else None


def _parse_date(value):
    text = str(value or "").strip().lower()
    for separator in ("/", "."):
        if separator in text:
            parts = text.split(separator)
            if len(parts) == 3:
                day, month, year = [int(part) for part in parts]
                if year < 100:
                    year += 2000
                return date(year, month, day)
    match = re.fullmatch(r"(\d{1,2})\s+([a-z]+)\s+(\d{4})", text)
    if match and match.group(2) in ITALIAN_MONTHS:
        return date(int(match.group(3)), ITALIAN_MONTHS[match.group(2)], int(match.group(1)))
    return None


def _billing_months(start, end):
    if not start or not end:
        return 1
    return max(1, (end.year - start.year) * 12 + end.month - start.month + 1)


def _first_match(text, patterns, flags=re.IGNORECASE):
    for pattern in patterns:
        match = re.search(pattern, text, flags)
        if match:
            return match
    return None


def _first_number(text):
    match = re.search(NUMBER_PATTERN, text)
    return _float(match.group(0)) if match else None


def _last_number(text):
    matches = re.findall(NUMBER_PATTERN, text)
    return _float(matches[-1]) if matches else None


def _amount_from_chunk(chunk):
    currency_totals = re.findall(rf"({NUMBER_PATTERN})\s*€(?!\s*/)", chunk)
    if currency_totals:
        return _float(currency_totals[-1])
    return _last_number(chunk)


def _expense_amount(section, label):
    match = re.search(
        rf"{label}(.*?)(?=di cui spesa|QUOTA|Quota|Spesa totale|Altre partite|Totale ricalcoli|Arrotondamento|Accise e IVA|Totale bolletta|$)",
        section,
        re.IGNORECASE,
    )
    return _amount_from_chunk(match.group(1)) if match else None


def _receipt_section(text):
    collapsed = _collapsed(text)
    for match in re.finditer(r"scontrino dell[' ]energia", collapsed, re.IGNORECASE):
        tail = collapsed[match.start() :]
        if re.search(r"(QUOTA PER CONSUMI|Quota consumi|QUANTIT[AÀ])", tail[:500], re.IGNORECASE):
            return tail
    return collapsed


def _section(text, start_pattern, end_pattern):
    match = re.search(rf"{start_pattern}(.*?)(?={end_pattern}|$)", text, re.IGNORECASE)
    return match.group(1) if match else ""


def _extract_name(text):
    match = _first_match(
        text,
        [
            r"Contratto intestato a:\s*\n\s*([A-ZÀ-ÖØ-Ý][A-ZÀ-ÖØ-Ý' ]{3,})\s*\n",
            r"MERCATO LIBERO\s*\n\s*([A-ZÀ-ÖØ-Ý][A-ZÀ-ÖØ-Ý' ]{3,})\s*\n\s*(?:VIA|VIALE|BORGO|CORSO|PIAZZA)",
            r"(?:^|\n)\s*([A-ZÀ-ÖØ-Ý][A-ZÀ-ÖØ-Ý' ]{3,})\s*\n\s*(?:VIA|VIALE|BORGO|CORSO|PIAZZA)\b",
        ],
        flags=re.IGNORECASE | re.MULTILINE,
    )
    if not match:
        return ""
    return " ".join(match.group(1).split()).title()


def _extract_period(text):
    match = _first_match(
        _collapsed(text),
        [
            rf"Periodo (?:oggetto )?di fatturazione\s*:?\s*(?:dal\s*)?({DATE_PATTERN})\s*(?:al|-)\s*({DATE_PATTERN})",
        ],
    )
    if not match:
        return None, None
    return _parse_date(match.group(1)), _parse_date(match.group(2))


def _extract_commodity(text):
    collapsed = _collapsed(text)
    if re.search(r"\bCodice POD\b.{0,80}IT\d{3}[A-Z0-9]{7,}|IT\d{3}[A-Z]\d{6,}", collapsed, re.IGNORECASE):
        return "EE"
    if re.search(r"\bCodice PDR\b.{0,100}\b\d{13,15}\b|\bPDR\b.{0,100}\b\d{13,15}\b", collapsed, re.IGNORECASE):
        return "GAS"
    if re.search(r"\bGAS NATURALE\b|\bGas\b", collapsed, re.IGNORECASE):
        return "GAS"
    if re.search(r"\benergia elettrica\b|\bLuce\b", collapsed, re.IGNORECASE):
        return "EE"
    return ""


def _extract_pod_pdr(text, commodity):
    collapsed = _collapsed(text)
    if commodity == "EE":
        match = re.search(r"\bPOD\b(?:\s+Potenza Disponibile)?\s*([A-Z]{2}\d[A-Z0-9]{8,})", collapsed, re.IGNORECASE)
        if not match:
            match = re.search(r"IT\d{3}[A-Z]\d{6,}", collapsed, re.IGNORECASE)
            return match.group(0).upper() if match else ""
    else:
        match = None
        for label in re.finditer(r"\bPDR\b", collapsed, re.IGNORECASE):
            match = re.search(r"\b(\d{13,15})\b", collapsed[label.end() : label.end() + 220])
            if match:
                break
    return match.group(1).upper() if match else ""


def _extract_consumption(text):
    collapsed = _collapsed(text)
    match = _first_match(
        collapsed,
        [
            rf"Consumo totale fatturato(?: del periodo)?\s*:?\s*({NUMBER_PATTERN})\s*(?:Smc|kWh)",
            rf"Consumo (?:totale )?del periodo(?: di fatturazione)?\s*:?\s*({NUMBER_PATTERN})\s*(?:Smc|kWh)",
            rf"Spesa totale quota consumi\s*({NUMBER_PATTERN})\s*(?:Smc|kWh)",
        ],
    )
    return _float(match.group(1)) if match else None


def _extract_annual_consumption(text):
    collapsed = _collapsed(text)
    match = _first_match(
        collapsed,
        [
            rf"(?:Il tuo consumo annuo aggiornato|In un anno hai consumato)\s*({NUMBER_PATTERN})\s*(?:Smc|kWh)",
            rf"CONSUMO ANNUO(?:\s+mc)?\s*({NUMBER_PATTERN})",
            rf"Da inizio fornitura hai consumato\s*({NUMBER_PATTERN})\s*(?:Smc|kWh)",
        ],
    )
    return _float(match.group(1)) if match else None


def _extract_bill_tariff_type(text):
    collapsed = _collapsed(text)
    if re.search(r"Tipologia offerta\s*:?\s*(?:Offerta )?a prezzo fisso", collapsed, re.IGNORECASE):
        return "FISSA"
    if re.search(r"Tipologia offerta\s*:?\s*(?:Offerta )?a prezzo variabile", collapsed, re.IGNORECASE):
        return "VARIABILE"
    return ""


def _extract_offer_expiry(text):
    collapsed = _collapsed(text)
    match = _first_match(
        collapsed,
        [
            rf"(?:Data di )?[Ss]cadenza condizioni economiche\s*:?\s*({DATE_PATTERN})",
            rf"Data scadenza offerta\s*:?\s*({DATE_PATTERN})",
            rf"condizioni economiche dell[' ]offerta sono valide fino al\s*({DATE_PATTERN})",
        ],
    )
    return _parse_date(match.group(1)) if match else None


def _extract_receipt_values(text, billing_months):
    receipt = _receipt_section(text)
    consumption = _section(
        receipt,
        r"(?:QUOTA PER CONSUMI|Quota consumi)",
        r"(?:QUOTA FISSA|Quota fissa(?: e quota potenza)?)",
    )
    fixed = _section(
        receipt,
        r"(?:Quota fissa e quota potenza|QUOTA FISSA)",
        r"(?:QUOTA POTENZA|Spesa totale quota potenza|Altre partite|Accise e IVA|Totale bolletta)",
    )
    sale_label = r"di cui spesa per (?:la )?vendita (?:di )?(?:energia elettrica|gas naturale)"
    network_label = r"di cui spesa per (?:la )?rete(?: e| gli)?(?: e| gli)? oneri generali di sistema"
    values = {}
    sale_consumption = _expense_amount(consumption, sale_label)
    network_consumption = _expense_amount(consumption, network_label)
    sale_fixed = _expense_amount(fixed, sale_label)
    network_fixed = _expense_amount(fixed, network_label)
    if sale_consumption is not None:
        values["b_vendita_consumo"] = sale_consumption
    if network_consumption is not None:
        values["b_rete_consumi"] = network_consumption
    if sale_fixed is not None:
        values["b_vendita_fissa"] = sale_fixed / billing_months
    if network_fixed is not None:
        values["b_rete_fissa"] = network_fixed / billing_months

    power_match = re.search(
        r"Spesa totale quota potenza(.*?)(?=di cui spesa|Altre partite|Arrotondamento|Accise e IVA|Totale bolletta|$)",
        receipt,
        re.IGNORECASE,
    )
    if not power_match:
        power_match = re.search(
            r"QUOTA POTENZA(.*?)(?=di cui spesa|Altre partite|Arrotondamento|Accise e IVA|Totale bolletta|$)",
            receipt,
        )
    if power_match:
        values["b_quota_potenza"] = _amount_from_chunk(power_match.group(1))

    for key, pattern in [
        ("b_ricalcoli", rf"Totale ricalcoli\s*({NUMBER_PATTERN})\s*€?"),
        ("b_arrotondamenti", rf"Arrotondamento.*?({NUMBER_PATTERN})\s*€"),
        ("b_accise_iva", rf"Accise e IVA\s*({NUMBER_PATTERN})\s*€?"),
    ]:
        match = re.search(pattern, receipt, re.IGNORECASE)
        if match:
            values[key] = _float(match.group(1))
    return values


def parse_bill_text(text):
    normalized = _normalize_text(text)
    if len(normalized.strip()) < 80:
        return ParsedBill(
            warnings=[
                "Il PDF sembra essere una scansione senza testo leggibile. Inserisci manualmente i valori della bolletta."
            ]
        )

    values = {}
    warnings = []
    start, end = _extract_period(normalized)
    commodity = _extract_commodity(normalized)
    if start:
        values["bill_start"] = start
    if end:
        values["bill_end"] = end
    if commodity:
        values["commodity"] = commodity

    name = _extract_name(normalized)
    if name:
        values["nome_cliente"] = name
    pod_pdr = _extract_pod_pdr(normalized, commodity)
    if pod_pdr:
        values["pod_pdr"] = pod_pdr
    consumo = _extract_consumption(normalized)
    if consumo is not None:
        values["consumo"] = consumo
    annual = _extract_annual_consumption(normalized)
    if annual is not None:
        values["tax_annual_consumption"] = annual
    power_match = re.search(
        r"Potenza impegnata.{{0,220}}?({})\s*kW".format(NUMBER_PATTERN),
        normalized,
        re.IGNORECASE | re.DOTALL,
    )
    if power_match:
        values["tax_power_kw"] = _float(power_match.group(1))
    tariff_type = _extract_bill_tariff_type(normalized)
    if tariff_type:
        values["bill_tariff_type"] = tariff_type
    expiry = _extract_offer_expiry(normalized)
    if expiry:
        values["bill_offer_expiry"] = expiry
    values.update(_extract_receipt_values(normalized, _billing_months(start, end)))

    if commodity == "GAS":
        values["tax_power_kw"] = 0
        values["b_quota_potenza"] = 0
    for field_name in ("b_sconti", "b_ricalcoli", "b_bonus_sociale", "b_arrotondamenti", "b_servizi_accessori"):
        values.setdefault(field_name, 0)
    if not expiry:
        warnings.append("Data fine offerta bolletta non riconosciuta: compilala manualmente.")
    if not name:
        warnings.append("Nome cliente non riconosciuto: compilalo manualmente.")
    if not start or not end:
        warnings.append("Periodo di fatturazione non riconosciuto: compilalo manualmente.")
    if consumo is None:
        warnings.append("Consumo della bolletta non riconosciuto: compilalo manualmente.")
    if annual is None:
        warnings.append("Consumo annuo non riconosciuto: compilalo manualmente.")
    if not tariff_type:
        warnings.append("Tipo tariffa bolletta non riconosciuto: seleziona manualmente fissa o variabile.")
    if commodity == "EE" and "tax_power_kw" not in values:
        warnings.append("Potenza impegnata non riconosciuta: compilala manualmente.")
    return ParsedBill(values=values, warnings=warnings)


def parse_uploaded_bill(uploaded_file):
    try:
        from pypdf import PdfReader
    except ImportError as exc:
        raise RuntimeError("La lettura dei PDF non è disponibile: installa la dipendenza pypdf.") from exc
    uploaded_file.seek(0)
    reader = PdfReader(uploaded_file)
    if len(reader.pages) > 60:
        raise ValueError("Il PDF contiene troppe pagine.")
    text = "\n".join(page.extract_text() or "" for page in reader.pages)
    uploaded_file.seek(0)
    return parse_bill_text(text)
