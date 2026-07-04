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
ABBREVIATED_ITALIAN_MONTHS = {
    "gen": 1,
    "feb": 2,
    "mar": 3,
    "apr": 4,
    "mag": 5,
    "giu": 6,
    "lug": 7,
    "ago": 8,
    "set": 9,
    "ott": 10,
    "nov": 11,
    "dic": 12,
}
NUMBER_PATTERN = r"[-+]?(?:\d{1,3}(?:\.\d{3})+|\d+)(?:,\d+)?"
DATE_PATTERN = (
    r"(?:\d{1,2}[/.]\d{1,2}[/.]\d{2,4}"
    r"|\d{1,2}\s+(?:gennaio|febbraio|marzo|aprile|maggio|giugno|luglio|agosto|settembre|ottobre|novembre|dicembre)\s+\d{4})"
)
ADDRESS_PREFIX_PATTERN = (
    r"(?:VIA|VIALE|BORGO|CORSO|PIAZZA|P\.ZZA|PZA|STRADA|LOCALIT[AÀ]|LOC\.?|LARGO|"
    r"RIVIERA|FONDAMENTA|CALLE|CAMPO|SALIZADA|CONTRADA|VICOLO)"
)
SUPPLY_ADDRESS_LABELS = (
    r"Indirizzo\s+(?:di\s+)?fornitura",
    r"Ubicazione\s+(?:della\s+)?fornitura",
    r"Luogo\s+di\s+fornitura",
    r"Punto\s+di\s+fornitura",
)
SUPPLY_ADDRESS_LABEL_PATTERN = "(?:" + "|".join(SUPPLY_ADDRESS_LABELS) + ")"
RESIDENCE_ADDRESS_LABEL_PATTERN = (
    r"(?:Residenza|Indirizzo\s+di\s+residenza|Domicilio|Recapito|Indirizzo\s+di\s+recapito|"
    r"Sede\s+legale|Sede\s+amministrativa|Indirizzo\s+cliente|Fatturare\s+a)"
)
POD_PDR_LABEL_PATTERN = r"(?:Codice\s+POD|Codice\s+PDR|\bPOD\b|\bPDR\b)"


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
        .replace("−", "-")
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


def _month_range(month_name, year):
    month = ITALIAN_MONTHS.get(str(month_name).lower()) or ABBREVIATED_ITALIAN_MONTHS.get(
        str(month_name).lower()[:3]
    )
    if not month:
        return None, None
    start = date(int(year), month, 1)
    if month == 12:
        next_month = date(int(year) + 1, 1, 1)
    else:
        next_month = date(int(year), month + 1, 1)
    return start, date.fromordinal(next_month.toordinal() - 1)


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
        return _float(currency_totals[0])
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


def _looks_like_address(value):
    return bool(re.match(rf"^{ADDRESS_PREFIX_PATTERN}\b", str(value or "").strip(), re.IGNORECASE))


def _clean_address(value):
    address = " ".join(str(value or "").split())
    address = re.split(
        r"\s+(?:Codice\s+POD|Codice\s+PDR|POD|PDR|Periodo|MERCATO|Contratto|Scontrino)\b",
        address,
        maxsplit=1,
        flags=re.IGNORECASE,
    )[0]
    return address.strip(" ,;-. ").title()


def _line_has_supply_address_context(value):
    return bool(re.search(SUPPLY_ADDRESS_LABEL_PATTERN, str(value or ""), re.IGNORECASE))


def _line_has_residence_address_context(value):
    return bool(re.search(RESIDENCE_ADDRESS_LABEL_PATTERN, str(value or ""), re.IGNORECASE))


def _line_has_pod_pdr_context(value):
    return bool(re.search(POD_PDR_LABEL_PATTERN, str(value or ""), re.IGNORECASE))


def _collect_address_candidates(lines, name=""):
    candidates = {}

    def remember(index, raw_value, score):
        cleaned = _clean_address(raw_value)
        if not cleaned:
            return
        current = candidates.get((index, cleaned))
        if current is None or score > current["score"]:
            candidates[(index, cleaned)] = {"index": index, "address": cleaned, "score": score}

    for index, line in enumerate(lines):
        explicit = re.search(
            rf"{SUPPLY_ADDRESS_LABEL_PATTERN}\s*:?[\s-]*({ADDRESS_PREFIX_PATTERN}\b.+)",
            line,
            re.IGNORECASE,
        )
        if explicit:
            remember(index, explicit.group(1), 100)
        if _line_has_supply_address_context(line):
            for offset, candidate in enumerate(lines[index + 1 : index + 4], start=1):
                if _looks_like_address(candidate):
                    remember(index + offset, candidate, 95 - offset)
                    break

    name_norm = " ".join(str(name).split()).lower()
    for index, line in enumerate(lines):
        if not _looks_like_address(line):
            continue
        window_lines = lines[max(0, index - 2) : min(len(lines), index + 3)]
        window_text = " ".join(window_lines)
        score = 10
        if _line_has_residence_address_context(window_text):
            score -= 100
        if _line_has_supply_address_context(window_text):
            score += 70
        if _line_has_pod_pdr_context(window_text):
            score += 40
        remember(index, line, score)

    if name_norm:
        for index, line in enumerate(lines):
            if " ".join(line.split()).lower() != name_norm:
                continue
            for offset, candidate in enumerate(lines[index + 1 : index + 4], start=1):
                if _looks_like_address(candidate):
                    remember(index + offset, candidate, 20 - offset)
                    break
    return sorted(candidates.values(), key=lambda item: (-item["score"], item["index"]))


def _extract_supply_address(text, name=""):
    lines = [" ".join(line.split()).strip() for line in _normalize_text(text).splitlines()]
    lines = [line for line in lines if line]
    candidates = _collect_address_candidates(lines, name=name)
    if candidates:
        strongest = candidates[0]
        if strongest["score"] >= 90:
            return strongest["address"]
        non_residence = [candidate for candidate in candidates if candidate["score"] >= 0]
        unique_addresses = {candidate["address"] for candidate in non_residence}
        if len(unique_addresses) == 1:
            return non_residence[0]["address"]
        pod_candidates = [candidate for candidate in non_residence if candidate["score"] >= 40]
        if pod_candidates:
            return pod_candidates[0]["address"]
    collapsed = _collapsed(text)
    explicit = re.search(
        rf"{SUPPLY_ADDRESS_LABEL_PATTERN}\s*:?[\s-]*({ADDRESS_PREFIX_PATTERN}\b.{{5,140}}?)(?=\s+(?:Codice\s+POD|Codice\s+PDR|POD|PDR|Periodo|MERCATO|Contratto|Scontrino|$))",
        collapsed,
        re.IGNORECASE,
    )
    return _clean_address(explicit.group(1)) if explicit else ""


def _extract_period(text):
    collapsed = _collapsed(text)
    match = _first_match(
        collapsed,
        [
            rf"Periodo (?:oggetto )?di fatturazione\s*:?\s*(?:dal\s*)?({DATE_PATTERN})\s*(?:al|-)\s*({DATE_PATTERN})",
            rf"Periodo di riferimento\s*:?.{{0,180}}?({DATE_PATTERN})\s*-\s*({DATE_PATTERN})",
            rf"Consumo Energia Attiva.*?({DATE_PATTERN})\s*-\s*({DATE_PATTERN})\s+Effettivo",
        ],
    )
    if match:
        return _parse_date(match.group(1)), _parse_date(match.group(2))
    month_match = _first_match(
        collapsed,
        [
            r"Periodo di riferimento\s*:?\s*([a-z]+)\s+(\d{4})",
            r"Periodo\s+([a-z]{3})\s+(\d{4})\s*-\s*([a-z]{3})\s+(\d{4})",
        ],
    )
    if not month_match:
        return None, None
    start, end = _month_range(month_match.group(1), month_match.group(2))
    if month_match.lastindex == 4:
        _, end = _month_range(month_match.group(3), month_match.group(4))
    return start, end


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
    receipt = _receipt_section(text)
    match = _first_match(
        receipt,
        [
            rf"Quota per consumi\s*({NUMBER_PATTERN})\s*(?:(?:Smc|kWh)\b|(?={NUMBER_PATTERN}))",
        ],
    )
    if match:
        return _float(match.group(1))
    match = _first_match(
        collapsed,
        [
            rf"Consumo totale fatturato(?: del periodo)?\s*:?\s*({NUMBER_PATTERN})\s*(?:Smc|kWh)",
            rf"Consumo fatturato(?: nel periodo di fatturazione)?\s*:?\s*({NUMBER_PATTERN})\s*(?:Smc|kWh)",
            rf"Consumo (?:totale )?del periodo(?: di fatturazione)?\s*:?\s*({NUMBER_PATTERN})\s*(?:Smc|kWh)",
            rf"Spesa totale quota consumi\s*({NUMBER_PATTERN})\s*(?:Smc|kWh)",
            rf"A quanto ammonta il consumo fatturato\?.{{0,120}}?({NUMBER_PATTERN})\s*(?:Smc|kWh)",
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
            rf"CONSUMO ANNUO\s*:?\s*Da\s+{DATE_PATTERN}\s+a\s+{DATE_PATTERN}\s*:\s*({NUMBER_PATTERN})\s*(?:Smc|kWh)",
            rf"Consumo annuo aggiornato(?:\s+dal\s+{DATE_PATTERN}\s+al\s+{DATE_PATTERN})?\s*:?\s*({NUMBER_PATTERN})\s*(?:Smc|kWh)",
            rf"Consumo annuo\s*:?\s*({NUMBER_PATTERN})\s*(?:Smc|kWh)",
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
    if re.search(r"Tipologia di offerta\s*:?\s*(?:Offerta )?a prezzo fisso|Tipologia di offerta\s*:?\s*prezzo fisso", collapsed, re.IGNORECASE):
        return "FISSA"
    if re.search(
        r"Tipologia di offerta\s*:?\s*(?:Offerta )?a prezzo variabile|Tipologia di offerta\s*:?\s*prezzo indicizzato",
        collapsed,
        re.IGNORECASE,
    ):
        return "VARIABILE"
    return ""


def _extract_offer_expiry(text):
    collapsed = _collapsed(text)
    match = _first_match(
        collapsed,
        [
            rf"(?:Data di )?[Ss]cadenza condizioni economiche\s*:?\s*({DATE_PATTERN})",
            rf"Data di scadenza delle condizioni economiche\s*:?\s*({DATE_PATTERN})",
            rf"Data scadenza offerta\s*:?\s*({DATE_PATTERN})",
            rf"condizioni economiche dell[' ]offerta sono valide fino al\s*({DATE_PATTERN})",
        ],
    )
    expiry = _parse_date(match.group(1)) if match else None
    return None if expiry and expiry.year >= 9999 else expiry


def _extract_receipt_values(text):
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
    network_label = r"di cui spesa per (?:la )?rete(?:\s+e)?\s+(?:gli\s+)?oneri(?:\s+generali)?\s+di\s+sistema\d*"
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
        values["b_vendita_fissa"] = sale_fixed
    if network_fixed is not None:
        values["b_rete_fissa"] = network_fixed

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
    if not power_match:
        power_match = re.search(
            rf"\b{NUMBER_PATTERN}\s*kW\s+(?:x|per)\s+{NUMBER_PATTERN}\s*mesi?.*?({NUMBER_PATTERN})\s*€(?!\s*/)",
            receipt,
            re.IGNORECASE,
        )
    if power_match:
        values["b_quota_potenza"] = _amount_from_chunk(power_match.group(1))

    for key, pattern in [
        ("b_ricalcoli", rf"Totale ricalcoli\s*({NUMBER_PATTERN})\s*€?"),
        (
            "b_arrotondamenti",
            rf"Arrotondament[oi]\s*(?:(?:\([^)]*\)|attuale|corrente)\s*)*:?\s*({NUMBER_PATTERN})\s*€",
        ),
        ("b_accise_iva", rf"Accise e IVA\s*({NUMBER_PATTERN})\s*€?"),
    ]:
        match = re.search(pattern, receipt, re.IGNORECASE)
        if match:
            values[key] = _float(match.group(1))
    social_bonus = re.search(
        rf"Altre partite[^€]{{0,160}}Bonus sociale[^€]{{0,100}}({NUMBER_PATTERN})\s*€",
        receipt,
        re.IGNORECASE,
    )
    if social_bonus:
        values["b_bonus_sociale"] = _float(social_bonus.group(1))
    supplier_discount = re.search(
        rf"Altre partite[^€]{{0,220}}Bonus(?! sociale).*?({NUMBER_PATTERN})\s*€",
        receipt,
        re.IGNORECASE,
    )
    if supplier_discount:
        values["b_sconti"] = _float(supplier_discount.group(1))
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
    supply_address = _extract_supply_address(normalized, name)
    if supply_address:
        values["indirizzo_fornitura"] = supply_address
    pod_pdr = _extract_pod_pdr(normalized, commodity)
    if pod_pdr:
        values["pod_pdr"] = pod_pdr
    consumo = _extract_consumption(normalized)
    if consumo is not None:
        values["consumo"] = consumo
    annual = _extract_annual_consumption(normalized)
    if annual is not None:
        values["tax_annual_consumption"] = annual
    power_match = _first_match(
        normalized,
        [
            rf"Potenza impegnata\s*\(kW\)\s*:?\s*({NUMBER_PATTERN})",
            rf"Potenza impegnata.{{0,220}}?({NUMBER_PATTERN})\s*kW(?!h)",
        ],
        flags=re.IGNORECASE | re.DOTALL,
    )
    if power_match:
        values["tax_power_kw"] = _float(power_match.group(1))
    tariff_type = _extract_bill_tariff_type(normalized)
    if tariff_type:
        values["bill_tariff_type"] = tariff_type
    expiry = _extract_offer_expiry(normalized)
    if expiry:
        values["bill_offer_expiry"] = expiry
    values.update(_extract_receipt_values(normalized))

    if commodity == "GAS":
        values["tax_power_kw"] = 0
        values["b_quota_potenza"] = 0
    for field_name in ("b_sconti", "b_ricalcoli", "b_bonus_sociale", "b_arrotondamenti", "b_servizi_accessori"):
        values.setdefault(field_name, 0)
    if not expiry:
        warnings.append("Data fine offerta bolletta non riconosciuta: compilala manualmente.")
    if not name:
        warnings.append("Nome cliente non riconosciuto: compilalo manualmente.")
    if not supply_address:
        warnings.append("Indirizzo fornitura non riconosciuto: compilalo manualmente.")
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
