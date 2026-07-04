from .utils import (
    KEYS,
    INDICI_XLSX,
    clean_text,
    normalize_provider,
    normalize_providers,
    normalize_bill_tariff_type,
    bill_tariff_type_label,
    normalize_tariff_selection_mode,
    tariff_selection_mode_label,
    normalize_accessory_services_vat_label,
    normalize_primary_home,
    primary_home_label,
    normalize_region,
    accessory_services_vat_rate,
    comparison_datetime_from_data,
    comparison_datetime_label,
    date_label_it,
    provider_list_label,
    provider_label,
    format_eur,
    format_percent,
    fiscal_unit,
    comparison_subtotal,
    comparison_total,
    comparison_value,
    bill_period_label,
    bill_period_outside_offer_validity,
    billing_months_from_dates,
    billing_divisor_from_months,
    billing_label_from_months,
    month_key_from_date,
    tax_incidence_ratio,
    safe_download_filename,
    bool_from_data,
)
from .tariffs import (
    get_file_valid_range,
    load_tariffe_from_path,
    load_tariffe_file_for_segment_with_effective_segment,
    filter_rows_by_context_with_fallback,
    filter_cve_rows_by_annual_context,
    filter_rows_by_offer,
    select_offer_name,
    load_indici_rows,
    select_indice_for_bill_period,
    valid_range_from_rows,
)
from .tax import (
    bill_tax_incidence_ratio,
    calculate_tax_breakdown,
    cap_tax_breakdown_to_bill_incidence,
)


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
    prezzo += comp_sum(rows_offer, "GAS", "FISSA", "ccv_quota_variabile")
    return float(consumo * prezzo), float(fixed_annua("GAS", "FISSA") / float(billing_divisor))


def format_fiscal_parameters(data, calc=None):
    comm = data.get("commodity", "EE")
    return (
        f"Prima casa/residente: {primary_home_label(data.get('tax_primary_home'))} | "
        f"Potenza: {float(data.get('tax_power_kw', 0.0)):g} kW | "
        f"Consumo annuo: {float(data.get('tax_annual_consumption', 0.0)):g} {fiscal_unit(comm)}/anno | "
        f"Regione: {normalize_region(data.get('tax_region'))} | "
        f"IVA servizi accessori: {normalize_accessory_services_vat_label(data.get('servizi_accessori_iva'))}"
    )


def build_comparison_table_rows(values):
    comm = values["commodity"]
    servizi_accessori_label = f"Servizi accessori (IVA {values['servizi_accessori_iva_label']})"
    accise_label = "Accise + Addizionale regionale" if comm == "GAS" else "Accise"
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
        ("servizi_accessori", servizi_accessori_label),
        ("accise", accise_label),
        ("iva", "IVA"),
        ("accise_iva", "Accise e Iva"),
        ("totale", "Totale"),
    ]
    out = []
    for key, label in rows_config:
        if comm == "GAS" and key in {"vendita_fissa_luce", "quota_potenza"}:
            continue
        cells = [format_eur(comparison_value(values["bolletta"], key, comm))]
        for column in values["offer_columns"]:
            value = comparison_value(column["vals"], key, comm) if column["has_offer"] else "N.D."
            cells.append(format_eur(value))
        row = {
            "voce": label,
            "cells": cells,
            "bolletta": cells[0],
        }
        if len(cells) >= 3:
            row["variabile"] = cells[1]
            row["fissa"] = cells[2]
        out.append(row)
    return out


def calculate_provider_result(data, base_calc, provider):
    bill_start = data["bill_start"]
    bill_end = data["bill_end"]
    segmento = data["segmento"]
    commodity = data["commodity"]
    provider_norm = normalize_provider(provider)
    consumo = float(data["consumo"])
    billing_divisor = base_calc["billing_divisor"]
    pun = base_calc["pun"]
    psv = base_calc["psv"]

    offer_file, effective_segment = load_tariffe_file_for_segment_with_effective_segment(
        segmento,
        provider_norm,
        base_calc.get("tariff_selection_mode", "LATEST"),
        base_calc.get("tariff_target_month", ""),
        commodity,
    )
    if offer_file:
        raw_tariffe_rows = load_tariffe_from_path(offer_file)
        tariffe_rows, effective_segment = filter_rows_by_context_with_fallback(
            raw_tariffe_rows, provider_norm, effective_segment, commodity
        )
        if provider_norm == "CVE":
            tariffe_rows = filter_cve_rows_by_annual_context(
                tariffe_rows, commodity, data.get("tax_annual_consumption", 0.0), data.get("cve_over70")
            )
    else:
        tariffe_rows = []
    offer_valid_from, offer_valid_to = get_file_valid_range(offer_file) if offer_file else (None, None)
    offer_var = select_offer_name(
        tariffe_rows,
        commodity,
        "VARIABILE",
        effective_segment,
        offer_choice_from_data(data, provider_norm, "VARIABILE"),
    )
    offer_fix = select_offer_name(
        tariffe_rows,
        commodity,
        "FISSA",
        effective_segment,
        offer_choice_from_data(data, provider_norm, "FISSA"),
    )

    sconto_var = float(data.get("ill_sconto_var", -3.0)) if provider_norm == "ILLUMIA" else 0.0
    sconto_fix = float(data.get("ill_sconto_fix", -3.0)) if provider_norm == "ILLUMIA" else 0.0
    rows_var = []
    rows_fix = []

    if offer_var:
        rows_var = filter_rows_by_offer(tariffe_rows, commodity, "VARIABILE", offer_var)
        v_cons, v_fix = calc_illumia_vendite(rows_var, commodity, "VARIABILE", consumo, pun, psv, billing_divisor)
        if provider_norm != "ILLUMIA":
            sconto_var = comp_sum(rows_var, commodity, "VARIABILE", "sconto_bonus")
    else:
        v_cons, v_fix = 0.0, 0.0

    if offer_fix:
        rows_fix = filter_rows_by_offer(tariffe_rows, commodity, "FISSA", offer_fix)
        f_cons, f_fix = calc_illumia_vendite(rows_fix, commodity, "FISSA", consumo, pun, psv, billing_divisor)
        if provider_norm != "ILLUMIA":
            sconto_fix = comp_sum(rows_fix, commodity, "FISSA", "sconto_bonus")
    else:
        f_cons, f_fix = 0.0, 0.0

    selected_valid_from, selected_valid_to = valid_range_from_rows(rows_var + rows_fix)
    if selected_valid_from or selected_valid_to:
        offer_valid_from, offer_valid_to = selected_valid_from, selected_valid_to
    offer_period_warning = ""
    if bill_period_outside_offer_validity(bill_start, bill_end, offer_valid_from, offer_valid_to):
        offer_period_warning = "Il periodo bolletta NON rientra nella validità dell'offerta selezionata."

    return {
        "provider": provider_norm,
        "provider_label": provider_label(provider_norm),
        "offer_file": str(offer_file) if offer_file else "",
        "requested_segment": segmento,
        "tariff_segment": effective_segment,
        "offer_valid_from": offer_valid_from,
        "offer_valid_to": offer_valid_to,
        "offer_expiry_label": date_label_it(offer_valid_to),
        "offer_period_warning": offer_period_warning,
        "offer_var": offer_var,
        "offer_fix": offer_fix,
        "v_cons": v_cons,
        "v_fix": v_fix,
        "f_cons": f_cons,
        "f_fix": f_fix,
        "sconto_var": sconto_var,
        "sconto_fix": sconto_fix,
    }


def offer_choice_from_data(data, provider, offer_type):
    provider_norm = normalize_provider(provider)
    suffix = provider_norm.lower()
    choice_type = "var" if offer_type.upper() == "VARIABILE" else "fix"
    choice = clean_text(data.get(f"offer_{choice_type}_choice_{suffix}", ""))
    if not choice and normalize_provider(data.get("provider", provider_norm)) == provider_norm:
        choice = clean_text(data.get(f"offer_{choice_type}_choice", ""))
    return choice


def build_offer_column_values(data, calc, base_values, provider_result, offer_type, bill_ratio=None):
    comm = data["commodity"]
    vals = base_values.copy()
    if offer_type == "VARIABILE":
        vals["vendita_consumo"] = provider_result["v_cons"]
        vals["vendita_fissa"] = provider_result["v_fix"]
        vals["sconti"] = float(provider_result.get("sconto_var", data.get("ill_sconto_var", -3.0)))
        offer_name = provider_result["offer_var"]
    else:
        vals["vendita_consumo"] = provider_result["f_cons"]
        vals["vendita_fissa"] = provider_result["f_fix"]
        vals["sconti"] = float(provider_result.get("sconto_fix", data.get("ill_sconto_fix", -3.0)))
        offer_name = provider_result["offer_fix"]
    vals["ricalcoli"] = 0.0
    vals["arrotondamenti"] = 0.0
    vals["servizi_accessori"] = 0.0
    subtotal = comparison_subtotal(vals, comm)
    raw_tax = calculate_tax_breakdown(data, calc, vals, comm)
    raw_tax_ratio = tax_incidence_ratio(raw_tax["accise_iva"], subtotal)
    tax = cap_tax_breakdown_to_bill_incidence(raw_tax, vals, comm, bill_ratio)
    final_tax_ratio = tax_incidence_ratio(tax["accise_iva"], subtotal)
    tax_cap_applied = bool(
        bill_ratio is not None and raw_tax_ratio is not None and raw_tax_ratio > float(bill_ratio) + 1e-9
    )
    vals["accise"] = tax["accise"]
    vals["iva"] = tax["iva"]
    vals["accise_iva"] = tax["accise_iva"]
    type_label = "Variabile" if offer_type == "VARIABILE" else "Fissa"
    return {
        "provider": provider_result["provider"],
        "provider_label": provider_result["provider_label"],
        "offer_type": offer_type,
        "offer_type_label": type_label,
        "offer_name": offer_name,
        "label": f"{provider_result['provider_label']} {type_label}",
        "vals": vals,
        "has_offer": bool(offer_name),
        "tax_cap_applied": tax_cap_applied,
        "tax_cap_status_label": "Si" if tax_cap_applied else "No",
        "bill_tax_ratio": bill_ratio,
        "bill_tax_ratio_label": format_percent(bill_ratio),
        "raw_tax_ratio": raw_tax_ratio,
        "raw_tax_ratio_label": format_percent(raw_tax_ratio),
        "tax_ratio": final_tax_ratio,
        "tax_ratio_label": format_percent(final_tax_ratio),
    }


def build_comparison_values(data, calc):
    comm = data["commodity"]
    accessory_vat_label = normalize_accessory_services_vat_label(data.get("servizi_accessori_iva"))
    accessory_vat_rate_val = accessory_services_vat_rate(accessory_vat_label)
    b_vals = {k: float(data.get(f"b_{k}", 0.0)) for k in KEYS}
    b_vals["bonus_sociale"] = -abs(float(b_vals.get("bonus_sociale", 0.0)))
    if comm == "GAS":
        b_vals["quota_potenza"] = 0.0
    b_vals["accise"] = None
    b_vals["iva"] = None
    bill_ratio = bill_tax_incidence_ratio(b_vals, comm)

    provider_results = calc.get("provider_results") or [
        {
            "provider": normalize_provider(calc.get("provider", data.get("provider", "ILLUMIA"))),
            "provider_label": provider_label(calc.get("provider", data.get("provider", "ILLUMIA"))),
            "offer_var": calc.get("offer_var", ""),
            "offer_fix": calc.get("offer_fix", ""),
            "v_cons": calc.get("v_cons", 0.0),
            "v_fix": calc.get("v_fix", 0.0),
            "f_cons": calc.get("f_cons", 0.0),
            "f_fix": calc.get("f_fix", 0.0),
            "sconto_var": calc.get("sconto_var", data.get("ill_sconto_var", -3.0)),
            "sconto_fix": calc.get("sconto_fix", data.get("ill_sconto_fix", -3.0)),
        }
    ]
    offer_columns = []
    for provider_result in provider_results:
        offer_columns.append(build_offer_column_values(data, calc, b_vals, provider_result, "VARIABILE", bill_ratio))
        offer_columns.append(build_offer_column_values(data, calc, b_vals, provider_result, "FISSA", bill_ratio))

    first_var = next((column for column in offer_columns if column["offer_type"] == "VARIABILE"), None)
    first_fix = next((column for column in offer_columns if column["offer_type"] == "FISSA"), None)
    return {
        "commodity": comm,
        "bolletta": b_vals,
        "offer_columns": offer_columns,
        "variabile": first_var["vals"] if first_var else b_vals.copy(),
        "fissa": first_fix["vals"] if first_fix else b_vals.copy(),
        "bill_tax_ratio": bill_ratio,
        "bill_tax_ratio_label": format_percent(bill_ratio),
        "servizi_accessori_iva_label": accessory_vat_label,
        "servizi_accessori_iva_rate": accessory_vat_rate_val,
        "has_offer_var": bool(first_var and first_var["has_offer"]),
        "has_offer_fix": bool(first_fix and first_fix["has_offer"]),
    }


def prepare_comparison(data):
    bill_start = data["bill_start"]
    bill_end = data["bill_end"]
    bill_offer_expiry = data.get("bill_offer_expiry")
    segmento = data["segmento"]
    commodity = data["commodity"]
    providers = normalize_providers(data.get("providers") or data.get("provider", "ILLUMIA"))
    if not providers:
        providers = [normalize_provider(data.get("provider", "ILLUMIA"))]
    tariff_selection_mode = normalize_tariff_selection_mode(data.get("tariff_selection_mode"))
    billing_months = billing_months_from_dates(bill_start, bill_end)
    billing_divisor = billing_divisor_from_months(billing_months)
    tariff_target_month = month_key_from_date(bill_end)

    indici_rows = load_indici_rows(INDICI_XLSX)
    indice, indice_reason = select_indice_for_bill_period(indici_rows, bill_start, bill_end)
    pun = float(indice["pun"]) if indice else 0.0
    psv = float(indice["psv"]) if indice else 0.0

    calc = {
        "nome_cliente": clean_text(data.get("nome_cliente")) or "Cliente",
        "indirizzo_fornitura": clean_text(data.get("indirizzo_fornitura")),
        "pod_pdr": clean_text(data.get("pod_pdr")),
        "commodity": commodity,
        "bill_tariff_type": normalize_bill_tariff_type(data.get("bill_tariff_type")),
        "bill_tariff_type_label": bill_tariff_type_label(data.get("bill_tariff_type")),
        "tax_primary_home": normalize_primary_home(data.get("tax_primary_home")),
        "tax_primary_home_label": primary_home_label(data.get("tax_primary_home")),
        "tax_power_kw": float(data.get("tax_power_kw", 0.0)),
        "tax_annual_consumption": float(data.get("tax_annual_consumption", 0.0)),
        "tax_annual_consumption_label": (
            f"{float(data.get('tax_annual_consumption', 0.0)):g} {fiscal_unit(commodity)}/anno"
        ),
        "tax_region": normalize_region(data.get("tax_region")),
        "fiscal_parameters_label": format_fiscal_parameters(data),
        "comparison_datetime": comparison_datetime_from_data(data.get("comparison_datetime")),
        "comparison_datetime_label": comparison_datetime_label(data.get("comparison_datetime")),
        "providers": providers,
        "providers_label": provider_list_label(providers),
        "tariff_selection_mode": tariff_selection_mode,
        "tariff_selection_mode_label": tariff_selection_mode_label(tariff_selection_mode),
        "tariff_target_month": tariff_target_month,
        "provider": providers[0],
        "provider_label": provider_label(providers[0]),
        "billing_months": billing_months,
        "billing_divisor": billing_divisor,
        "billing_label": billing_label_from_months(billing_months),
        "period_label": bill_period_label(bill_start, bill_end),
        "bill_offer_expiry": bill_offer_expiry,
        "bill_offer_expiry_label": date_label_it(bill_offer_expiry),
        "indice": indice,
        "indice_reason": indice_reason,
        "pun": pun,
        "psv": psv,
        "servizi_accessori_iva_label": normalize_accessory_services_vat_label(data.get("servizi_accessori_iva")),
        "cve_over70": bool_from_data(data.get("cve_over70")),
        "cve_over70_label": "Si" if bool_from_data(data.get("cve_over70")) else "No",
        "cve_selected": "CVE" in providers,
    }
    provider_results = [calculate_provider_result(data, calc, provider) for provider in providers]
    calc["provider_results"] = provider_results
    primary = provider_results[0] if provider_results else {}
    calc.update(
        {
            "offer_file": primary.get("offer_file", ""),
            "offer_valid_from": primary.get("offer_valid_from"),
            "offer_valid_to": primary.get("offer_valid_to"),
            "offer_expiry_label": calc["bill_offer_expiry_label"],
            "offer_period_warning": " ".join(
                f"{result['provider_label']}: {result['offer_period_warning']}"
                for result in provider_results
                if result.get("offer_period_warning")
            ),
            "offer_var": primary.get("offer_var", ""),
            "offer_fix": primary.get("offer_fix", ""),
            "v_cons": primary.get("v_cons", 0.0),
            "v_fix": primary.get("v_fix", 0.0),
            "f_cons": primary.get("f_cons", 0.0),
            "f_fix": primary.get("f_fix", 0.0),
            "sconto_var": primary.get("sconto_var", 0.0),
            "sconto_fix": primary.get("sconto_fix", 0.0),
        }
    )
    values = build_comparison_values(data, calc)
    columns = [{"label": "Bolletta"}] + [{"label": column["label"]} for column in values["offer_columns"]]
    return {"calc": calc, "values": values, "rows": build_comparison_table_rows(values), "columns": columns}
