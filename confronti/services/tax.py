from .utils import (
    ACCESSORY_SERVICES_VAT_RATES,
    EE_EXCISE_RATE,
    EE_RESIDENTIAL_EXEMPT_KWH_PER_MONTH,
    GAS_EXCISE_BRACKETS,
    GAS_REGIONAL_ADDITIONAL_MIN_RATES,
    GAS_VAT_REDUCED_ANNUAL_THRESHOLD,
    accessory_services_vat_rate,
    annual_progressive_tax_for_period,
    comparison_subtotal,
    normalize_accessory_services_vat_label,
    normalize_primary_home,
    normalize_region,
    tax_incidence_ratio,
)


def calculate_tax_breakdown(data, calc, vals, commodity):
    period_consumption = max(0.0, float(data.get("consumo", 0.0)))
    accessory_vat_rate_value = accessory_services_vat_rate(data.get("servizi_accessori_iva"))
    accessory_services = float(vals.get("servizi_accessori", 0.0))
    social_bonus = float(vals.get("bonus_sociale", 0.0))
    taxable_supply_subtotal = comparison_subtotal(vals, commodity) - accessory_services - social_bonus
    accessory_services_vat = accessory_services * accessory_vat_rate_value

    if commodity == "EE":
        power_kw = max(0.0, float(data.get("tax_power_kw", 0.0)))
        residential = data.get("segmento") == "RESIDENZIALE"
        primary_home = normalize_primary_home(data.get("tax_primary_home")) == "SI"
        exempt_kwh = EE_RESIDENTIAL_EXEMPT_KWH_PER_MONTH * float(calc.get("billing_months", 1))
        if residential and primary_home and power_kw <= 3.0:
            taxable_kwh = max(0.0, period_consumption - exempt_kwh)
        else:
            taxable_kwh = period_consumption
        excise = taxable_kwh * EE_EXCISE_RATE
        vat_rate = 0.10 if residential and primary_home else 0.22
        vat = ((taxable_supply_subtotal + excise) * vat_rate) + accessory_services_vat
        return {
            "accise": excise,
            "iva": vat,
            "accise_iva": excise + vat,
            "addizionale_regionale": 0.0,
        }

    annual_consumption = max(0.0, float(data.get("tax_annual_consumption", 0.0)))
    excise = annual_progressive_tax_for_period(period_consumption, annual_consumption, GAS_EXCISE_BRACKETS)
    regional = period_consumption * float(GAS_REGIONAL_ADDITIONAL_MIN_RATES[normalize_region(data.get("tax_region"))])
    variable_base = (
        float(vals["vendita_consumo"])
        + float(vals["rete_consumi"])
        + float(vals["sconti"])
        + excise
        + regional
    )
    fixed_base = (
        float(vals["vendita_fissa"])
        + float(vals["rete_fissa"])
        + float(vals["ricalcoli"])
        + float(vals["arrotondamenti"])
    )
    reduced_ratio = min(annual_consumption, GAS_VAT_REDUCED_ANNUAL_THRESHOLD) / annual_consumption if annual_consumption else 0.0
    reduced_ratio = min(max(reduced_ratio, 0.0), 1.0)
    vat = (
        (variable_base * reduced_ratio * 0.10)
        + (variable_base * (1.0 - reduced_ratio) * 0.22)
        + (fixed_base * 0.22)
        + accessory_services_vat
    )
    return {
        "accise": excise + regional,
        "iva": vat,
        "accise_iva": excise + regional + vat,
        "addizionale_regionale": regional,
    }


def calculate_accise_iva(data, calc, vals, commodity):
    return calculate_tax_breakdown(data, calc, vals, commodity)["accise_iva"]


def bill_tax_incidence_ratio(vals, commodity):
    subtotal = comparison_subtotal(vals, commodity)
    return tax_incidence_ratio(vals.get("accise_iva", 0.0), subtotal)


def cap_tax_breakdown_to_bill_incidence(tax, vals, commodity, bill_ratio):
    if bill_ratio is None:
        return tax
    subtotal = comparison_subtotal(vals, commodity)
    current_total = max(0.0, float(tax.get("accise_iva", 0.0)))
    capped_total = max(0.0, float(bill_ratio)) * max(0.0, subtotal)
    if subtotal <= 0 or current_total <= 0 or current_total <= capped_total:
        return tax
    scale = capped_total / current_total
    accise = float(tax.get("accise", 0.0)) * scale
    iva = float(tax.get("iva", 0.0)) * scale
    return {
        "accise": accise,
        "iva": iva,
        "accise_iva": accise + iva,
        "addizionale_regionale": float(tax.get("addizionale_regionale", 0.0)) * scale,
    }
