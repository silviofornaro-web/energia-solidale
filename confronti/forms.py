from calendar import monthrange
from datetime import date

from django import forms

from . import services


class ItalianDecimalField(forms.DecimalField):
    def to_python(self, value):
        if value in self.empty_values:
            return None
        if isinstance(value, str):
            value = value.strip().replace("€", "").replace("\u20ac", "").replace(" ", "")
            if "," in value and "." in value:
                value = value.replace(".", "").replace(",", ".")
            else:
                value = value.replace(",", ".")
        return super().to_python(value)


class ConfrontoForm(forms.Form):
    SEGMENTI = [("RESIDENZIALE", "Residenziale"), ("MICROBUSINESS", "Microbusiness"), ("BUSINESS", "Business")]
    COMMODITIES = [("GAS", "Gas"), ("EE", "Luce")]
    PROVIDERS = [("ILLUMIA", "Illumia"), ("EON", "E.ON")]
    BILL_TARIFF_TYPES = [("VARIABILE", "Variabile"), ("FISSA", "Fissa")]
    PRIMARY_HOME_CHOICES = [("SI", "Sì"), ("NO", "No")]
    SEGMENT_CHOICES = [("", "Seleziona segmento")] + SEGMENTI
    COMMODITY_CHOICES = [("", "Seleziona fornitura")] + COMMODITIES
    PROVIDER_CHOICES = PROVIDERS
    BILL_TARIFF_TYPE_CHOICES = [("", "Seleziona tariffa")] + BILL_TARIFF_TYPES
    REGION_CHOICES = [(region, region) for region in services.GAS_REGIONAL_ADDITIONAL_MIN_RATES]
    MONTH_INPUT_FORMATS = ["%Y-%m", "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y"]

    nome_cliente = forms.CharField(label="Nome e Cognome", max_length=120)
    segmento = forms.ChoiceField(label="Segmento", choices=SEGMENT_CHOICES)
    commodity = forms.ChoiceField(label="Fornitura", choices=COMMODITY_CHOICES)
    bill_tariff_type = forms.ChoiceField(label="Tipo tariffa bolletta", choices=BILL_TARIFF_TYPE_CHOICES)
    providers = forms.MultipleChoiceField(
        label="Fornitori confronto",
        choices=PROVIDER_CHOICES,
        widget=forms.CheckboxSelectMultiple,
    )
    tax_primary_home = forms.ChoiceField(
        label="Prima casa / residente",
        choices=PRIMARY_HOME_CHOICES,
        initial="SI",
    )
    tax_power_kw = ItalianDecimalField(
        label="Potenza impegnata (kW, solo luce)",
        min_value=0,
        decimal_places=2,
        max_digits=8,
        initial=0,
        required=False,
    )
    tax_annual_consumption = ItalianDecimalField(
        label="Consumo annuo stimato (Smc/kWh anno)",
        min_value=0.0001,
        decimal_places=4,
        max_digits=12,
    )
    tax_region = forms.ChoiceField(label="Regione (solo gas)", choices=REGION_CHOICES, initial="Veneto")
    servizi_accessori_iva = forms.ChoiceField(
        label="IVA servizi accessori",
        choices=[(label, label) for label in services.ACCESSORY_SERVICES_VAT_OPTIONS],
        initial="22%",
    )
    offer_var_choice_illumia = forms.ChoiceField(label="Illumia - Offerta variabile", required=False)
    offer_fix_choice_illumia = forms.ChoiceField(label="Illumia - Offerta fissa", required=False)
    offer_var_choice_eon = forms.ChoiceField(label="E.ON - Offerta variabile", required=False)
    offer_fix_choice_eon = forms.ChoiceField(label="E.ON - Offerta fissa", required=False)
    bill_start = forms.DateField(
        label="Dal mese",
        input_formats=MONTH_INPUT_FORMATS,
        widget=forms.DateInput(format="%Y-%m", attrs={"type": "month"}),
    )
    bill_end = forms.DateField(
        label="Al mese",
        input_formats=MONTH_INPUT_FORMATS,
        widget=forms.DateInput(format="%Y-%m", attrs={"type": "month"}),
    )
    consumo = ItalianDecimalField(label="Consumo", min_value=0, decimal_places=4, max_digits=12, initial=0)

    b_vendita_consumo = ItalianDecimalField(label="Vendita consumo", decimal_places=4, max_digits=12, initial=0)
    b_rete_consumi = ItalianDecimalField(label="Rete/oneri consumi", decimal_places=4, max_digits=12, initial=0)
    b_vendita_fissa = ItalianDecimalField(label="Vendita fissa mensile", decimal_places=4, max_digits=12, initial=0)
    b_rete_fissa = ItalianDecimalField(label="Rete/oneri fissa mensile", decimal_places=4, max_digits=12, initial=0)
    b_quota_potenza = ItalianDecimalField(
        label="Quota potenza",
        decimal_places=4,
        max_digits=12,
        initial=0,
        required=False,
    )
    b_sconti = ItalianDecimalField(label="Sconti", decimal_places=4, max_digits=12, initial=0)
    b_ricalcoli = ItalianDecimalField(label="Ricalcoli/Partite pregresse", decimal_places=4, max_digits=12, initial=0)
    b_bonus_sociale = ItalianDecimalField(
        label="Bonus Sociale",
        decimal_places=4,
        max_digits=12,
        initial=0,
        required=False,
        help_text="Inserire l'importo come valore negativo.",
    )
    b_arrotondamenti = ItalianDecimalField(label="Arrotondamenti", decimal_places=4, max_digits=12, initial=0)
    b_servizi_accessori = ItalianDecimalField(
        label="Servizi accessori (imponibile)",
        min_value=0,
        decimal_places=4,
        max_digits=12,
        initial=0,
        required=False,
    )
    bill_offer_expiry = forms.DateField(
        label="Data fine offerta bolletta",
        input_formats=["%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y"],
        widget=forms.DateInput(format="%Y-%m-%d", attrs={"type": "date"}),
    )
    b_accise_iva = ItalianDecimalField(label="Accise + IVA", decimal_places=4, max_digits=12, initial=0)

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        segmento = self._current_value("segmento")
        commodity = self._current_value("commodity")
        self.fields["offer_var_choice_illumia"].choices = self._offer_choices("ILLUMIA", segmento, commodity, "VARIABILE")
        self.fields["offer_fix_choice_illumia"].choices = self._offer_choices("ILLUMIA", segmento, commodity, "FISSA")
        self.fields["offer_var_choice_eon"].choices = self._offer_choices("EON", segmento, commodity, "VARIABILE")
        self.fields["offer_fix_choice_eon"].choices = self._offer_choices("EON", segmento, commodity, "FISSA")
        self._apply_commodity_rules(commodity)

    def _current_value(self, field_name):
        if self.data and field_name in self.data:
            return self.data.get(field_name) or ""
        return self.initial.get(field_name, self.fields[field_name].initial or "")

    def _offer_choices(self, provider, segmento, commodity, offer_type):
        if not provider or not segmento or not commodity:
            return [("", "N.D.")]
        payload = services.offer_options_payload()
        key = f"{services.normalize_provider(provider)}|{str(segmento).upper()}|{str(commodity).upper()}"
        names = payload.get(key, {}).get(offer_type, [])
        return [("", "Automatica")] + [(name, name) for name in names]

    def _apply_commodity_rules(self, commodity):
        if str(commodity).upper() != "GAS":
            return
        for field_name in ("tax_power_kw", "b_quota_potenza"):
            self.fields[field_name].disabled = True
            self.fields[field_name].required = False
            self.fields[field_name].widget.attrs["disabled"] = "disabled"

    def clean(self):
        cleaned = super().clean()
        bill_start = cleaned.get("bill_start")
        bill_end = cleaned.get("bill_end")
        if bill_start:
            cleaned["bill_start"] = date(bill_start.year, bill_start.month, 1)
            bill_start = cleaned["bill_start"]
        if bill_end:
            last_day = monthrange(bill_end.year, bill_end.month)[1]
            cleaned["bill_end"] = date(bill_end.year, bill_end.month, last_day)
            bill_end = cleaned["bill_end"]
        if cleaned.get("bill_start") and cleaned.get("bill_end") and cleaned["bill_end"] < cleaned["bill_start"]:
            raise forms.ValidationError("Il mese finale non può essere precedente al mese iniziale.")
        commodity = cleaned.get("commodity")
        providers = [services.normalize_provider(provider) for provider in cleaned.get("providers", [])]
        cleaned["providers"] = providers
        cleaned["provider"] = providers[0] if providers else ""
        if cleaned.get("tax_annual_consumption") is None:
            self.add_error("tax_annual_consumption", "Indica il consumo annuo stimato.")
        if commodity == "GAS":
            cleaned["tax_power_kw"] = 0
            cleaned["b_quota_potenza"] = 0
        if commodity == "EE" and not cleaned.get("tax_power_kw"):
            self.add_error("tax_power_kw", "Indica la potenza impegnata per il calcolo accise/IVA luce.")
        cleaned["tax_primary_home"] = services.normalize_primary_home(cleaned.get("tax_primary_home"))
        cleaned["tax_region"] = services.normalize_region(cleaned.get("tax_region"))
        bonus = cleaned.get("b_bonus_sociale")
        cleaned["b_bonus_sociale"] = -abs(bonus) if bonus else 0
        cleaned["b_servizi_accessori"] = cleaned.get("b_servizi_accessori") or 0
        cleaned["servizi_accessori_iva"] = services.normalize_accessory_services_vat_label(
            cleaned.get("servizi_accessori_iva")
        )
        return cleaned

    def service_data(self):
        data = dict(self.cleaned_data)
        data["ill_sconto_var"] = -3.0
        data["ill_sconto_fix"] = -3.0
        return data

    def session_data(self):
        data = self.service_data()
        out = {}
        for key, value in data.items():
            if isinstance(value, date):
                out[key] = value.isoformat()
            elif isinstance(value, (list, tuple)):
                out[key] = ",".join(str(item) for item in value)
            else:
                out[key] = str(value)
        return out


def session_to_service_data(raw):
    data = dict(raw)
    data["bill_start"] = services.parse_date_any(data.get("bill_start"))
    data["bill_end"] = services.parse_date_any(data.get("bill_end"))
    data["bill_offer_expiry"] = services.parse_date_any(data.get("bill_offer_expiry"))
    for key in [
        "consumo",
        "tax_power_kw",
        "tax_annual_consumption",
    ] + [f"b_{k}" for k in services.KEYS] + ["ill_sconto_var", "ill_sconto_fix"]:
        data[key] = services.parse_number(data.get(key))
    data["providers"] = services.normalize_providers(data.get("providers") or data.get("provider", "ILLUMIA"))
    data["provider"] = data["providers"][0] if data["providers"] else "ILLUMIA"
    data["bill_tariff_type"] = services.normalize_bill_tariff_type(data.get("bill_tariff_type"))
    data["tax_primary_home"] = services.normalize_primary_home(data.get("tax_primary_home"))
    data["tax_region"] = services.normalize_region(data.get("tax_region"))
    data["servizi_accessori_iva"] = services.normalize_accessory_services_vat_label(data.get("servizi_accessori_iva"))
    data["comparison_datetime"] = data.get("comparison_datetime")
    data["offer_var_choice_illumia"] = data.get("offer_var_choice_illumia", data.get("offer_var_choice", ""))
    data["offer_fix_choice_illumia"] = data.get("offer_fix_choice_illumia", data.get("offer_fix_choice", ""))
    data["offer_var_choice_eon"] = data.get("offer_var_choice_eon", "")
    data["offer_fix_choice_eon"] = data.get("offer_fix_choice_eon", "")
    return data
