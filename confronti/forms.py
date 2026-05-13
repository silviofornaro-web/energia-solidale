from calendar import monthrange
from datetime import date

from django import forms

from . import services


class ConfrontoForm(forms.Form):
    SEGMENTI = [("RESIDENZIALE", "Residenziale"), ("MICROBUSINESS", "Microbusiness"), ("BUSINESS", "Business")]
    COMMODITIES = [("GAS", "Gas"), ("EE", "Luce")]
    PROVIDERS = [("ILLUMIA", "Illumia"), ("EON", "E.ON")]
    MONTH_INPUT_FORMATS = ["%Y-%m", "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y"]

    nome_cliente = forms.CharField(label="Nome e Cognome", max_length=120, initial="Cliente")
    segmento = forms.ChoiceField(label="Segmento", choices=SEGMENTI, initial="RESIDENZIALE")
    commodity = forms.ChoiceField(label="Fornitura", choices=COMMODITIES, initial="GAS")
    provider = forms.ChoiceField(label="Fornitore confronto", choices=PROVIDERS, initial="ILLUMIA")
    offer_var_choice = forms.ChoiceField(label="Offerta variabile", required=False)
    offer_fix_choice = forms.ChoiceField(label="Offerta fissa", required=False)
    bill_start = forms.DateField(
        label="Dal mese",
        initial=date.today,
        input_formats=MONTH_INPUT_FORMATS,
        widget=forms.DateInput(format="%Y-%m", attrs={"type": "month"}),
    )
    bill_end = forms.DateField(
        label="Al mese",
        initial=date.today,
        input_formats=MONTH_INPUT_FORMATS,
        widget=forms.DateInput(format="%Y-%m", attrs={"type": "month"}),
    )
    consumo = forms.DecimalField(label="Consumo", min_value=0, decimal_places=4, max_digits=12, initial=0)

    b_vendita_consumo = forms.DecimalField(label="Vendita consumo", decimal_places=4, max_digits=12, initial=0)
    b_rete_consumi = forms.DecimalField(label="Rete/oneri consumi", decimal_places=4, max_digits=12, initial=0)
    b_vendita_fissa = forms.DecimalField(label="Vendita fissa mensile", decimal_places=4, max_digits=12, initial=0)
    b_rete_fissa = forms.DecimalField(label="Rete/oneri fissa mensile", decimal_places=4, max_digits=12, initial=0)
    b_quota_potenza = forms.DecimalField(label="Quota potenza", decimal_places=4, max_digits=12, initial=0)
    b_sconti = forms.DecimalField(label="Sconti", decimal_places=4, max_digits=12, initial=0)
    b_ricalcoli = forms.DecimalField(label="Ricalcoli/Partite pregresse", decimal_places=4, max_digits=12, initial=0)
    b_bonus_sociale = forms.DecimalField(
        label="Bonus Sociale",
        decimal_places=4,
        max_digits=12,
        initial=0,
        required=False,
        help_text="Inserire l'importo come valore negativo.",
    )
    b_arrotondamenti = forms.DecimalField(label="Arrotondamenti", decimal_places=4, max_digits=12, initial=0)
    b_accise_iva = forms.DecimalField(label="Accise + IVA", decimal_places=4, max_digits=12, initial=0)

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        provider = self._current_value("provider", "ILLUMIA")
        segmento = self._current_value("segmento", "RESIDENZIALE")
        commodity = self._current_value("commodity", "GAS")
        self.fields["offer_var_choice"].choices = self._offer_choices(provider, segmento, commodity, "VARIABILE")
        self.fields["offer_fix_choice"].choices = self._offer_choices(provider, segmento, commodity, "FISSA")

    def _current_value(self, field_name, default):
        if self.data and field_name in self.data:
            return self.data.get(field_name) or default
        return self.initial.get(field_name, self.fields[field_name].initial or default)

    def _offer_choices(self, provider, segmento, commodity, offer_type):
        payload = services.offer_options_payload()
        key = f"{services.normalize_provider(provider)}|{str(segmento).upper()}|{str(commodity).upper()}"
        names = payload.get(key, {}).get(offer_type, [])
        return [("", "Automatica")] + [(name, name) for name in names]

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
        bonus = cleaned.get("b_bonus_sociale")
        cleaned["b_bonus_sociale"] = -abs(bonus) if bonus else 0
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
            else:
                out[key] = str(value)
        return out


def session_to_service_data(raw):
    data = dict(raw)
    data["bill_start"] = services.parse_date_any(data.get("bill_start"))
    data["bill_end"] = services.parse_date_any(data.get("bill_end"))
    for key in ["consumo"] + [f"b_{k}" for k in services.KEYS] + ["ill_sconto_var", "ill_sconto_fix"]:
        data[key] = services.parse_number(data.get(key))
    data["provider"] = services.normalize_provider(data.get("provider", "ILLUMIA"))
    data["offer_var_choice"] = data.get("offer_var_choice", "")
    data["offer_fix_choice"] = data.get("offer_fix_choice", "")
    return data
