from datetime import date

from django import forms

from . import services


class ConfrontoForm(forms.Form):
    SEGMENTI = [("RESIDENZIALE", "Residenziale"), ("BUSINESS", "Business")]
    COMMODITIES = [("GAS", "Gas"), ("EE", "Luce")]

    nome_cliente = forms.CharField(label="Nome e Cognome", max_length=120, initial="Cliente")
    segmento = forms.ChoiceField(label="Segmento", choices=SEGMENTI, initial="RESIDENZIALE")
    commodity = forms.ChoiceField(label="Fornitura", choices=COMMODITIES, initial="GAS")
    bill_start = forms.DateField(label="Dal", initial=date.today, widget=forms.DateInput(attrs={"type": "date"}))
    bill_end = forms.DateField(label="Al", initial=date.today, widget=forms.DateInput(attrs={"type": "date"}))
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

    def clean(self):
        cleaned = super().clean()
        if cleaned.get("bill_start") and cleaned.get("bill_end") and cleaned["bill_end"] < cleaned["bill_start"]:
            raise forms.ValidationError("La data finale non può essere precedente alla data iniziale.")
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
    return data
