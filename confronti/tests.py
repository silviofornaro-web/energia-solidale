from datetime import date
from io import BytesIO
from unittest.mock import patch

from django.contrib.auth import get_user_model
from django.core.files.uploadedfile import SimpleUploadedFile
from django.test import Client, SimpleTestCase, TestCase
from openpyxl import load_workbook

from .bill_parser import ParsedBill, parse_bill_text
from .forms import ConfrontoForm
from . import services


def valid_payload(**overrides):
    data = {
        "nome_cliente": "Mario Rossi",
        "pod_pdr": "IT001E12345678",
        "segmento": "RESIDENZIALE",
        "commodity": "EE",
        "bill_tariff_type": "VARIABILE",
        "tariff_selection_mode": "LATEST",
        "provider": "ILLUMIA",
        "providers": ["ILLUMIA"],
        "tax_primary_home": "SI",
        "tax_power_kw": "3",
        "tax_annual_consumption": "1200",
        "tax_region": "Veneto",
        "servizi_accessori_iva": "22%",
        "offer_var_choice": "",
        "offer_fix_choice": "",
        "offer_var_choice_illumia": "",
        "offer_fix_choice_illumia": "",
        "offer_var_choice_eon": "",
        "offer_fix_choice_eon": "",
        "bill_start": "2026-01",
        "bill_end": "2026-03",
        "consumo": "100",
        "bill_fixed_values_are_monthly": "1",
        "b_vendita_consumo": "39.86",
        "b_rete_consumi": "8.47",
        "b_vendita_fissa": "3.70",
        "b_rete_fissa": "0.3967",
        "b_quota_potenza": "12.64",
        "b_sconti": "0",
        "b_ricalcoli": "0",
        "b_bonus_sociale": "21.60",
        "b_arrotondamenti": "0",
        "b_servizi_accessori": "0",
        "bill_offer_expiry": "2026-12-31",
        "b_accise_iva": "12.12",
        "action": "calculate",
    }
    data.update(overrides)
    if "provider" in overrides and "providers" not in overrides:
        data["providers"] = [data["provider"]]
    selected_provider = services.normalize_provider(data["providers"][0]) if data.get("providers") else "ILLUMIA"
    if "offer_var_choice" in overrides:
        data[f"offer_var_choice_{selected_provider.lower()}"] = data["offer_var_choice"]
    if "offer_fix_choice" in overrides:
        data[f"offer_fix_choice_{selected_provider.lower()}"] = data["offer_fix_choice"]
    return data


def service_data(**overrides):
    form = ConfrontoForm(valid_payload(**overrides))
    assert form.is_valid(), form.errors.as_json()
    return form.service_data()


class ServiceUtilityTests(SimpleTestCase):
    def test_parse_number_handles_italian_currency(self):
        self.assertEqual(services.parse_number("€ 1.234,56"), 1234.56)
        self.assertEqual(services.parse_number("-€ 21,60"), -21.6)
        self.assertEqual(services.parse_number(""), 0.0)

    def test_billing_months_and_labels(self):
        self.assertEqual(services.billing_months_from_dates(date(2026, 1, 1), date(2026, 3, 31)), 3)
        self.assertEqual(services.billing_label_from_months(1), "MENSILE")
        self.assertEqual(services.billing_label_from_months(3), "TRIMESTRALE")
        self.assertEqual(services.bill_period_label(date(2026, 1, 1), date(2026, 3, 31)), "Gennaio 2026 - Marzo 2026")
        self.assertTrue(
            services.bill_period_outside_offer_validity(
                date(2026, 1, 1),
                date(2026, 3, 31),
                date(2026, 4, 1),
                date(2026, 4, 30),
            )
        )
        self.assertFalse(
            services.bill_period_outside_offer_validity(
                date(2026, 4, 1),
                date(2026, 4, 30),
                date(2026, 4, 1),
                date(2026, 4, 30),
            )
        )

    def test_select_indice_uses_latest_month_inside_bill_period(self):
        rows = [
            {"mese": "2026-01", "pun": 1, "psv": 10},
            {"mese": "2026-02", "pun": 2, "psv": 20},
            {"mese": "2026-03", "pun": 3, "psv": 30},
            {"mese": "2026-04", "pun": 4, "psv": 40},
        ]
        row, reason = services.select_indice_for_bill_period(rows, date(2026, 1, 1), date(2026, 3, 31))
        self.assertEqual(reason, "period")
        self.assertEqual(row["mese"], "2026-03")

    def test_build_comparison_values_multiplies_monthly_fixed_items_and_copies_bonus(self):
        data = {
            "commodity": "EE",
            "segmento": "RESIDENZIALE",
            "consumo": 100,
            "tax_primary_home": "SI",
            "tax_power_kw": 3,
            "tax_annual_consumption": 1200,
            "tax_region": "Veneto",
            "b_vendita_consumo": 40,
            "b_vendita_fissa": 10,
            "b_rete_consumi": 8,
            "b_rete_fissa": 1,
            "b_quota_potenza": 12,
            "b_sconti": 0,
            "b_ricalcoli": 0,
            "b_bonus_sociale": 21.6,
            "b_arrotondamenti": 0,
            "b_servizi_accessori": 0,
            "b_accise_iva": 12,
            "servizi_accessori_iva": "22%",
            "ill_sconto_var": -3,
            "ill_sconto_fix": -3,
        }
        calc = {
            "billing_months": 3,
            "v_cons": 35,
            "v_fix": 15,
            "f_cons": 30,
            "f_fix": 18,
            "offer_var": "Variabile",
            "offer_fix": "Fissa",
        }
        values = services.build_comparison_values(data, calc)
        self.assertEqual(values["bolletta"]["vendita_fissa"], 30)
        self.assertEqual(values["bolletta"]["rete_fissa"], 3)
        self.assertEqual(values["bolletta"]["bonus_sociale"], -21.6)
        self.assertEqual(values["variabile"]["bonus_sociale"], -21.6)
        self.assertEqual(values["fissa"]["bonus_sociale"], -21.6)
        self.assertAlmostEqual(values["fissa"]["accise"], 0.0, places=6)
        self.assertAlmostEqual(values["fissa"]["iva"], 6.8, places=6)
        self.assertAlmostEqual(values["fissa"]["accise_iva"], 6.8, places=6)

    def test_accessory_services_stay_only_on_bill_column(self):
        data = {
            "commodity": "EE",
            "segmento": "RESIDENZIALE",
            "consumo": 100,
            "tax_primary_home": "SI",
            "tax_power_kw": 3,
            "tax_annual_consumption": 1200,
            "tax_region": "Veneto",
            "b_vendita_consumo": 40,
            "b_vendita_fissa": 10,
            "b_rete_consumi": 8,
            "b_rete_fissa": 1,
            "b_quota_potenza": 12,
            "b_sconti": 0,
            "b_ricalcoli": 0,
            "b_bonus_sociale": 21.6,
            "b_arrotondamenti": 0,
            "b_servizi_accessori": 5,
            "b_accise_iva": 12,
            "servizi_accessori_iva": "10%",
            "ill_sconto_var": -3,
            "ill_sconto_fix": -3,
        }
        calc = {
            "billing_months": 3,
            "v_cons": 35,
            "v_fix": 15,
            "f_cons": 30,
            "f_fix": 18,
            "offer_var": "Variabile",
            "offer_fix": "Fissa",
        }
        values = services.build_comparison_values(data, calc)
        self.assertEqual(values["bolletta"]["servizi_accessori"], 5)
        self.assertEqual(values["variabile"]["servizi_accessori"], 0)
        self.assertEqual(values["fissa"]["servizi_accessori"], 0)
        self.assertEqual(values["servizi_accessori_iva_label"], "10%")
        self.assertAlmostEqual(values["fissa"]["accise_iva"], 6.8, places=6)
        rows = services.build_comparison_table_rows(values)
        servizi_row = next(row for row in rows if row["voce"] == "Servizi accessori (IVA 10%)")
        self.assertEqual(servizi_row["cells"], ["€ 5,00", "€ 0,00", "€ 0,00"])

    def test_gas_annual_consumption_changes_accise_iva(self):
        data = {
            "commodity": "GAS",
            "segmento": "RESIDENZIALE",
            "consumo": 100,
            "tax_primary_home": "SI",
            "tax_power_kw": 0,
            "tax_annual_consumption": 1200,
            "tax_region": "Veneto",
            "b_vendita_consumo": 40,
            "b_vendita_fissa": 10,
            "b_rete_consumi": 8,
            "b_rete_fissa": 1,
            "b_quota_potenza": 0,
            "b_sconti": 0,
            "b_ricalcoli": 0,
            "b_bonus_sociale": 0,
            "b_arrotondamenti": 0,
            "b_servizi_accessori": 0,
            "b_accise_iva": 12,
            "servizi_accessori_iva": "22%",
            "ill_sconto_var": -3,
            "ill_sconto_fix": -3,
        }
        calc = {
            "billing_months": 1,
            "v_cons": 45,
            "v_fix": 6,
            "f_cons": 42,
            "f_fix": 7,
            "offer_var": "Variabile",
            "offer_fix": "Fissa",
        }
        offer_vals = {k: float(data.get(f"b_{k}", 0.0)) for k in services.KEYS}
        offer_vals["bonus_sociale"] = -abs(float(offer_vals.get("bonus_sociale", 0.0)))
        offer_vals["vendita_consumo"] = calc["f_cons"]
        offer_vals["vendita_fissa"] = calc["f_fix"]
        offer_vals["sconti"] = data["ill_sconto_fix"]
        offer_vals["ricalcoli"] = 0.0
        offer_vals["arrotondamenti"] = 0.0
        offer_vals["servizi_accessori"] = 0.0

        high_annual = services.calculate_tax_breakdown(data, calc, offer_vals, "GAS")["accise_iva"]
        low_annual = services.calculate_tax_breakdown(
            {**data, "tax_annual_consumption": 300},
            calc,
            offer_vals,
            "GAS",
        )["accise_iva"]
        self.assertNotEqual(round(high_annual, 6), round(low_annual, 6))

    def test_offer_tax_is_capped_to_bill_percentage_incidence(self):
        data = {
            "commodity": "GAS",
            "segmento": "RESIDENZIALE",
            "consumo": 100,
            "tax_primary_home": "SI",
            "tax_power_kw": 0,
            "tax_annual_consumption": 1200,
            "tax_region": "Veneto",
            "b_vendita_consumo": 40,
            "b_vendita_fissa": 10,
            "b_rete_consumi": 8,
            "b_rete_fissa": 1,
            "b_quota_potenza": 0,
            "b_sconti": 0,
            "b_ricalcoli": 0,
            "b_bonus_sociale": 0,
            "b_arrotondamenti": 0,
            "b_servizi_accessori": 0,
            "b_accise_iva": 5,
            "servizi_accessori_iva": "22%",
            "ill_sconto_var": -3,
            "ill_sconto_fix": -3,
        }
        calc = {
            "billing_months": 1,
            "v_cons": 45,
            "v_fix": 6,
            "f_cons": 42,
            "f_fix": 7,
            "offer_var": "Variabile",
            "offer_fix": "Fissa",
        }

        bill_vals = {k: float(data.get(f"b_{k}", 0.0)) for k in services.KEYS}
        bill_vals["bonus_sociale"] = -abs(float(bill_vals.get("bonus_sociale", 0.0)))
        bill_ratio = services.bill_tax_incidence_ratio(bill_vals, "GAS")

        uncapped_offer_vals = bill_vals.copy()
        uncapped_offer_vals["vendita_consumo"] = calc["f_cons"]
        uncapped_offer_vals["vendita_fissa"] = calc["f_fix"]
        uncapped_offer_vals["sconti"] = data["ill_sconto_fix"]
        uncapped_offer_vals["ricalcoli"] = 0.0
        uncapped_offer_vals["arrotondamenti"] = 0.0
        uncapped_offer_vals["servizi_accessori"] = 0.0
        uncapped_tax = services.calculate_tax_breakdown(data, calc, uncapped_offer_vals, "GAS")

        values = services.build_comparison_values(data, calc)
        capped_vals = values["fissa"]
        capped_total = services.comparison_subtotal(capped_vals, "GAS") * bill_ratio
        scale = capped_total / uncapped_tax["accise_iva"]
        fixed_column = next(column for column in values["offer_columns"] if column["offer_type"] == "FISSA")

        self.assertGreater(uncapped_tax["accise_iva"], capped_total)
        self.assertAlmostEqual(capped_vals["accise_iva"], capped_total, places=6)
        self.assertAlmostEqual(capped_vals["accise"], uncapped_tax["accise"] * scale, places=6)
        self.assertAlmostEqual(capped_vals["iva"], uncapped_tax["iva"] * scale, places=6)
        self.assertAlmostEqual(values["bill_tax_ratio"], bill_ratio, places=6)
        self.assertEqual(values["bill_tax_ratio_label"], "8,47%")
        self.assertTrue(fixed_column["tax_cap_applied"])
        self.assertEqual(fixed_column["tax_cap_status_label"], "Si")
        self.assertAlmostEqual(fixed_column["raw_tax_ratio"], uncapped_tax["accise_iva"] / services.comparison_subtotal(uncapped_offer_vals, "GAS"), places=6)
        self.assertAlmostEqual(fixed_column["tax_ratio"], bill_ratio, places=6)
        self.assertEqual(fixed_column["tax_ratio_label"], "8,47%")

    def test_table_marks_missing_illumia_offers_as_nd(self):
        data = service_data()
        calc = {
            "billing_months": 3,
            "v_cons": 0,
            "v_fix": 0,
            "f_cons": 0,
            "f_fix": 0,
            "offer_var": "",
            "offer_fix": "",
        }
        rows = services.build_comparison_table_rows(services.build_comparison_values(data, calc))
        self.assertTrue(all(row["variabile"] == "N.D." for row in rows))
        self.assertTrue(all(row["fissa"] == "N.D." for row in rows))

    def test_eon_offer_options_include_residential_and_business(self):
        options = services.offer_options_payload()
        self.assertIn("E.ON Flex Gas", options["EON|RESIDENZIALE|GAS"]["VARIABILE"])
        self.assertIn("E.ON Gas Tua", options["EON|RESIDENZIALE|GAS"]["FISSA"])
        self.assertIn("E.ON Gas Impresa CLSC", options["EON|MICROBUSINESS|GAS"]["VARIABILE"])
        self.assertIn("E.ON LuceDinamica ECO CLSE", options["EON|MICROBUSINESS|EE"]["VARIABILE"])
        self.assertIn("E.ON Profilo Dinamico Gas P", options["EON|BUSINESS|GAS"]["VARIABILE"])
        self.assertIn("E.ON Profilo Sicuro T", options["EON|BUSINESS|EE"]["FISSA"])

    def test_missing_microbusiness_tariffe_falls_back_to_business(self):
        latest = services.load_tariffe_file_for_segment("MICROBUSINESS", "ILLUMIA", "LATEST", "2026-05")
        self.assertIn("/business/", str(latest))

        options = services.offer_options_payload()
        self.assertIn("GAS BUSINESS PREMIUM FLEX", options["ILLUMIA|MICROBUSINESS|GAS"]["VARIABILE"])

        prepared = services.prepare_comparison(
            service_data(
                segmento="MICROBUSINESS",
                provider="ILLUMIA",
                providers=["ILLUMIA"],
                commodity="GAS",
                tax_power_kw="0",
                b_quota_potenza="0",
            )
        )
        self.assertIn("/business/", prepared["calc"]["offer_file"])
        self.assertEqual(prepared["calc"]["provider_results"][0]["tariff_segment"], "BUSINESS")
        self.assertEqual(prepared["calc"]["offer_var"], "GAS PREMIUM FLEX BUSINESS")

    def test_tariff_selection_can_use_bill_period_month(self):
        latest = services.load_tariffe_file_for_segment("RESIDENZIALE", "ILLUMIA", "LATEST", "2026-03")
        period = services.load_tariffe_file_for_segment("RESIDENZIALE", "ILLUMIA", "PERIOD", "2026-03")
        self.assertIn("2026-05", str(latest))
        self.assertIn("2026-03", str(period))
        self.assertIsNone(services.load_tariffe_file_for_segment("RESIDENZIALE", "EON", "PERIOD", "2026-03"))

    def test_prepare_comparison_uses_period_tariffe_when_selected(self):
        latest = services.prepare_comparison(service_data(tariff_selection_mode="LATEST"))
        period = services.prepare_comparison(service_data(tariff_selection_mode="PERIOD"))
        self.assertIn("2026-05", latest["calc"]["offer_file"])
        self.assertIn("2026-03", period["calc"]["offer_file"])
        self.assertEqual(period["calc"]["tariff_selection_mode_label"], "Tariffe del periodo bolletta")
        self.assertEqual(period["calc"]["tariff_target_month"], "2026-03")

    def test_prepare_comparison_uses_selected_eon_fixed_gas_offer(self):
        data = service_data(
            provider="EON",
            commodity="GAS",
            offer_var_choice="E.ON Flex Gas",
            offer_fix_choice="E.ON Gas Tua",
        )
        prepared = services.prepare_comparison(data)
        self.assertEqual(prepared["calc"]["provider_label"], "E.ON")
        self.assertEqual(prepared["calc"]["offer_var"], "E.ON Flex Gas")
        self.assertEqual(prepared["calc"]["offer_fix"], "E.ON Gas Tua")
        self.assertEqual(prepared["calc"]["offer_valid_to"], date(2026, 5, 21))
        self.assertEqual(prepared["calc"]["bill_offer_expiry"], date(2026, 12, 31))
        self.assertEqual(prepared["calc"]["bill_offer_expiry_label"], "31/12/2026")
        self.assertEqual(prepared["calc"]["offer_expiry_label"], "31/12/2026")
        self.assertEqual(prepared["values"]["variabile"]["sconti"], -10)
        self.assertEqual(prepared["values"]["fissa"]["sconti"], -10)
        self.assertAlmostEqual(prepared["values"]["fissa"]["vendita_consumo"], 54.8, places=4)


class BillParserTests(SimpleTestCase):
    def test_parses_gas_bill_with_textual_months(self):
        parsed = parse_bill_text(
            """
            GAS NATURALE
            MERCATO LIBERO
            MASSIMO PASSARELLA
            VIA SIGNORIA 79
            Periodo oggetto di fatturazione: 1 ottobre 2025 - 31 ottobre 2025
            Consumo totale fatturato: 21,41202 Smc
            Consumo da inizio contratto (mc): 564
            Scadenza condizioni economiche: 30/09/2027
            SCONTRINO DELL'ENERGIA
            Codice PDR: 15351410010036
            QUOTA PER CONSUMI
            21,412020 Smc 0,714552 €/Smc 15,30
            di cui spesa per vendita gas naturale 0,487110 €/Smc 10,43
            di cui spesa per la rete e gli oneri generali di sistema 0,227442 €/Smc 4,87
            QUOTA FISSA
            1 mesi 16,330000 €/mese 16,33
            di cui spesa per vendita gas naturale 9,000000 €/mese 9,00
            di cui spesa per la rete e gli oneri generali di sistema 7,330000 €/mese 7,33
            Accise e IVA 12,07
            """
        )
        self.assertEqual(parsed.values["commodity"], "GAS")
        self.assertEqual(parsed.values["pod_pdr"], "15351410010036")
        self.assertEqual(parsed.values["bill_start"], date(2025, 10, 1))
        self.assertEqual(parsed.values["bill_end"], date(2025, 10, 31))
        self.assertAlmostEqual(parsed.values["b_rete_fissa"], 7.33)


class ConfrontoFormTests(SimpleTestCase):
    def test_dashboard_context_fields_are_required_and_not_prefilled(self):
        form = ConfrontoForm()
        self.assertIsNone(form.fields["nome_cliente"].initial)
        self.assertFalse(form.fields["pod_pdr"].required)
        self.assertEqual(form.fields["segmento"].choices[0], ("", "Seleziona segmento"))
        self.assertEqual(form.fields["commodity"].choices[0], ("", "Seleziona fornitura"))
        self.assertEqual(form.fields["bill_tariff_type"].choices[0], ("", "Seleziona tariffa"))
        self.assertEqual(form.fields["providers"].choices[0], ("ILLUMIA", "Illumia"))
        self.assertIsNone(form.fields["tariff_selection_mode"].initial)
        self.assertIsNone(form.fields["bill_start"].initial)
        self.assertIsNone(form.fields["bill_end"].initial)
        self.assertIsNone(form.fields["bill_offer_expiry"].initial)
        self.assertIsNone(form.fields["tax_annual_consumption"].initial)

        invalid = ConfrontoForm(
            valid_payload(
                nome_cliente="",
                segmento="",
                commodity="",
                bill_tariff_type="",
                tariff_selection_mode="",
                providers=[],
                tax_annual_consumption="",
                bill_start="",
                bill_end="",
                bill_offer_expiry="",
            )
        )
        self.assertFalse(invalid.is_valid())
        for field in [
            "nome_cliente",
            "segmento",
            "commodity",
            "bill_tariff_type",
            "tariff_selection_mode",
            "providers",
            "bill_start",
            "bill_offer_expiry",
            "tax_annual_consumption",
        ]:
            self.assertIn(field, invalid.errors)

    def test_month_fields_are_normalized_to_full_bill_period(self):
        form = ConfrontoForm(valid_payload())
        self.assertTrue(form.is_valid(), form.errors.as_data())
        self.assertEqual(form.cleaned_data["bill_start"], date(2026, 1, 1))
        self.assertEqual(form.cleaned_data["bill_end"], date(2026, 3, 31))
        self.assertEqual(form.cleaned_data["bill_offer_expiry"], date(2026, 12, 31))
        self.assertEqual(form.cleaned_data["tariff_selection_mode"], "LATEST")
        self.assertEqual(float(form.cleaned_data["b_bonus_sociale"]), -21.6)

    def test_missing_end_month_defaults_to_start_month(self):
        form = ConfrontoForm(valid_payload(bill_start="2026-02", bill_end=""))
        self.assertTrue(form.is_valid(), form.errors.as_data())
        self.assertEqual(form.cleaned_data["bill_start"], date(2026, 2, 1))
        self.assertEqual(form.cleaned_data["bill_end"], date(2026, 2, 28))

    def test_bonus_sociale_is_optional(self):
        form = ConfrontoForm(valid_payload(b_bonus_sociale=""))
        self.assertTrue(form.is_valid(), form.errors.as_data())
        self.assertEqual(form.cleaned_data["b_bonus_sociale"], 0)

    def test_italian_decimal_commas_are_accepted(self):
        form = ConfrontoForm(
            valid_payload(
                b_vendita_consumo="85,520",
                b_rete_consumi="24,61",
                b_vendita_fissa="12,105",
                b_rete_fissa="1,9",
                b_accise_iva="23,97",
            )
        )
        self.assertTrue(form.is_valid(), form.errors.as_data())
        self.assertEqual(float(form.cleaned_data["b_vendita_consumo"]), 85.52)
        self.assertEqual(float(form.cleaned_data["b_rete_consumi"]), 24.61)

    def test_service_data_normalizes_fixed_bill_amounts_to_monthly_values(self):
        form = ConfrontoForm(
            valid_payload(
                bill_fixed_values_are_monthly="0",
                bill_start="2026-01",
                bill_end="2026-02",
                b_vendita_fissa="10",
                b_rete_fissa="4",
            )
        )
        self.assertTrue(form.is_valid(), form.errors.as_data())
        data = form.service_data()
        self.assertEqual(data["bill_fixed_values_are_monthly"], "1")
        self.assertEqual(data["b_vendita_fissa"], 5.0)
        self.assertEqual(data["b_rete_fissa"], 2.0)

    def test_service_data_keeps_already_monthly_fixed_bill_amounts(self):
        form = ConfrontoForm(
            valid_payload(
                bill_fixed_values_are_monthly="1",
                bill_start="2026-01",
                bill_end="2026-02",
                b_vendita_fissa="5",
                b_rete_fissa="2",
            )
        )
        self.assertTrue(form.is_valid(), form.errors.as_data())
        data = form.service_data()
        self.assertEqual(data["bill_fixed_values_are_monthly"], "1")
        self.assertEqual(data["b_vendita_fissa"], 5)
        self.assertEqual(data["b_rete_fissa"], 2)

    def test_gas_disables_and_zeros_power_fields(self):
        form = ConfrontoForm(
            valid_payload(
                commodity="GAS",
                tax_power_kw="9",
                b_quota_potenza="99",
            )
        )
        self.assertTrue(form.fields["tax_power_kw"].disabled)
        self.assertTrue(form.fields["b_quota_potenza"].disabled)
        self.assertTrue(form.is_valid(), form.errors.as_data())
        self.assertEqual(form.cleaned_data["tax_power_kw"], 0)
        self.assertEqual(form.cleaned_data["b_quota_potenza"], 0)

    def test_end_month_before_start_month_is_rejected(self):
        form = ConfrontoForm(valid_payload(bill_start="2026-03", bill_end="2026-01"))
        self.assertFalse(form.is_valid())
        self.assertIn("Il mese finale non può essere precedente al mese iniziale.", form.non_field_errors())


class ConfrontoViewTests(TestCase):
    def setUp(self):
        self.user = get_user_model().objects.create_user(username="tester", password="secret")
        self.client = Client()

    def login(self):
        self.assertTrue(self.client.login(username="tester", password="secret"))

    def test_login_is_required(self):
        response = self.client.get("/")
        self.assertEqual(response.status_code, 302)
        self.assertIn("/accounts/login/", response["Location"])

    def test_form_renders_month_selectors_and_reset_button(self):
        self.login()
        response = self.client.get("/")
        self.assertNotContains(response, 'type="month"')
        self.assertContains(response, "Build locale CAP-FISCALE 2026-06-06")
        self.assertContains(response, 'data-month-field="bill_start"')
        self.assertContains(response, 'data-month-field="bill_end"')
        self.assertContains(response, 'data-month-part="year"')
        self.assertContains(response, 'type="date"')
        self.assertContains(response, "Nuova bolletta")
        self.assertContains(response, "Fornitori confronto")
        self.assertContains(response, "Importa bolletta PDF")

    def test_invalid_form_renders_field_errors(self):
        self.login()
        response = self.client.post("/", valid_payload(tax_power_kw="0"))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Il confronto non è partito")
        self.assertContains(response, "Indica la potenza impegnata")
        self.assertNotContains(response, "Scarica Excel")

    def test_calculate_comparison_stores_session_and_renders_period(self):
        self.login()
        response = self.client.post("/", valid_payload())
        self.assertEqual(response.status_code, 200)
        html = response.content.decode()
        self.assertIn("Gennaio 2026 - Marzo 2026", html)
        self.assertIn("Cliente:</strong> Mario Rossi", html)
        self.assertIn("Codice POD/PDR:</strong> IT001E12345678", html)
        self.assertIn("Confronto eseguito:</strong>", html)
        self.assertIn("Tipo tariffa bolletta:</strong> Variabile", html)
        self.assertIn("Consumo annuo stimato:</strong> 1200 kWh/anno", html)
        self.assertIn("Scadenza offerta bolletta:</strong> 31/12/2026", html)
        self.assertIn("Logica tariffe:</strong> Ultime tariffe disponibili", html)
        self.assertNotIn("periodo bolletta NON rientra", html)
        self.assertIn("TRIMESTRALE", html)
        self.assertIn('data-download-excel', html)
        self.assertIn("Scarica Excel", html)
        self.assertIn("last_confronto", self.client.session)
        self.assertEqual(self.client.session["last_confronto"]["bill_start"], "2026-01-01")
        self.assertEqual(self.client.session["last_confronto"]["bill_end"], "2026-03-31")
        self.assertEqual(self.client.session["last_confronto"]["bill_offer_expiry"], "2026-12-31")
        self.assertEqual(self.client.session["last_confronto"]["tariff_selection_mode"], "LATEST")

    def test_calculate_normalizes_multi_month_fixed_bill_fields_in_rendered_form(self):
        self.login()
        response = self.client.post(
            "/",
            valid_payload(
                bill_fixed_values_are_monthly="0",
                bill_start="2026-01",
                bill_end="2026-02",
                b_vendita_fissa="10",
                b_rete_fissa="4",
            ),
        )
        self.assertEqual(response.status_code, 200)
        form = response.context["form"]
        self.assertEqual(form["bill_fixed_values_are_monthly"].value(), "1")
        self.assertEqual(form["b_vendita_fissa"].value(), 5.0)
        self.assertEqual(form["b_rete_fissa"].value(), 2.0)

    def test_calculate_renders_tax_cap_summary_for_capped_offer(self):
        self.login()
        response = self.client.post(
            "/",
            valid_payload(
                commodity="GAS",
                bill_tariff_type="VARIABILE",
                tax_power_kw="0",
                tax_annual_consumption="1200",
                consumo="100",
                b_vendita_consumo="40",
                b_vendita_fissa="10",
                b_rete_consumi="8",
                b_rete_fissa="1",
                b_quota_potenza="0",
                b_sconti="0",
                b_ricalcoli="0",
                b_bonus_sociale="0",
                b_arrotondamenti="0",
                b_servizi_accessori="0",
                b_accise_iva="5",
            ),
        )
        self.assertEqual(response.status_code, 200)
        html = response.content.decode()
        self.assertIn("Incidenza fiscale bolletta:</strong> 6,17%", html)
        self.assertIn("Cap fiscale applicato:</strong> Si", html)
        self.assertIn("Incidenza teorica senza cap:</strong>", html)

    def test_calculate_can_compare_selected_eon_offer(self):
        self.login()
        response = self.client.post(
            "/",
            valid_payload(
                provider="EON",
                offer_var_choice="E.ON Flex Luce Casa",
                offer_fix_choice="E.ON Luce Tua",
            ),
        )
        self.assertEqual(response.status_code, 200)
        html = response.content.decode()
        self.assertIn("Fornitori confronto:</strong> E.ON", html)
        self.assertIn("E.ON Flex Luce Casa", html)
        self.assertIn("E.ON Luce Tua", html)
        self.assertIn("E.ON Variabile", html)
        self.assertEqual(self.client.session["last_confronto"]["provider"], "EON")

    def test_calculate_can_compare_illumia_and_eon_together(self):
        self.login()
        response = self.client.post(
            "/",
            valid_payload(
                providers=["ILLUMIA", "EON"],
                offer_var_choice_eon="E.ON Flex Luce Casa",
                offer_fix_choice_eon="E.ON Luce Tua",
            ),
        )
        self.assertEqual(response.status_code, 200)
        html = response.content.decode()
        self.assertIn("Fornitori confronto:</strong> Illumia + E.ON", html)
        self.assertIn("Illumia Variabile", html)
        self.assertIn("Illumia Fissa", html)
        self.assertIn("E.ON Variabile", html)
        self.assertIn("E.ON Fissa", html)
        self.assertEqual(response.context["prepared"]["columns"][0]["label"], "Bolletta")
        self.assertEqual(len(response.context["prepared"]["columns"]), 5)
        self.assertEqual(self.client.session["last_confronto"]["providers"], "ILLUMIA,EON")

    def test_change_bill_resets_bill_values_and_download_session(self):
        self.login()
        self.client.post("/", valid_payload())
        self.assertIn("last_confronto", self.client.session)
        session = self.client.session
        session["last_uploaded_bill_name"] = "bolletta.pdf"
        session.save()
        response = self.client.post(
            "/",
            valid_payload(action="reset_bill", b_vendita_consumo="999", tax_annual_consumption="9999"),
        )
        html = response.content.decode()
        self.assertEqual(response.status_code, 200)
        self.assertNotIn("last_confronto", self.client.session)
        self.assertNotIn("Scarica Excel", html)
        self.assertIn("Mario Rossi", html)
        self.assertNotIn("999", html)
        self.assertEqual(response.context["form"].initial["tax_annual_consumption"], "")
        self.assertNotIn("last_uploaded_bill_name", self.client.session)
        self.assertEqual(response.context["uploaded_bill_name"], "")

    @patch("confronti.views.parse_uploaded_bill")
    def test_pdf_upload_prefills_recognized_bill_values(self, mock_parse_uploaded_bill):
        mock_parse_uploaded_bill.return_value = ParsedBill(
            values={
                "nome_cliente": "Federico Boetto",
                "pod_pdr": "00881906523889",
                "commodity": "GAS",
                "consumo": 83,
                "tax_annual_consumption": 83,
                "b_vendita_consumo": 37.81,
            },
            warnings=["Data fine offerta bolletta non riconosciuta: compilala manualmente."],
        )
        self.login()
        response = self.client.post(
            "/",
            {
                "action": "extract_bill",
                "bill_pdf": SimpleUploadedFile("bolletta.pdf", b"%PDF-test", content_type="application/pdf"),
            },
        )
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Ho riportato nella dashboard 6 valori riconosciuti")
        self.assertContains(response, "Federico Boetto")
        self.assertContains(response, "00881906523889")
        self.assertContains(response, "Data fine offerta bolletta non riconosciuta")
        self.assertContains(response, "File caricato: bolletta.pdf")
        self.assertEqual(response.context["uploaded_bill_name"], "bolletta.pdf")
        self.assertEqual(self.client.session["last_uploaded_bill_name"], "bolletta.pdf")

    def test_download_requires_a_previous_comparison(self):
        self.login()
        response = self.client.get("/scarica-excel/")
        self.assertEqual(response.status_code, 400)

    def test_excel_download_contains_expected_labels_values_and_formulas(self):
        self.login()
        self.client.post("/", valid_payload(b_servizi_accessori="5", servizi_accessori_iva="10%"))
        response = self.client.get("/scarica-excel/")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            response["Content-Type"],
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        self.assertIn("confronto_illumia_Mario_Rossi_EE.xlsx", response["Content-Disposition"])
        self.assertEqual(response.content[:2], b"PK")

        wb = load_workbook(BytesIO(response.content), data_only=False)
        ws = wb["Confronto"]
        self.assertEqual(ws["A1"].value, "Mario Rossi")
        self.assertEqual(ws["F3"].value, "Scadenza offerta bolletta: 31/12/2026")
        self.assertTrue(ws["F3"].font.bold)
        self.assertEqual(ws["F3"].font.color.rgb[-6:], "B3261E")
        self.assertNotIn("scadenza", ws["F2"].value.lower())
        self.assertEqual(ws["F6"].value, "Tipo tariffa bolletta: Variabile")
        self.assertTrue(str(ws["F7"].value).startswith("Confronto eseguito: "))
        self.assertEqual(ws["F8"].value, "IVA servizi accessori: 10%")
        self.assertEqual(ws["F9"].value, "Consumo annuo stimato: 1200 kWh/anno")
        self.assertTrue(str(ws["F10"].value).startswith("Parametri Accise/IVA: "))
        self.assertEqual(ws["F11"].value, "Logica tariffe: Ultime tariffe disponibili")
        self.assertEqual(ws["F12"].value, "Codice POD/PDR: IT001E12345678")
        self.assertEqual(ws["F13"].value, "Fornitura: Luce")
        self.assertEqual(ws["F14"].value, "Periodo bolletta: Gennaio 2026 - Marzo 2026")
        self.assertEqual(ws["A12"].value, "Bonus Sociale")
        self.assertEqual(ws["A13"].value, "Arrotondamenti")
        self.assertEqual(ws["A14"].value, "Servizi accessori (IVA 10%)")
        self.assertEqual(ws["A15"].value, "Accise")
        self.assertEqual(ws["A16"].value, "IVA")
        self.assertEqual(ws["A17"].value, "Accise e Iva")
        self.assertEqual(ws["A18"].value, "Totale")
        self.assertEqual(ws["B12"].value, -21.6)
        self.assertEqual(ws["C12"].value, -21.6)
        self.assertEqual(ws["D12"].value, -21.6)
        self.assertEqual(ws["B14"].value, 5)
        self.assertEqual(ws["B15"].value, "N.D.")
        self.assertEqual(ws["B16"].value, "N.D.")
        self.assertEqual(ws["C17"].value, ws["C15"].value + ws["C16"].value)
        self.assertEqual(ws["C14"].value, 0)
        self.assertEqual(ws["D14"].value, 0)
        self.assertAlmostEqual(ws["B6"].value, 11.1, places=4)
        self.assertAlmostEqual(ws["B8"].value, 1.1901, places=4)
        self.assertIsInstance(ws["C15"].value, (int, float))
        self.assertIsInstance(ws["D15"].value, (int, float))
        self.assertEqual(ws["B17"].value, 12.12)
        self.assertEqual(ws["B18"].value, "=SUM(B4:B14)+B17")

    def test_excel_download_contains_both_provider_columns(self):
        self.login()
        self.client.post(
            "/",
            valid_payload(
                providers=["ILLUMIA", "EON"],
                offer_var_choice_eon="E.ON Flex Luce Casa",
                offer_fix_choice_eon="E.ON Luce Tua",
            ),
        )
        response = self.client.get("/scarica-excel/")
        self.assertEqual(response.status_code, 200)
        self.assertIn("confronto_illumia_eon_Mario_Rossi_EE.xlsx", response["Content-Disposition"])

        wb = load_workbook(BytesIO(response.content), data_only=False)
        ws = wb["Confronto"]
        self.assertEqual(ws["B3"].value, "Bolletta")
        self.assertEqual(ws["C3"].value, "Illumia Variabile")
        self.assertEqual(ws["D3"].value, "Illumia Fissa")
        self.assertEqual(ws["E3"].value, "E.ON Variabile")
        self.assertEqual(ws["F3"].value, "E.ON Fissa")
        self.assertEqual(ws["H1"].value, "Fornitori confronto: Illumia + E.ON")
        self.assertIn("Illumia:", ws["H2"].value)
        self.assertIn("E.ON:", ws["H2"].value)
        self.assertEqual(ws["F18"].value, "=SUM(F4:F14)+F17")

    @patch("confronti.services.load_tariffe_file_for_segment_with_effective_segment", return_value=(None, "RESIDENZIALE"))
    def test_missing_illumia_offer_keeps_bill_and_marks_offers_nd(self, _mock_load_file):
        self.login()
        response = self.client.post("/", valid_payload())
        rows = response.context["rows"]
        self.assertTrue(all(row["variabile"] == "N.D." for row in rows))
        self.assertTrue(all(row["fissa"] == "N.D." for row in rows))
        self.assertNotEqual(next(row for row in rows if row["voce"] == "Totale")["bolletta"], "N.D.")
        self.assertNotEqual(next(row for row in rows if row["voce"] == "Accise e Iva")["bolletta"], "N.D.")
