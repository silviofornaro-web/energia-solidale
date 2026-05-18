from django.contrib.auth.decorators import login_required
from django.http import HttpResponse
from django.shortcuts import render
from django.utils import timezone

from .forms import ConfrontoForm, session_to_service_data
from .services import build_excel_bytes, offer_options_payload, prepare_comparison, provider_label, safe_download_filename


@login_required
def confronto(request):
    prepared = None
    rows = None
    if request.method == "POST":
        if request.POST.get("action") == "reset_bill":
            request.session.pop("last_confronto", None)
            providers = request.POST.getlist("providers") or [request.POST.get("provider") or ""]
            form = ConfrontoForm(
                initial={
                    "nome_cliente": request.POST.get("nome_cliente") or "",
                    "segmento": request.POST.get("segmento") or "",
                    "commodity": request.POST.get("commodity") or "",
                    "bill_tariff_type": request.POST.get("bill_tariff_type") or "",
                    "providers": [provider for provider in providers if provider],
                    "tax_primary_home": request.POST.get("tax_primary_home") or "SI",
                    "tax_power_kw": request.POST.get("tax_power_kw") or "0",
                    "tax_annual_consumption": request.POST.get("tax_annual_consumption") or "",
                    "tax_region": request.POST.get("tax_region") or "Veneto",
                    "servizi_accessori_iva": "22%",
                    "offer_var_choice_illumia": request.POST.get("offer_var_choice_illumia") or "",
                    "offer_fix_choice_illumia": request.POST.get("offer_fix_choice_illumia") or "",
                    "offer_var_choice_eon": request.POST.get("offer_var_choice_eon") or "",
                    "offer_fix_choice_eon": request.POST.get("offer_fix_choice_eon") or "",
                }
            )
        else:
            form = ConfrontoForm(request.POST)
            if form.is_valid():
                data = form.service_data()
                data["comparison_datetime"] = timezone.localtime().isoformat(timespec="seconds")
                prepared = prepare_comparison(data)
                rows = prepared["rows"]
                session_data = form.session_data()
                session_data["comparison_datetime"] = data["comparison_datetime"]
                request.session["last_confronto"] = session_data
    else:
        form = ConfrontoForm()

    return render(
        request,
        "confronti/confronto.html",
        {
            "form": form,
            "prepared": prepared,
            "rows": rows,
            "offer_options": offer_options_payload(),
        },
    )


@login_required
def scarica_excel(request):
    raw = request.session.get("last_confronto")
    if not raw:
        return HttpResponse("Nessun confronto pronto. Torna alla pagina principale e calcola il confronto.", status=400)
    data = session_to_service_data(raw)
    prepared = prepare_comparison(data)
    content = build_excel_bytes(data, prepared)
    provider_name = "_".join(provider_label(provider).lower().replace(".", "") for provider in data.get("providers", []))
    provider_name = provider_name or provider_label(data.get("provider", "ILLUMIA")).lower().replace(".", "")
    nome = safe_download_filename(
        f"confronto_{provider_name}_{data.get('nome_cliente', 'Cliente')}_{data.get('commodity', 'ENERGIA')}.xlsx"
    )
    response = HttpResponse(
        content,
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = f'attachment; filename="{nome}"'
    response["Content-Length"] = str(len(content))
    return response
