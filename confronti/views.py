from django.contrib.auth.decorators import login_required
from django.http import HttpResponse
from django.shortcuts import render

from .forms import ConfrontoForm, session_to_service_data
from .services import build_excel_bytes, offer_options_payload, prepare_comparison, provider_label, safe_download_filename


@login_required
def confronto(request):
    prepared = None
    rows = None
    if request.method == "POST":
        if request.POST.get("action") == "reset_bill":
            request.session.pop("last_confronto", None)
            form = ConfrontoForm(
                initial={
                    "nome_cliente": request.POST.get("nome_cliente") or "Cliente",
                    "segmento": request.POST.get("segmento") or "RESIDENZIALE",
                    "commodity": request.POST.get("commodity") or "GAS",
                    "provider": request.POST.get("provider") or "ILLUMIA",
                    "offer_var_choice": request.POST.get("offer_var_choice") or "",
                    "offer_fix_choice": request.POST.get("offer_fix_choice") or "",
                }
            )
        else:
            form = ConfrontoForm(request.POST)
            if form.is_valid():
                data = form.service_data()
                prepared = prepare_comparison(data)
                rows = prepared["rows"]
                request.session["last_confronto"] = form.session_data()
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
    provider_name = provider_label(data.get("provider", "ILLUMIA")).lower().replace(".", "")
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
