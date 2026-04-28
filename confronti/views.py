from django.contrib.auth.decorators import login_required
from django.http import HttpResponse
from django.shortcuts import render

from .forms import ConfrontoForm, session_to_service_data
from .services import build_excel_bytes, prepare_comparison, safe_download_filename


@login_required
def confronto(request):
    prepared = None
    rows = None
    if request.method == "POST":
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
    nome = safe_download_filename(
        f"confronto_illumia_{data.get('nome_cliente', 'Cliente')}_{data.get('commodity', 'ENERGIA')}.xlsx"
    )
    response = HttpResponse(
        content,
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = f'attachment; filename="{nome}"'
    response["Content-Length"] = str(len(content))
    return response
