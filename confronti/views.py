import logging

from django.contrib.auth.decorators import login_required
from django.http import HttpResponse
from django.shortcuts import render
from django.utils import timezone

from .bill_parser import parse_uploaded_bill
from .forms import BillUploadForm, ConfrontoForm, session_to_service_data
from .services import (
    archive_excel_bytes,
    build_excel_bytes,
    google_drive_archive_configured,
    offer_options_payload,
    prepare_comparison,
    provider_label,
    safe_download_filename,
    upload_excel_to_google_drive,
)


logger = logging.getLogger(__name__)


@login_required
def confronto(request):
    prepared = None
    rows = None
    extraction_warnings = []
    extraction_count = 0
    upload_form = BillUploadForm()
    if request.method == "POST":
        action = request.POST.get("action")
        if action == "extract_bill":
            request.session.pop("last_confronto", None)
            upload_form = BillUploadForm(request.POST, request.FILES)
            if upload_form.is_valid():
                try:
                    parsed = parse_uploaded_bill(upload_form.cleaned_data["bill_pdf"])
                except Exception:
                    logger.exception("Impossibile leggere il PDF della bolletta.")
                    form = ConfrontoForm()
                    extraction_warnings = ["Non sono riuscito a leggere questo PDF. Inserisci manualmente i valori della bolletta."]
                else:
                    form = ConfrontoForm(initial=parsed.values)
                    extraction_warnings = parsed.warnings
                    extraction_count = len(parsed.values)
                    upload_form = BillUploadForm()
            else:
                form = ConfrontoForm()
        elif action == "reset_bill":
            request.session.pop("last_confronto", None)
            providers = request.POST.getlist("providers") or [request.POST.get("provider") or ""]
            form = ConfrontoForm(
                initial={
                    "nome_cliente": request.POST.get("nome_cliente") or "",
                    "segmento": request.POST.get("segmento") or "",
                    "commodity": request.POST.get("commodity") or "",
                    "bill_tariff_type": request.POST.get("bill_tariff_type") or "",
                    "tariff_selection_mode": request.POST.get("tariff_selection_mode") or "",
                    "providers": [provider for provider in providers if provider],
                    "tax_primary_home": request.POST.get("tax_primary_home") or "SI",
                    "tax_power_kw": request.POST.get("tax_power_kw") or "0",
                    "tax_annual_consumption": "",
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
            "upload_form": upload_form,
            "extraction_warnings": extraction_warnings,
            "extraction_count": extraction_count,
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
    try:
        archive_excel_bytes(content, data)
    except OSError:
        logger.exception("Impossibile archiviare il file Excel del confronto.")
    if google_drive_archive_configured():
        try:
            upload_excel_to_google_drive(content, data)
        except Exception:
            logger.exception("Impossibile archiviare il file Excel del confronto su Google Drive.")
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
