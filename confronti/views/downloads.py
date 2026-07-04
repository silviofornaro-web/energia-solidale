import os

from django.contrib.auth.decorators import login_required
from django.http import FileResponse, HttpResponse
from django.shortcuts import get_object_or_404, redirect

from ..models import ComparisonReport
from ..roles import is_illumia_operator, is_internal_user
from ..services import (
    build_excel_bytes,
    build_reports_summary_excel,
    prepare_comparison,
    provider_label,
    safe_download_filename,
    session_to_service_data,
)
from .helpers import (
    LAST_COMPARISON_KEY,
    LAST_REPORT_SUMMARY_KEY,
    force_customer_mode_data,
    force_operator_mode_data,
    is_archive_admin,
    mode_config,
    redirect_for_non_archive_admin,
)


@login_required
def scarica_excel(request):
    if not is_internal_user(request.user):
        return redirect("accesso_clienti")
    return _scarica_excel(request, operator_mode=is_illumia_operator(request.user))


@login_required
def scarica_excel_cliente_illumia(request):
    return _scarica_excel(request, customer_mode=True)


def _scarica_excel(request, customer_mode=False, operator_mode=False):
    mode = mode_config(customer_mode, operator_mode)
    raw = request.session.get(mode["comparison_session_key"])
    if not raw:
        return HttpResponse("Nessun confronto pronto. Torna alla pagina principale e calcola il confronto.", status=400)
    data = prepare_session_data(raw, customer_mode, operator_mode)
    prepared = prepare_comparison(data)
    content = build_excel_bytes(data, prepared)
    provider_name = "_".join(provider_label(p).lower().replace(".", "") for p in data.get("providers", []))
    provider_name = provider_name or provider_label(data.get("provider", "ILLUMIA")).lower().replace(".", "")
    nome = safe_download_filename(
        f"confronto_{provider_name}_{data.get('nome_cliente', 'Cliente')}_{data.get('commodity', 'ENERGIA')}.xlsx"
    )
    response = HttpResponse(content, content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response["Content-Disposition"] = f'attachment; filename="{nome}"'
    response["Content-Length"] = str(len(content))
    return response


def prepare_session_data(raw, customer_mode=False, operator_mode=False):
    from ..services import session_to_service_data
    data = session_to_service_data(raw)
    if customer_mode:
        data = force_customer_mode_data(data)
    elif operator_mode:
        data = force_operator_mode_data(data)
    return data


@login_required
def scarica_report_archiviato(request, report_id):
    if not is_archive_admin(request.user):
        return redirect_for_non_archive_admin(request)
    report = get_object_or_404(ComparisonReport, pk=report_id)
    filename = report.original_filename or os.path.basename(report.report_file.name)
    report.report_file.open("rb")
    response = FileResponse(
        report.report_file,
        as_attachment=True,
        filename=filename,
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Length"] = report.report_file.size
    return response


@login_required
def scarica_sunto_report(request):
    if not is_internal_user(request.user):
        return redirect("accesso_clienti")
    summary = request.session.get(LAST_REPORT_SUMMARY_KEY)
    if not summary or not summary.get("rows"):
        return HttpResponse("Nessun sunto pronto. Carica i report dalla dashboard admin.", status=400)
    content = build_reports_summary_excel(summary)
    nome = safe_download_filename("sunto_report_confronti.xlsx")
    response = HttpResponse(content, content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response["Content-Disposition"] = f'attachment; filename="{nome}"'
    response["Content-Length"] = str(len(content))
    return response
