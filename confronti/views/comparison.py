import logging
import os

from django.contrib import messages
from django.contrib.auth import get_user_model
from django.contrib.auth.decorators import login_required
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse
from django.utils import timezone

from ..bill_parser import parse_uploaded_bill
from ..forms import BillUploadForm, ConfrontoForm, CustomerInviteForm, ReportSummaryUploadForm
from ..models import InviteCode
from ..roles import is_illumia_operator, is_internal_user
from ..services import build_reports_summary, offer_options_payload, prepare_comparison, provider_label, safe_download_filename
from .helpers import (
    LAST_ARCHIVED_REPORT_CLIENT_KEY,
    LAST_ARCHIVED_REPORT_KEY,
    LAST_COMPARISON_CLIENT_KEY,
    LAST_COMPARISON_KEY,
    LAST_REPORT_SUMMARY_KEY,
    LAST_UPLOADED_BILL_NAME_CLIENT_KEY,
    LAST_UPLOADED_BILL_NAME_KEY,
    WHATSAPP_SENDER_NUMBER,
    WHATSAPP_WEB_HOME_URL,
    admin_tabs,
    build_whatsapp_invite,
    comparison_form,
    create_archived_comparison_report,
    customer_optional_section_open,
    customer_registration_url,
    customer_status_snapshot,
    display_upload_name,
    force_customer_mode_data,
    force_operator_mode_data,
    home_page,
    is_archive_admin,
    mode_config,
    normalize_admin_focus_panel,
    public_access_page,
    store_report_summary_session,
)


logger = logging.getLogger(__name__)


def accesso_clienti(request):
    if request.user.is_authenticated:
        return _confronto_page(request, customer_mode=True)
    return public_access_page(request)


def _confronto_page(request, customer_mode=False, operator_mode=False):
    mode = mode_config(customer_mode, operator_mode)
    prepared = None
    rows = None
    extraction_warnings = []
    extraction_count = 0
    customer_invite_form = CustomerInviteForm()
    customer_invite_result = None
    customer_status = None
    report_summary_form = ReportSummaryUploadForm()
    report_summary = None
    report_summary_warnings = []
    archived_report = None
    admin_focus_panel = "confronto" if not customer_mode else ""
    uploaded_bill_name = request.session.get(mode["upload_name_session_key"], "")
    upload_form = BillUploadForm()

    if request.method == "POST":
        action = request.POST.get("action")

        if action == "send_customer_invite" and not customer_mode and not operator_mode:
            admin_focus_panel = "genera-codici"
            customer_invite_form = CustomerInviteForm(request.POST)
            form = comparison_form(customer_mode=customer_mode, operator_mode=operator_mode)
            if customer_invite_form.is_valid():
                customer_name = customer_invite_form.cleaned_data.get("customer_name", "").strip()
                customer_phone = customer_invite_form.cleaned_data["customer_phone"]
                customer_phone_display = customer_invite_form.cleaned_data["customer_phone_display"]
                invite_label = customer_name or f"Cliente {customer_phone_display}"
                invite_note = f"WhatsApp cliente: {customer_phone_display} | Creato da: {request.user.get_username() or 'utente'}"
                invite = InviteCode.objects.create(label=invite_label, note=invite_note)
                whatsapp_payload = build_whatsapp_invite(customer_name, customer_phone, invite.code)
                customer_invite_result = {
                    "code": invite.code,
                    "label": invite.label,
                    "customer_phone_display": customer_phone_display,
                    "registration_url": whatsapp_payload["registration_url"],
                    "message": whatsapp_payload["message"],
                    "whatsapp_url": whatsapp_payload["whatsapp_url"],
                }
                customer_invite_form = CustomerInviteForm(
                    initial={"customer_name": customer_name, "customer_phone": customer_phone_display, "whatsapp_ready": True}
                )

        elif action == "show_customer_status" and not customer_mode and not operator_mode:
            admin_focus_panel = "status-clienti"
            customer_status = customer_status_snapshot()
            form = comparison_form(customer_mode=customer_mode, operator_mode=operator_mode)

        elif action == "delete_customer_user" and not customer_mode and not operator_mode:
            admin_focus_panel = "status-clienti"
            form = comparison_form(customer_mode=customer_mode, operator_mode=operator_mode)
            customer_user = get_object_or_404(
                get_user_model(), pk=request.POST.get("customer_user_id"), is_staff=False, is_superuser=False
            )
            deleted_label = customer_user.email or customer_user.get_username() or f"cliente #{customer_user.pk}"
            customer_user.delete()
            messages.success(request, f"Cliente registrato eliminato: {deleted_label}.")
            customer_status = customer_status_snapshot()

        elif action == "build_report_summary" and not customer_mode:
            admin_focus_panel = "sunto-report"
            form = comparison_form(customer_mode=customer_mode, operator_mode=operator_mode)
            report_summary_form = ReportSummaryUploadForm(request.POST, request.FILES)
            if report_summary_form.is_valid():
                report_summary = build_reports_summary(report_summary_form.cleaned_data["report_files"])
                report_summary_warnings = report_summary.get("warnings", [])
                store_report_summary_session(request, report_summary)
                report_summary_form = ReportSummaryUploadForm()

        elif action == "extract_bill":
            if not customer_mode:
                admin_focus_panel = "confronto"
            request.session.pop(mode["comparison_session_key"], None)
            request.session.pop(mode["archive_report_session_key"], None)
            upload_form = BillUploadForm(request.POST, request.FILES)
            uploaded_file = request.FILES.get("bill_pdf")
            if uploaded_file:
                uploaded_bill_name = display_upload_name(uploaded_file.name)
                request.session[mode["upload_name_session_key"]] = uploaded_bill_name
            else:
                uploaded_bill_name = ""
                request.session.pop(mode["upload_name_session_key"], None)
            if upload_form.is_valid():
                try:
                    parsed = parse_uploaded_bill(upload_form.cleaned_data["bill_pdf"])
                except Exception:
                    logger.exception("Impossibile leggere il PDF della bolletta.")
                    form = comparison_form(customer_mode=customer_mode, operator_mode=operator_mode)
                    extraction_warnings = ["Non sono riuscito a leggere questo PDF. Inserisci manualmente i valori della bolletta."]
                else:
                    initial_data = parsed.values
                    if customer_mode:
                        initial_data = force_customer_mode_data(initial_data)
                    elif operator_mode:
                        initial_data = force_operator_mode_data(initial_data)
                    form = comparison_form(initial=initial_data, customer_mode=customer_mode, operator_mode=operator_mode)
                    extraction_warnings = parsed.warnings
                    extraction_count = len(parsed.values)
                    upload_form = BillUploadForm()
            else:
                form = comparison_form(customer_mode=customer_mode, operator_mode=operator_mode)

        elif action == "reset_bill":
            if not customer_mode:
                admin_focus_panel = "confronto"
            request.session.pop(mode["comparison_session_key"], None)
            request.session.pop(mode["archive_report_session_key"], None)
            request.session.pop(mode["upload_name_session_key"], None)
            uploaded_bill_name = ""
            providers = ["ILLUMIA"] if customer_mode else (request.POST.getlist("providers") or [request.POST.get("provider") or ""])
            initial_data = {
                "nome_cliente": request.POST.get("nome_cliente") or "",
                "indirizzo_fornitura": request.POST.get("indirizzo_fornitura") or "",
                "email_cliente": request.POST.get("email_cliente") or "",
                "telefono_cliente": request.POST.get("telefono_cliente") or "",
                "pod_pdr": request.POST.get("pod_pdr") or "",
                "segmento": request.POST.get("segmento") or "",
                "commodity": request.POST.get("commodity") or "",
                "bill_tariff_type": request.POST.get("bill_tariff_type") or "",
                "tariff_selection_mode": request.POST.get("tariff_selection_mode") or "",
                "providers": [p for p in providers if p],
                "tax_primary_home": request.POST.get("tax_primary_home") or "SI",
                "tax_power_kw": request.POST.get("tax_power_kw") or "0",
                "tax_annual_consumption": "",
                "tax_region": request.POST.get("tax_region") or "Veneto",
                "servizi_accessori_iva": "22%",
                "offer_var_choice_illumia": request.POST.get("offer_var_choice_illumia") or "",
                "offer_fix_choice_illumia": request.POST.get("offer_fix_choice_illumia") or "",
                "offer_var_choice_eon": request.POST.get("offer_var_choice_eon") or "",
                "offer_fix_choice_eon": request.POST.get("offer_fix_choice_eon") or "",
                "offer_var_choice_cve": request.POST.get("offer_var_choice_cve") or "",
                "offer_fix_choice_cve": request.POST.get("offer_fix_choice_cve") or "",
                "cve_over70": bool(request.POST.get("cve_over70")),
            }
            if customer_mode:
                initial_data = force_customer_mode_data(initial_data)
            elif operator_mode:
                initial_data = force_operator_mode_data(initial_data)
            form = comparison_form(initial=initial_data, customer_mode=customer_mode, operator_mode=operator_mode)

        else:
            if not customer_mode:
                admin_focus_panel = "confronto"
            form = comparison_form(request.POST, customer_mode=customer_mode, operator_mode=operator_mode)
            if form.is_valid():
                data = form.service_data()
                if customer_mode:
                    data = force_customer_mode_data(data)
                elif operator_mode:
                    data = force_operator_mode_data(data)
                data["comparison_datetime"] = timezone.localtime().isoformat(timespec="seconds")
                prepared = prepare_comparison(data)
                rows = prepared["rows"]
                session_data = form.session_data()
                if customer_mode:
                    session_data = force_customer_mode_data(session_data)
                elif operator_mode:
                    session_data = force_operator_mode_data(session_data)
                session_data["comparison_datetime"] = data["comparison_datetime"]
                request.session[mode["comparison_session_key"]] = session_data
                archived_report = create_archived_comparison_report(data, prepared, request.user)
                request.session[mode["archive_report_session_key"]] = archived_report.pk
                form = comparison_form(initial=data, customer_mode=customer_mode, operator_mode=operator_mode)

    else:
        if not customer_mode:
            requested_panel = normalize_admin_focus_panel(request.GET.get("panel"))
            admin_focus_panel = requested_panel if not operator_mode or requested_panel == "sunto-report" else "confronto"
            if admin_focus_panel == "status-clienti" and not operator_mode:
                customer_status = customer_status_snapshot()
            if admin_focus_panel == "sunto-report":
                report_summary = request.session.get(LAST_REPORT_SUMMARY_KEY)
                if report_summary:
                    report_summary_warnings = report_summary.get("warnings", [])
        initial = None
        if customer_mode:
            initial = {"providers": ["ILLUMIA"], "tariff_selection_mode": "LATEST", "pod_pdr": "", "email_cliente": "", "telefono_cliente": ""}
        elif operator_mode:
            initial = {"tariff_selection_mode": "LATEST"}
        form = comparison_form(initial=initial, customer_mode=customer_mode, operator_mode=operator_mode)

    return render(
        request,
        "confronti/confronto.html",
        {
            "form": form,
            "prepared": prepared,
            "rows": rows,
            "offer_options": offer_options_payload(),
            "upload_form": upload_form,
            "uploaded_bill_name": uploaded_bill_name,
            "extraction_warnings": extraction_warnings,
            "extraction_count": extraction_count,
            "customer_mode": customer_mode,
            "operator_mode": operator_mode,
            "illumia_only_mode": customer_mode,
            "customer_optional_open": customer_optional_section_open(form),
            "page_title": mode["page_title"],
            "page_intro": mode["page_intro"],
            "download_url_name": mode["download_url_name"],
            "brand_home_url": reverse("accesso_clienti") if customer_mode else reverse("confronto"),
            "customer_dashboard_local_url": reverse("accesso_clienti"),
            "customer_register_url": customer_registration_url(),
            "customer_invite_form": customer_invite_form,
            "customer_invite_result": customer_invite_result,
            "customer_status": customer_status,
            "archived_report": archived_report,
            "can_view_archive": is_archive_admin(request.user),
            "report_summary_form": report_summary_form,
            "report_summary": report_summary,
            "report_summary_warnings": report_summary_warnings,
            "admin_focus_panel": admin_focus_panel,
            "admin_tabs": admin_tabs(admin_focus_panel, operator_mode) if not customer_mode else [],
            "whatsapp_sender_number": WHATSAPP_SENDER_NUMBER,
            "whatsapp_web_home_url": WHATSAPP_WEB_HOME_URL,
        },
    )


def confronto(request):
    if not request.user.is_authenticated:
        return home_page(request)
    if not is_internal_user(request.user):
        return redirect("accesso_clienti")
    return _confronto_page(request, operator_mode=is_illumia_operator(request.user))


@login_required
def confronto_cliente_illumia(request):
    return redirect("accesso_clienti")
