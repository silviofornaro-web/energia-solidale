import io
import logging
import os
import shutil
import subprocess
from datetime import datetime
from urllib.parse import quote, urlencode, urljoin

from django.conf import settings
from django.contrib import messages
from django.contrib.auth import get_user_model
from django.contrib.auth.decorators import login_required
from django.core.files.base import ContentFile
from django.db.models import Count, Max, Q
from django.http import FileResponse, HttpResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse
from django.utils import timezone

from .bill_parser import parse_uploaded_bill
from .forms import (
    ArchiveFolderForm,
    ArchiveReportForm,
    ArchiveReportReplaceFileForm,
    ArchiveReportUploadForm,
    BillUploadForm,
    ConfrontoForm,
    CustomerInviteForm,
    ReportSummaryUploadForm,
    session_to_service_data,
)
from .models import ComparisonReport, CustomerArchiveFolder, InviteCode
from .roles import is_illumia_operator, is_internal_user
from .services import (
    build_excel_bytes,
    build_reports_summary,
    build_reports_summary_excel,
    normalize_providers,
    offer_options_payload,
    prepare_comparison,
    provider_label,
    safe_download_filename,
)


logger = logging.getLogger(__name__)
LAST_UPLOADED_BILL_NAME_KEY = "last_uploaded_bill_name"
LAST_UPLOADED_BILL_NAME_CLIENT_KEY = "last_uploaded_bill_name_cliente_illumia"
LAST_COMPARISON_KEY = "last_confronto"
LAST_COMPARISON_CLIENT_KEY = "last_confronto_cliente_illumia"
LAST_REPORT_SUMMARY_KEY = "last_report_summary"
WHATSAPP_SENDER_NUMBER = os.environ.get("WHATSAPP_SENDER_NUMBER", "3271044102")
WHATSAPP_WEB_HOME_URL = "https://web.whatsapp.com/"
CUSTOMER_PORTAL_BASE_URL = (
    os.environ.get("CUSTOMER_PORTAL_BASE_URL")
    or (f"https://{os.environ['RENDER_EXTERNAL_HOSTNAME']}" if os.environ.get("RENDER_EXTERNAL_HOSTNAME") else "")
    or "https://energia-solidale.onrender.com"
).rstrip("/")


def _display_upload_name(raw_name):
    return (raw_name or "").replace("\\", "/").split("/")[-1].strip()


def _public_customer_url(path):
    return urljoin(f"{CUSTOMER_PORTAL_BASE_URL}/", path.lstrip("/"))


def _customer_registration_url(invite_code=""):
    base_url = _public_customer_url(reverse("register"))
    normalized_code = InviteCode.normalize_code(invite_code)
    if not normalized_code:
        return base_url
    return f"{base_url}?invite_code={normalized_code}"


def _build_whatsapp_invite(customer_name, customer_phone, invite_code):
    registration_url = _customer_registration_url(invite_code)
    saluto = f"Ciao {customer_name}," if customer_name else "Ciao,"
    message = (
        f"{saluto} per creare il tuo accesso all'area clienti Energia Solidale apri questo link: "
        f"{registration_url} "
        f"e completa la registrazione. "
        f"Il tuo codice invito e {invite_code}. "
        "Dopo l'accesso potrai usare la dashboard cliente per confrontare la bolletta con Illumia."
    )
    whatsapp_url = f"https://web.whatsapp.com/send?phone={customer_phone}&text={quote(message)}"
    return {
        "message": message,
        "registration_url": registration_url,
        "whatsapp_url": whatsapp_url,
    }


def _mode_config(customer_mode=False, operator_mode=False):
    if customer_mode:
        return {
            "comparison_session_key": LAST_COMPARISON_CLIENT_KEY,
            "upload_name_session_key": LAST_UPLOADED_BILL_NAME_CLIENT_KEY,
            "page_title": "Confronto bolletta con Illumia",
            "page_intro": (
                "Area clienti riservata. Dopo l'accesso puoi confrontare la tua bolletta "
                "autonomamente solo con le tariffe Illumia."
            ),
            "download_url_name": "scarica_excel_cliente_illumia",
        }
    if operator_mode:
        return {
            "comparison_session_key": LAST_COMPARISON_KEY,
            "upload_name_session_key": LAST_UPLOADED_BILL_NAME_KEY,
            "page_title": "Confronto bollette",
            "page_intro": "",
            "download_url_name": "scarica_excel",
        }
    return {
        "comparison_session_key": LAST_COMPARISON_KEY,
        "upload_name_session_key": LAST_UPLOADED_BILL_NAME_KEY,
        "page_title": "Confronto bollette",
        "page_intro": "",
        "download_url_name": "scarica_excel",
    }


def _force_illumia_only_data(data, *, clear_pod_pdr=False):
    payload = dict(data)
    if clear_pod_pdr:
        payload["pod_pdr"] = ""
    payload["providers"] = ["ILLUMIA"]
    payload["provider"] = "ILLUMIA"
    payload["tariff_selection_mode"] = "LATEST"
    payload["offer_var_choice_eon"] = ""
    payload["offer_fix_choice_eon"] = ""
    payload["offer_var_choice_cve"] = ""
    payload["offer_fix_choice_cve"] = ""
    payload["cve_over70"] = False
    return payload


def _force_customer_mode_data(data):
    return _force_illumia_only_data(data, clear_pod_pdr=True)


def _force_operator_mode_data(data):
    payload = dict(data)
    providers = [provider for provider in normalize_providers(payload.get("providers") or payload.get("provider")) if provider in {"ILLUMIA", "EON"}]
    if not providers:
        providers = ["ILLUMIA"]
    payload["providers"] = providers
    payload["provider"] = providers[0]
    payload["tariff_selection_mode"] = "LATEST"
    payload["offer_var_choice_cve"] = ""
    payload["offer_fix_choice_cve"] = ""
    payload["cve_over70"] = False
    return payload


def _comparison_form(*args, customer_mode=False, operator_mode=False, **kwargs):
    return ConfrontoForm(*args, customer_mode=customer_mode, operator_mode=operator_mode, **kwargs)


def _customer_optional_section_open(form):
    if not getattr(form, "customer_mode", False):
        return False
    return any(field_name in form.errors for field_name in ConfrontoForm.CUSTOMER_OPTIONAL_FIELDS)


def _customer_status_snapshot():
    User = get_user_model()
    users = list(User.objects.filter(is_staff=False, is_superuser=False).order_by("-id"))
    available_invites = list(
        InviteCode.objects.filter(is_active=True, used_at__isnull=True, used_by__isnull=True).order_by("-created_at")
    )
    used_invites = list(
        InviteCode.objects.exclude(used_at__isnull=True, used_by__isnull=True)
        .select_related("used_by")
        .order_by("-used_at", "-created_at")
    )
    return {
        "users": users,
        "available_invites": available_invites,
        "used_invites": used_invites,
        "user_count": len(users),
        "available_count": len(available_invites),
        "used_count": len(used_invites),
    }


def _is_archive_admin(user):
    return bool(
        getattr(user, "is_authenticated", False)
        and (getattr(user, "is_superuser", False) or (getattr(user, "is_staff", False) and not is_illumia_operator(user)))
    )


def _redirect_for_non_archive_admin(request):
    if request.user.is_authenticated and is_internal_user(request.user):
        return redirect("confronto")
    return redirect("accesso_clienti")


def _parse_comparison_datetime(value):
    if not value:
        return timezone.localtime()
    try:
        parsed = datetime.fromisoformat(str(value))
    except ValueError:
        return timezone.localtime()
    if timezone.is_naive(parsed):
        parsed = timezone.make_aware(parsed, timezone.get_current_timezone())
    return parsed


def _archive_match_folder(customer_name, customer_email="", customer_phone=""):
    normalized_name = CustomerArchiveFolder.normalize_customer_name(customer_name)
    normalized_email = CustomerArchiveFolder.normalize_customer_email(customer_email)
    normalized_phone = CustomerArchiveFolder.normalize_customer_phone(customer_phone)
    folders = CustomerArchiveFolder.objects.all()
    if normalized_email:
        folder = folders.filter(customer_email=normalized_email).order_by("-updated_at", "-created_at").first()
        if folder:
            return folder
    if normalized_phone:
        folder = folders.filter(customer_phone=normalized_phone).order_by("-updated_at", "-created_at").first()
        if folder:
            return folder
    if normalized_name:
        return folders.filter(customer_name__iexact=normalized_name).order_by("-updated_at", "-created_at").first()
    return None


def _get_or_create_archive_folder(data, user):
    customer_name = CustomerArchiveFolder.normalize_customer_name(data.get("nome_cliente")) or "Cliente senza nome"
    customer_email = CustomerArchiveFolder.normalize_customer_email(data.get("email_cliente"))
    customer_phone = CustomerArchiveFolder.normalize_customer_phone(data.get("telefono_cliente"))
    folder = _archive_match_folder(customer_name, customer_email, customer_phone)
    if folder is None:
        return CustomerArchiveFolder.objects.create(
            customer_name=customer_name,
            customer_email=customer_email,
            customer_phone=customer_phone,
            created_by=user,
        )
    changed_fields = []
    if customer_name and folder.customer_name != customer_name:
        folder.customer_name = customer_name
        changed_fields.append("customer_name")
    if customer_email and folder.customer_email != customer_email:
        folder.customer_email = customer_email
        changed_fields.append("customer_email")
    if customer_phone and folder.customer_phone != customer_phone:
        folder.customer_phone = customer_phone
        changed_fields.append("customer_phone")
    if changed_fields:
        folder.save(update_fields=changed_fields + ["updated_at"])
    return folder


def _archive_report_payload(data, prepared):
    return {
        "customer_name": data.get("nome_cliente", ""),
        "supply_address": data.get("indirizzo_fornitura", ""),
        "customer_email": data.get("email_cliente", ""),
        "customer_phone": data.get("telefono_cliente", ""),
        "commodity": data.get("commodity", ""),
        "providers": list(data.get("providers", [])),
        "providers_label": prepared["calc"].get("providers_label", ""),
        "bill_start": data.get("bill_start").isoformat() if data.get("bill_start") else "",
        "bill_end": data.get("bill_end").isoformat() if data.get("bill_end") else "",
        "bill_period_label": prepared["calc"].get("period_label", ""),
        "comparison_datetime": data.get("comparison_datetime", ""),
    }


def _archive_report_filename(data):
    provider_name = "_".join(provider_label(provider).lower().replace(".", "") for provider in data.get("providers", []))
    provider_name = provider_name or provider_label(data.get("provider", "ILLUMIA")).lower().replace(".", "")
    return safe_download_filename(
        f"confronto_{provider_name}_{data.get('nome_cliente', 'Cliente')}_{data.get('commodity', 'ENERGIA')}.xlsx"
    )


def _archive_report_title(data, prepared):
    provider_name = prepared["calc"].get("providers_label", "Illumia")
    commodity = prepared["calc"].get("commodity_label", data.get("commodity", ""))
    period = prepared["calc"].get("period_label", "")
    title_parts = [provider_name, commodity]
    if period:
        title_parts.append(period)
    return " - ".join(part for part in title_parts if part)


def _archive_report_queryset(search_query=""):
    folders = CustomerArchiveFolder.objects.annotate(report_count=Count("reports"))
    query = str(search_query or "").strip()
    if query:
        folders = folders.filter(
            Q(customer_name__icontains=query)
            | Q(customer_email__icontains=query)
            | Q(customer_phone__icontains=query)
            | Q(folder_name__icontains=query)
            | Q(reports__title__icontains=query)
            | Q(reports__providers_label__icontains=query)
        ).distinct()
    return folders.order_by("-updated_at", "-created_at")


def _archive_reports_queryset(search_query=""):
    reports = ComparisonReport.objects.select_related("folder")
    query = str(search_query or "").strip()
    if query:
        reports = reports.filter(
            Q(folder__customer_name__icontains=query)
            | Q(folder__customer_email__icontains=query)
            | Q(folder__customer_phone__icontains=query)
            | Q(folder__folder_name__icontains=query)
            | Q(title__icontains=query)
            | Q(providers_label__icontains=query)
            | Q(original_filename__icontains=query)
            | Q(bill_period_label__icontains=query)
        )
    return reports.order_by("-comparison_datetime", "-created_at")


def _archive_summary(search_query=""):
    report_qs = _archive_reports_queryset(search_query)
    query = str(search_query or "").strip()
    return {
        "folder_count": _archive_report_queryset(query).count(),
        "report_count": report_qs.count(),
        "latest_report_at": report_qs.aggregate(last_saved=Max("created_at"))["last_saved"],
    }


def _store_report_summary_session(request, report_summary):
    request.session[LAST_REPORT_SUMMARY_KEY] = {
        "columns": report_summary.get("columns", []),
        "rows": report_summary.get("rows", []),
        "warnings": report_summary.get("warnings", []),
        "count": report_summary.get("count", 0),
    }


def _archive_folder_absolute_path(folder):
    return os.path.join(str(settings.MEDIA_ROOT), "report_archive", folder.folder_name)


def _archive_report_absolute_path(report):
    try:
        return report.report_file.path
    except (NotImplementedError, ValueError):
        return os.path.join(_archive_folder_absolute_path(report.folder), os.path.basename(report.report_file.name))


def _can_open_local_archive_path():
    if os.environ.get("RENDER_EXTERNAL_HOSTNAME"):
        return False
    if os.name == "nt":
        return True
    return bool(shutil.which("open") or shutil.which("xdg-open"))


def _archive_storage_label():
    if os.environ.get("RENDER_EXTERNAL_HOSTNAME"):
        return "Cloud app (Render + disco persistente)"
    return "Computer locale"


def _open_local_archive_path(path):
    if not _can_open_local_archive_path():
        raise RuntimeError("Apertura cartella disponibile solo in locale.")
    if not os.path.isdir(path):
        raise FileNotFoundError(path)
    if os.name == "nt":
        command = ["explorer", path]
    elif shutil.which("open"):
        command = ["open", path]
    else:
        command = ["xdg-open", path]
    subprocess.Popen(command)


def _touch_archive_folder(folder):
    CustomerArchiveFolder.objects.filter(pk=folder.pk).update(updated_at=timezone.now())


def _archive_uploaded_report_title(filename):
    stem = os.path.splitext(_display_upload_name(filename))[0]
    normalized = " ".join(stem.replace("_", " ").replace("-", " ").split())
    return normalized or "Report confronto"


def _create_uploaded_archive_report(folder, form, user):
    uploaded_file = form.cleaned_data["report_file"]
    report = ComparisonReport(
        folder=folder,
        title=str(form.cleaned_data.get("title") or "").strip() or _archive_uploaded_report_title(uploaded_file.name),
        notes=str(form.cleaned_data.get("notes") or "").strip(),
        original_filename=_display_upload_name(uploaded_file.name),
        comparison_datetime=timezone.now(),
        created_by=user,
    )
    report.report_file.save(report.original_filename, uploaded_file, save=False)
    report.save()
    _touch_archive_folder(folder)
    return report


def _replace_archive_report_file(report, uploaded_file):
    old_name = report.report_file.name
    report.report_file.save(_display_upload_name(uploaded_file.name), uploaded_file, save=False)
    report.original_filename = _display_upload_name(uploaded_file.name)
    if not report.title:
        report.title = _archive_uploaded_report_title(uploaded_file.name)
    if report.comparison_datetime is None:
        report.comparison_datetime = timezone.now()
    report.save()
    if old_name and old_name != report.report_file.name:
        try:
            report.report_file.storage.delete(old_name)
        except OSError:
            logger.warning("Impossibile eliminare il vecchio file archivio %s", old_name, exc_info=True)
    _touch_archive_folder(report.folder)


def _delete_archive_report(report):
    stored_name = report.report_file.name
    folder = report.folder
    report.delete()
    if stored_name:
        try:
            report.report_file.storage.delete(stored_name)
        except OSError:
            logger.warning("Impossibile eliminare il file archivio %s", stored_name, exc_info=True)
    _touch_archive_folder(folder)


def _delete_archive_folder(folder):
    reports = list(folder.reports.all())
    folder_name = folder.folder_name
    folder_path = _archive_folder_absolute_path(folder)
    report_count = len(reports)
    for report in reports:
        _delete_archive_report(report)
    folder.delete()
    if os.path.isdir(folder_path):
        shutil.rmtree(folder_path, ignore_errors=True)
    return {"folder_name": folder_name, "report_count": report_count}


def _build_reports_summary_from_archived_reports(reports):
    summary_files = []
    try:
        for report in reports:
            report.report_file.open("rb")
            try:
                content = report.report_file.read()
            finally:
                report.report_file.close()
            stream = io.BytesIO(content)
            stream.name = report.original_filename or os.path.basename(report.report_file.name) or f"report-{report.pk}.xlsx"
            summary_files.append(stream)
        return build_reports_summary(summary_files)
    finally:
        for stream in summary_files:
            stream.close()


def _selected_archive_folder_ids(values):
    selected_ids = []
    seen_ids = set()
    for value in values:
        if not str(value).isdigit():
            continue
        folder_id = int(value)
        if folder_id in seen_ids:
            continue
        seen_ids.add(folder_id)
        selected_ids.append(folder_id)
    return selected_ids


def _selected_archive_folders_and_reports(folder_ids):
    folders = list(
        CustomerArchiveFolder.objects.filter(pk__in=folder_ids)
        .annotate(report_count=Count("reports"))
        .order_by("customer_name", "-updated_at", "-created_at")
    )
    reports = list(
        ComparisonReport.objects.filter(folder_id__in=[folder.pk for folder in folders])
        .select_related("folder")
        .order_by("folder__customer_name", "-comparison_datetime", "-created_at")
    )
    empty_folders = [folder for folder in folders if not folder.report_count]
    return folders, reports, empty_folders


def _archive_context(
    request,
    *,
    selected_folder=None,
    folder_form=None,
    add_report_form=None,
    selected_report_ids=None,
    selected_archive_folder_ids=None,
    selected_archive_folders=None,
    report_summary=None,
    report_summary_warnings=None,
    report_summary_scope="reports",
    active_report_id=None,
    active_report_form=None,
    active_replace_report_id=None,
    active_replace_report_form=None,
):
    search_query = str(request.GET.get("q") or request.POST.get("q") or "").strip()
    folders = list(_archive_report_queryset(search_query))
    report_entries = []
    all_report_entries = []
    selected_report_ids = [int(report_id) for report_id in (selected_report_ids or [])]
    selected_archive_folder_ids = [int(folder_id) for folder_id in (selected_archive_folder_ids or [])]
    selected_archive_folders = list(selected_archive_folders or [])
    report_summary_warnings = list(report_summary_warnings or [])
    archive_folder_absolute_path = ""
    archive_local_open_available = _can_open_local_archive_path()
    if selected_folder is not None:
        selected_folder = get_object_or_404(
            CustomerArchiveFolder.objects.prefetch_related("reports").annotate(report_count=Count("reports")),
            pk=selected_folder.pk,
        )
        archive_folder_absolute_path = _archive_folder_absolute_path(selected_folder)
        reports = list(selected_folder.reports.order_by("-comparison_datetime", "-created_at"))
        if folder_form is None:
            folder_form = ArchiveFolderForm(instance=selected_folder)
        if add_report_form is None:
            add_report_form = ArchiveReportUploadForm(prefix="add-report")
        for report in reports:
            form = (
                active_report_form
                if active_report_id == report.pk and active_report_form is not None
                else ArchiveReportForm(instance=report, prefix=f"report-{report.pk}")
            )
            replace_form = (
                active_replace_report_form
                if active_replace_report_id == report.pk and active_replace_report_form is not None
                else ArchiveReportReplaceFileForm(prefix=f"replace-report-{report.pk}")
            )
            report_entries.append(
                {
                    "report": report,
                    "form": form,
                    "replace_form": replace_form,
                    "absolute_path": _archive_report_absolute_path(report),
                }
            )
    else:
        for report in _archive_reports_queryset(search_query):
            all_report_entries.append(
                {
                    "report": report,
                    "absolute_path": _archive_report_absolute_path(report),
                }
            )
    return {
        "brand_home_url": reverse("confronto"),
        "page_title": "Archivio report",
        "admin_tabs": _admin_tabs("archivio-report"),
        "archive_query": search_query,
        "archive_summary": _archive_summary(search_query),
        "archive_folders": folders,
        "selected_folder": selected_folder,
        "folder_form": folder_form,
        "add_report_form": add_report_form,
        "archive_storage_label": _archive_storage_label(),
        "archive_folder_absolute_path": archive_folder_absolute_path,
        "archive_local_open_available": archive_local_open_available,
        "selected_report_ids": selected_report_ids,
        "selected_archive_folder_ids": selected_archive_folder_ids,
        "selected_archive_folders": selected_archive_folders,
        "archive_report_summary": report_summary,
        "archive_report_summary_warnings": report_summary_warnings,
        "archive_report_summary_scope": report_summary_scope,
        "report_entries": report_entries,
        "all_report_entries": all_report_entries,
    }


def _archive_redirect_url(search_query="", *, folder_id=None):
    if folder_id is not None:
        base_url = reverse("archivio_report_cartella", kwargs={"folder_id": folder_id})
    else:
        base_url = reverse("archivio_report")
    query = str(search_query or "").strip()
    if not query:
        return base_url
    return f"{base_url}?{urlencode({'q': query})}"


def _public_access_page(request):
    return render(
        request,
        "registration/access_choice.html",
        {
            "brand_home_url": reverse("accesso_clienti"),
            "login_url": reverse("login"),
            "register_url": reverse("register"),
        },
    )


def _home_tabs():
    root_url = reverse("confronto")
    login_url = reverse("login")
    return [
        {
            "key": "mostra-dashboard-clienti",
            "label": "Mostra Dashboard Clienti",
            "href": f"{login_url}?{urlencode({'next': reverse('accesso_clienti')})}",
        },
        {
            "key": "mostra-registrazione-clienti",
            "label": "Mostra Registrazione Clienti",
            "href": reverse("register"),
        },
        {
            "key": "stato-clienti",
            "label": "Stato Clienti",
            "href": f"{login_url}?{urlencode({'next': f'{root_url}?panel=status-clienti'})}",
        },
        {
            "key": "genera-codici",
            "label": "Genera Codici",
            "href": f"{login_url}?{urlencode({'next': f'{root_url}?panel=genera-codici'})}",
        },
        {
            "key": "archivio-report",
            "label": "Apri archivio file",
            "href": f"{login_url}?{urlencode({'next': reverse('archivio_report')})}",
        },
        {
            "key": "confronto-bollette-offerte",
            "label": "Confronto Bollette/Offerte",
            "href": f"{login_url}?{urlencode({'next': f'{root_url}?panel=confronto'})}",
        },
        {
            "key": "sunto-report",
            "label": "Sunto Report",
            "href": f"{login_url}?{urlencode({'next': f'{root_url}?panel=sunto-report'})}",
        },
    ]


def _normalize_admin_focus_panel(panel):
    return panel if panel in {"confronto", "genera-codici", "status-clienti", "sunto-report", "archivio-report"} else "confronto"


def _admin_tabs(active_panel, operator_mode=False):
    root_url = reverse("confronto")
    comparison_tab = {
        "key": "confronto",
        "label": "Esegui confronto",
        "href": f"{root_url}?panel=confronto#confronto-bollette-offerte",
        "active": active_panel == "confronto",
    }
    if operator_mode:
        return [
            {
                "key": "sunto-report",
                "label": "Sunto report",
                "href": f"{root_url}?panel=sunto-report#sunto-report",
                "active": active_panel == "sunto-report",
            }
        ]
    return [
        comparison_tab,
        {
            "key": "genera-codici",
            "label": "Genera Codice invito",
            "href": f"{root_url}?panel=genera-codici#genera-codici",
            "active": active_panel == "genera-codici",
        },
        {
            "key": "status-clienti",
            "label": "Stato clienti",
            "href": f"{root_url}?panel=status-clienti#stato-clienti",
            "active": active_panel == "status-clienti",
        },
        {
            "key": "archivio-report",
            "label": "Apri archivio file",
            "href": reverse("archivio_report"),
            "active": active_panel == "archivio-report",
        },
        {
            "key": "sunto-report",
            "label": "Sunto report",
            "href": f"{root_url}?panel=sunto-report#sunto-report",
            "active": active_panel == "sunto-report",
        },
        {
            "key": "apri-dashboard-cliente",
            "label": "Apri dashboard cliente",
            "href": reverse("accesso_clienti"),
            "active": False,
            "new_tab": True,
        },
        {
            "key": "apri-registrazione-cliente",
            "label": "Apri registrazione cliente",
            "href": reverse("register"),
            "active": False,
            "new_tab": True,
        },
    ]


def _home_page(request):
    return render(
        request,
        "confronti/home.html",
        {
            "brand_home_url": reverse("confronto"),
            "page_title": "Energia Solidale",
            "hide_header_brand": True,
            "admin_tabs": _home_tabs(),
            "customer_dashboard_url": reverse("accesso_clienti"),
        },
    )


def accesso_clienti(request):
    if request.user.is_authenticated:
        return _confronto_page(request, customer_mode=True)
    return _public_access_page(request)


def _confronto_page(request, *, customer_mode=False, operator_mode=False):
    illumia_only_mode = customer_mode
    mode = _mode_config(customer_mode, operator_mode)
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
    admin_focus_panel = "confronto" if not customer_mode else ""
    uploaded_bill_name = request.session.get(mode["upload_name_session_key"], "")
    upload_form = BillUploadForm()
    if request.method == "POST":
        action = request.POST.get("action")
        if action == "send_customer_invite" and not customer_mode and not operator_mode:
            admin_focus_panel = "genera-codici"
            customer_invite_form = CustomerInviteForm(request.POST)
            form = _comparison_form(customer_mode=customer_mode, operator_mode=operator_mode)
            if customer_invite_form.is_valid():
                customer_name = customer_invite_form.cleaned_data.get("customer_name", "").strip()
                customer_phone = customer_invite_form.cleaned_data["customer_phone"]
                customer_phone_display = customer_invite_form.cleaned_data["customer_phone_display"]
                invite_label = customer_name or f"Cliente {customer_phone_display}"
                invite_note = (
                    f"WhatsApp cliente: {customer_phone_display} | Creato da: {request.user.get_username() or 'utente'}"
                )
                invite = InviteCode.objects.create(label=invite_label, note=invite_note)
                whatsapp_payload = _build_whatsapp_invite(customer_name, customer_phone, invite.code)
                customer_invite_result = {
                    "code": invite.code,
                    "label": invite.label,
                    "customer_phone_display": customer_phone_display,
                    "registration_url": whatsapp_payload["registration_url"],
                    "message": whatsapp_payload["message"],
                    "whatsapp_url": whatsapp_payload["whatsapp_url"],
                }
                customer_invite_form = CustomerInviteForm(
                    initial={
                        "customer_name": customer_name,
                        "customer_phone": customer_phone_display,
                        "whatsapp_ready": True,
                    }
                )
        elif action == "show_customer_status" and not customer_mode and not operator_mode:
            admin_focus_panel = "status-clienti"
            customer_status = _customer_status_snapshot()
            form = _comparison_form(customer_mode=customer_mode, operator_mode=operator_mode)
        elif action == "delete_customer_user" and not customer_mode and not operator_mode:
            admin_focus_panel = "status-clienti"
            form = _comparison_form(customer_mode=customer_mode, operator_mode=operator_mode)
            customer_user = get_object_or_404(
                get_user_model(),
                pk=request.POST.get("customer_user_id"),
                is_staff=False,
                is_superuser=False,
            )
            deleted_label = customer_user.email or customer_user.get_username() or f"cliente #{customer_user.pk}"
            customer_user.delete()
            messages.success(request, f"Cliente registrato eliminato: {deleted_label}.")
            customer_status = _customer_status_snapshot()
        elif action == "build_report_summary" and not customer_mode:
            admin_focus_panel = "sunto-report"
            form = _comparison_form(customer_mode=customer_mode, operator_mode=operator_mode)
            report_summary_form = ReportSummaryUploadForm(request.POST, request.FILES)
            if report_summary_form.is_valid():
                report_summary = build_reports_summary(report_summary_form.cleaned_data["report_files"])
                report_summary_warnings = report_summary.get("warnings", [])
                _store_report_summary_session(request, report_summary)
                report_summary_form = ReportSummaryUploadForm()
        elif action == "extract_bill":
            if not customer_mode:
                admin_focus_panel = "confronto"
            request.session.pop(mode["comparison_session_key"], None)
            upload_form = BillUploadForm(request.POST, request.FILES)
            uploaded_file = request.FILES.get("bill_pdf")
            if uploaded_file:
                uploaded_bill_name = _display_upload_name(uploaded_file.name)
                request.session[mode["upload_name_session_key"]] = uploaded_bill_name
            else:
                uploaded_bill_name = ""
                request.session.pop(mode["upload_name_session_key"], None)
            if upload_form.is_valid():
                try:
                    parsed = parse_uploaded_bill(upload_form.cleaned_data["bill_pdf"])
                except Exception:
                    logger.exception("Impossibile leggere il PDF della bolletta.")
                    form = _comparison_form(customer_mode=customer_mode, operator_mode=operator_mode)
                    extraction_warnings = [
                        "Non sono riuscito a leggere questo PDF. Inserisci manualmente i valori della bolletta."
                    ]
                else:
                    initial_data = parsed.values
                    if customer_mode:
                        initial_data = _force_customer_mode_data(initial_data)
                    elif operator_mode:
                        initial_data = _force_operator_mode_data(initial_data)
                    form = _comparison_form(
                        initial=initial_data,
                        customer_mode=customer_mode,
                        operator_mode=operator_mode,
                    )
                    extraction_warnings = parsed.warnings
                    extraction_count = len(parsed.values)
                    upload_form = BillUploadForm()
            else:
                form = _comparison_form(customer_mode=customer_mode, operator_mode=operator_mode)
        elif action == "reset_bill":
            if not customer_mode:
                admin_focus_panel = "confronto"
            request.session.pop(mode["comparison_session_key"], None)
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
                "offer_var_choice_cve": request.POST.get("offer_var_choice_cve") or "",
                "offer_fix_choice_cve": request.POST.get("offer_fix_choice_cve") or "",
                "cve_over70": bool(request.POST.get("cve_over70")),
            }
            if customer_mode:
                initial_data = _force_customer_mode_data(initial_data)
            elif operator_mode:
                initial_data = _force_operator_mode_data(initial_data)
            form = _comparison_form(initial=initial_data, customer_mode=customer_mode, operator_mode=operator_mode)
        else:
            if not customer_mode:
                admin_focus_panel = "confronto"
            form = _comparison_form(request.POST, customer_mode=customer_mode, operator_mode=operator_mode)
            if form.is_valid():
                data = form.service_data()
                if customer_mode:
                    data = _force_customer_mode_data(data)
                elif operator_mode:
                    data = _force_operator_mode_data(data)
                data["comparison_datetime"] = timezone.localtime().isoformat(timespec="seconds")
                prepared = prepare_comparison(data)
                rows = prepared["rows"]
                session_data = form.session_data()
                if customer_mode:
                    session_data = _force_customer_mode_data(session_data)
                elif operator_mode:
                    session_data = _force_operator_mode_data(session_data)
                session_data["comparison_datetime"] = data["comparison_datetime"]
                request.session[mode["comparison_session_key"]] = session_data
                form = _comparison_form(initial=data, customer_mode=customer_mode, operator_mode=operator_mode)
    else:
        if not customer_mode:
            requested_panel = _normalize_admin_focus_panel(request.GET.get("panel"))
            admin_focus_panel = requested_panel if not operator_mode or requested_panel == "sunto-report" else "confronto"
            if admin_focus_panel == "status-clienti" and not operator_mode:
                customer_status = _customer_status_snapshot()
            if admin_focus_panel == "sunto-report":
                report_summary = request.session.get(LAST_REPORT_SUMMARY_KEY)
                if report_summary:
                    report_summary_warnings = report_summary.get("warnings", [])
        initial = None
        if customer_mode:
            initial = {
                "providers": ["ILLUMIA"],
                "tariff_selection_mode": "LATEST",
                "pod_pdr": "",
                "email_cliente": "",
                "telefono_cliente": "",
            }
        elif operator_mode:
            initial = {"tariff_selection_mode": "LATEST"}
        form = _comparison_form(initial=initial, customer_mode=customer_mode, operator_mode=operator_mode)

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
            "illumia_only_mode": illumia_only_mode,
            "customer_optional_open": _customer_optional_section_open(form),
            "page_title": mode["page_title"],
            "page_intro": mode["page_intro"],
            "download_url_name": mode["download_url_name"],
            "brand_home_url": reverse("accesso_clienti") if customer_mode else reverse("confronto"),
            "customer_dashboard_local_url": reverse("accesso_clienti"),
            "customer_register_url": _customer_registration_url(),
            "customer_invite_form": customer_invite_form,
            "customer_invite_result": customer_invite_result,
            "customer_status": customer_status,
            "report_summary_form": report_summary_form,
            "report_summary": report_summary,
            "report_summary_warnings": report_summary_warnings,
            "admin_focus_panel": admin_focus_panel,
            "admin_tabs": _admin_tabs(admin_focus_panel, operator_mode) if not customer_mode else [],
            "whatsapp_sender_number": WHATSAPP_SENDER_NUMBER,
            "whatsapp_web_home_url": WHATSAPP_WEB_HOME_URL,
        },
    )


def _scarica_excel(request, *, customer_mode=False, operator_mode=False):
    mode = _mode_config(customer_mode, operator_mode)
    raw = request.session.get(mode["comparison_session_key"])
    if not raw:
        return HttpResponse("Nessun confronto pronto. Torna alla pagina principale e calcola il confronto.", status=400)
    data = session_to_service_data(raw)
    if customer_mode:
        data = _force_customer_mode_data(data)
    elif operator_mode:
        data = _force_operator_mode_data(data)
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


@login_required
def archivia_report_corrente(request):
    if not is_internal_user(request.user):
        return redirect("accesso_clienti")
    if request.method != "POST":
        return redirect("confronto")
    raw = request.session.get(LAST_COMPARISON_KEY)
    if not raw:
        messages.error(request, "Non c'e nessun confronto interno pronto da archiviare.")
        return redirect(f"{reverse('confronto')}?panel=confronto")
    data = session_to_service_data(raw)
    prepared = prepare_comparison(data)
    content = build_excel_bytes(data, prepared)
    folder = _get_or_create_archive_folder(data, request.user)
    report = ComparisonReport(
        folder=folder,
        title=_archive_report_title(data, prepared),
        commodity=data.get("commodity", ""),
        providers_label=prepared["calc"].get("providers_label", ""),
        bill_period_label=prepared["calc"].get("period_label", ""),
        comparison_datetime=_parse_comparison_datetime(data.get("comparison_datetime")),
        original_filename=_archive_report_filename(data),
        comparison_data=_archive_report_payload(data, prepared),
        created_by=request.user,
    )
    report.report_file.save(report.original_filename, ContentFile(content), save=False)
    report.save()
    _touch_archive_folder(folder)
    messages.success(request, f"Report archiviato nella cartella {folder.folder_name}.")
    if _is_archive_admin(request.user):
        return redirect("archivio_report_cartella", folder_id=folder.pk)
    return redirect(f"{reverse('confronto')}?panel=confronto")


@login_required
def archivio_report(request):
    if not _is_archive_admin(request.user):
        return _redirect_for_non_archive_admin(request)
    if request.method == "POST":
        action = request.POST.get("action")
        search_query = request.POST.get("q")
        if action == "delete_archive_report_global":
            report = get_object_or_404(ComparisonReport.objects.select_related("folder"), pk=request.POST.get("report_id"))
            folder_name = report.folder.folder_name
            _delete_archive_report(report)
            messages.success(request, f"Report eliminato dall'archivio della cartella {folder_name}.")
            return redirect(_archive_redirect_url(search_query))
        if action == "delete_archive_folder":
            folder = get_object_or_404(CustomerArchiveFolder, pk=request.POST.get("folder_id"))
            deleted = _delete_archive_folder(folder)
            messages.success(
                request,
                f"Cartella {deleted['folder_name']} eliminata con {deleted['report_count']} report.",
            )
            return redirect(_archive_redirect_url(search_query))
        if action == "build_archive_folder_summary":
            selected_folder_ids = _selected_archive_folder_ids(request.POST.getlist("selected_folder_ids"))
            if not selected_folder_ids:
                messages.error(request, "Seleziona almeno una cartella cliente da includere nel sunto.")
                return render(
                    request,
                    "confronti/archive.html",
                    _archive_context(request, selected_archive_folder_ids=selected_folder_ids, report_summary_scope="folders"),
                )
            selected_folders, selected_reports, empty_folders = _selected_archive_folders_and_reports(selected_folder_ids)
            if not selected_folders:
                messages.error(request, "Le cartelle selezionate non sono disponibili nell'archivio.")
                return render(
                    request,
                    "confronti/archive.html",
                    _archive_context(request, selected_archive_folder_ids=selected_folder_ids, report_summary_scope="folders"),
                )
            if not selected_reports:
                messages.error(request, "Le cartelle selezionate non contengono report da usare nel sunto.")
                return render(
                    request,
                    "confronti/archive.html",
                    _archive_context(
                        request,
                        selected_archive_folder_ids=selected_folder_ids,
                        selected_archive_folders=selected_folders,
                        report_summary_scope="folders",
                    ),
                )
            report_summary = _build_reports_summary_from_archived_reports(selected_reports)
            report_summary_warnings = report_summary.get("warnings", [])
            if empty_folders:
                folder_labels = ", ".join(folder.customer_name for folder in empty_folders)
                report_summary_warnings = [
                    f"Queste cartelle non contenevano report e sono state saltate: {folder_labels}.",
                    *report_summary_warnings,
                ]
            _store_report_summary_session(request, report_summary)
            messages.success(
                request,
                f"Sunto creato per {report_summary.get('count', 0)} report da {len(selected_folders)} cartelle.",
            )
            return render(
                request,
                "confronti/archive.html",
                _archive_context(
                    request,
                    selected_archive_folder_ids=selected_folder_ids,
                    selected_archive_folders=selected_folders,
                    report_summary=report_summary,
                    report_summary_warnings=report_summary_warnings,
                    report_summary_scope="folders",
                ),
            )
    return render(request, "confronti/archive.html", _archive_context(request))


@login_required
def archivio_report_cartella(request, folder_id):
    if not _is_archive_admin(request.user):
        return _redirect_for_non_archive_admin(request)
    folder = get_object_or_404(CustomerArchiveFolder, pk=folder_id)
    if request.method == "POST":
        action = request.POST.get("action")
        search_query = request.POST.get("q")
        if action == "update_archive_folder":
            folder_form = ArchiveFolderForm(request.POST, instance=folder)
            if folder_form.is_valid():
                folder_form.save()
                messages.success(request, "Dati cartella archivio aggiornati.")
                return redirect(_archive_redirect_url(search_query, folder_id=folder.pk))
            return render(
                request,
                "confronti/archive.html",
                _archive_context(request, selected_folder=folder, folder_form=folder_form),
            )
        if action == "add_archive_report":
            add_report_form = ArchiveReportUploadForm(request.POST, request.FILES, prefix="add-report")
            if add_report_form.is_valid():
                _create_uploaded_archive_report(folder, add_report_form, request.user)
                messages.success(request, f"File aggiunto nella cartella {folder.folder_name}.")
                return redirect(_archive_redirect_url(search_query, folder_id=folder.pk))
            return render(
                request,
                "confronti/archive.html",
                _archive_context(request, selected_folder=folder, add_report_form=add_report_form),
            )
        if action == "build_archive_report_summary":
            selected_report_ids = [int(value) for value in request.POST.getlist("selected_reports") if str(value).isdigit()]
            if not selected_report_ids:
                messages.error(request, "Seleziona almeno un report da includere nel sunto.")
                return render(
                    request,
                    "confronti/archive.html",
                    _archive_context(request, selected_folder=folder, selected_report_ids=selected_report_ids),
                )
            selected_reports = list(
                folder.reports.filter(pk__in=selected_report_ids).order_by("-comparison_datetime", "-created_at")
            )
            if not selected_reports:
                messages.error(request, "I report selezionati non sono disponibili in questa cartella.")
                return render(
                    request,
                    "confronti/archive.html",
                    _archive_context(request, selected_folder=folder),
                )
            report_summary = _build_reports_summary_from_archived_reports(selected_reports)
            report_summary_warnings = report_summary.get("warnings", [])
            _store_report_summary_session(request, report_summary)
            messages.success(request, f"Sunto creato per {report_summary.get('count', 0)} report selezionati.")
            return render(
                request,
                "confronti/archive.html",
                _archive_context(
                    request,
                    selected_folder=folder,
                    selected_report_ids=[report.pk for report in selected_reports],
                    report_summary=report_summary,
                    report_summary_warnings=report_summary_warnings,
                ),
            )
        if action == "update_archive_report":
            report = get_object_or_404(ComparisonReport, pk=request.POST.get("report_id"), folder=folder)
            report_form = ArchiveReportForm(request.POST, instance=report, prefix=f"report-{report.pk}")
            if report_form.is_valid():
                report_form.save()
                _touch_archive_folder(folder)
                messages.success(request, "Report archiviato aggiornato.")
                return redirect(_archive_redirect_url(search_query, folder_id=folder.pk))
            return render(
                request,
                "confronti/archive.html",
                _archive_context(
                    request,
                    selected_folder=folder,
                    active_report_id=report.pk,
                    active_report_form=report_form,
                ),
            )
        if action == "replace_archive_report_file":
            report = get_object_or_404(ComparisonReport, pk=request.POST.get("report_id"), folder=folder)
            replace_form = ArchiveReportReplaceFileForm(request.POST, request.FILES, prefix=f"replace-report-{report.pk}")
            if replace_form.is_valid():
                _replace_archive_report_file(report, replace_form.cleaned_data["report_file"])
                messages.success(request, "File report sostituito.")
                return redirect(_archive_redirect_url(search_query, folder_id=folder.pk))
            return render(
                request,
                "confronti/archive.html",
                _archive_context(
                    request,
                    selected_folder=folder,
                    active_replace_report_id=report.pk,
                    active_replace_report_form=replace_form,
                ),
            )
        if action == "delete_archive_report":
            report = get_object_or_404(ComparisonReport, pk=request.POST.get("report_id"), folder=folder)
            _delete_archive_report(report)
            messages.success(request, "Report archiviato eliminato.")
            return redirect(_archive_redirect_url(search_query, folder_id=folder.pk))
        if action == "delete_archive_folder":
            deleted = _delete_archive_folder(folder)
            messages.success(
                request,
                f"Cartella {deleted['folder_name']} eliminata con {deleted['report_count']} report.",
            )
            return redirect(_archive_redirect_url(search_query))
    return render(request, "confronti/archive.html", _archive_context(request, selected_folder=folder))


@login_required
def scarica_report_archiviato(request, report_id):
    if not _is_archive_admin(request.user):
        return _redirect_for_non_archive_admin(request)
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
def apri_cartella_archivio_locale(request, folder_id):
    if not _is_archive_admin(request.user):
        return _redirect_for_non_archive_admin(request)
    folder = get_object_or_404(CustomerArchiveFolder, pk=folder_id)
    if request.method != "POST":
        return redirect("archivio_report_cartella", folder_id=folder.pk)
    folder_path = _archive_folder_absolute_path(folder)
    try:
        _open_local_archive_path(folder_path)
    except FileNotFoundError:
        messages.error(request, "La cartella archivio non esiste ancora sul disco.")
    except RuntimeError as exc:
        messages.error(request, str(exc))
    except OSError:
        logger.exception("Impossibile aprire la cartella archivio %s", folder_path)
        messages.error(request, "Non sono riuscito ad aprire la cartella file.")
    else:
        messages.success(request, "Cartella file aperta nel computer locale.")
    return redirect("archivio_report_cartella", folder_id=folder.pk)


@login_required
def scarica_sunto_report(request):
    if not is_internal_user(request.user):
        return redirect("accesso_clienti")
    summary = request.session.get(LAST_REPORT_SUMMARY_KEY)
    if not summary or not summary.get("rows"):
        return HttpResponse("Nessun sunto pronto. Carica i report dalla dashboard admin.", status=400)
    content = build_reports_summary_excel(summary)
    nome = safe_download_filename("sunto_report_confronti.xlsx")
    response = HttpResponse(
        content,
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = f'attachment; filename="{nome}"'
    response["Content-Length"] = str(len(content))
    return response


def confronto(request):
    if not request.user.is_authenticated:
        return _home_page(request)
    if not is_internal_user(request.user):
        return redirect("accesso_clienti")
    return _confronto_page(request, operator_mode=is_illumia_operator(request.user))


@login_required
def confronto_cliente_illumia(request):
    return redirect("accesso_clienti")


@login_required
def scarica_excel(request):
    if not is_internal_user(request.user):
        return redirect("accesso_clienti")
    return _scarica_excel(request, operator_mode=is_illumia_operator(request.user))


@login_required
def scarica_excel_cliente_illumia(request):
    return _scarica_excel(request, customer_mode=True)
