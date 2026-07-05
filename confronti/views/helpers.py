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
from django.db.models import Count, Max, Q
from django.shortcuts import get_object_or_404, redirect
from django.urls import reverse
from django.utils import timezone

from ..forms import (
    ArchiveFolderCreateForm,
    ArchiveFolderForm,
    ArchiveFolderMergeForm,
    ArchiveReportForm,
    ArchiveReportMoveForm,
    ArchiveReportReplaceFileForm,
    ArchiveReportUploadForm,
    ConfrontoForm,
)
from ..models import ComparisonReport, CustomerArchiveFolder, InviteCode, archive_report_upload_to
from ..roles import is_illumia_operator, is_internal_user
from ..services import (
    build_excel_bytes,
    build_reports_summary,
    provider_label,
    safe_download_filename,
)


logger = logging.getLogger(__name__)
LAST_UPLOADED_BILL_NAME_KEY = "last_uploaded_bill_name"
LAST_UPLOADED_BILL_NAME_CLIENT_KEY = "last_uploaded_bill_name_cliente_illumia"
LAST_COMPARISON_KEY = "last_confronto"
LAST_COMPARISON_CLIENT_KEY = "last_confronto_cliente_illumia"
LAST_ARCHIVED_REPORT_KEY = "last_archived_report_id"
LAST_ARCHIVED_REPORT_CLIENT_KEY = "last_archived_report_id_cliente_illumia"
LAST_REPORT_SUMMARY_KEY = "last_report_summary"
WHATSAPP_SENDER_NUMBER = os.environ.get("WHATSAPP_SENDER_NUMBER", "3271044102")
WHATSAPP_WEB_HOME_URL = "https://web.whatsapp.com/"
CUSTOMER_PORTAL_BASE_URL = (
    os.environ.get("CUSTOMER_PORTAL_BASE_URL")
    or (f"https://{os.environ['RENDER_EXTERNAL_HOSTNAME']}" if os.environ.get("RENDER_EXTERNAL_HOSTNAME") else "")
    or "https://energia-solidale.onrender.com"
).rstrip("/")


def display_upload_name(raw_name):
    return (raw_name or "").replace("\\", "/").split("/")[-1].strip()


def public_customer_url(path):
    return urljoin(f"{CUSTOMER_PORTAL_BASE_URL}/", path.lstrip("/"))


def customer_registration_url(invite_code=""):
    base_url = public_customer_url(reverse("register"))
    normalized_code = InviteCode.normalize_code(invite_code)
    if not normalized_code:
        return base_url
    return f"{base_url}?invite_code={normalized_code}"


def build_whatsapp_invite(customer_name, customer_phone, invite_code):
    registration_url = customer_registration_url(invite_code)
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


def mode_config(customer_mode=False, operator_mode=False):
    if customer_mode:
        return {
            "comparison_session_key": LAST_COMPARISON_CLIENT_KEY,
            "archive_report_session_key": LAST_ARCHIVED_REPORT_CLIENT_KEY,
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
            "archive_report_session_key": LAST_ARCHIVED_REPORT_KEY,
            "upload_name_session_key": LAST_UPLOADED_BILL_NAME_KEY,
            "page_title": "Confronto bollette",
            "page_intro": "",
            "download_url_name": "scarica_excel",
        }
    return {
        "comparison_session_key": LAST_COMPARISON_KEY,
        "archive_report_session_key": LAST_ARCHIVED_REPORT_KEY,
        "upload_name_session_key": LAST_UPLOADED_BILL_NAME_KEY,
        "page_title": "Confronto bollette",
        "page_intro": "",
        "download_url_name": "scarica_excel",
    }


def force_illumia_only_data(data, clear_pod_pdr=False):
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


def force_customer_mode_data(data):
    return force_illumia_only_data(data, clear_pod_pdr=True)


def force_operator_mode_data(data):
    from ..services import normalize_providers
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


def comparison_form(*args, customer_mode=False, operator_mode=False, **kwargs):
    return ConfrontoForm(*args, customer_mode=customer_mode, operator_mode=operator_mode, **kwargs)


def customer_optional_section_open(form):
    if not getattr(form, "customer_mode", False):
        return False
    return any(field_name in form.errors for field_name in ConfrontoForm.CUSTOMER_OPTIONAL_FIELDS)


def customer_status_snapshot():
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


def is_archive_admin(user):
    return bool(
        getattr(user, "is_authenticated", False)
        and (getattr(user, "is_superuser", False) or (getattr(user, "is_staff", False) and not is_illumia_operator(user)))
    )


def redirect_for_non_archive_admin(request):
    if request.user.is_authenticated and is_internal_user(request.user):
        return redirect("confronto")
    return redirect("accesso_clienti")


def parse_comparison_datetime(value):
    if not value:
        return timezone.localtime()
    try:
        parsed = datetime.fromisoformat(str(value))
    except ValueError:
        return timezone.localtime()
    if timezone.is_naive(parsed):
        parsed = timezone.make_aware(parsed, timezone.get_current_timezone())
    return parsed


def create_archive_folder(customer_name, user, customer_email="", customer_phone="", notes=""):
    folder_label = CustomerArchiveFolder.normalize_customer_name(customer_name) or "Nuova cartella"
    return CustomerArchiveFolder.objects.create(
        customer_name=folder_label,
        customer_email=CustomerArchiveFolder.normalize_customer_email(customer_email),
        customer_phone=CustomerArchiveFolder.normalize_customer_phone(customer_phone),
        notes=str(notes or "").strip(),
        created_by=user,
    )


def create_archive_folder_from_report_data(data, user, folder_name=""):
    return create_archive_folder(
        folder_name or data.get("nome_cliente"),
        user,
        customer_email=data.get("email_cliente", ""),
        customer_phone=data.get("telefono_cliente", ""),
    )


def archive_report_payload(data, prepared):
    return {
        "customer_name": data.get("nome_cliente", ""),
        "pod_pdr": data.get("pod_pdr", ""),
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


def archive_report_filename(data):
    provider_name = "_".join(provider_label(provider).lower().replace(".", "") for provider in data.get("providers", []))
    provider_name = provider_name or provider_label(data.get("provider", "ILLUMIA")).lower().replace(".", "")
    return safe_download_filename(
        f"confronto_{provider_name}_{data.get('nome_cliente', 'Cliente')}_{data.get('commodity', 'ENERGIA')}.xlsx"
    )


def archive_report_title(data, prepared):
    provider_name = prepared["calc"].get("providers_label", "Illumia")
    commodity = prepared["calc"].get("commodity_label", data.get("commodity", ""))
    period = prepared["calc"].get("period_label", "")
    title_parts = [provider_name, commodity]
    if period:
        title_parts.append(period)
    return " - ".join(part for part in title_parts if part)


def archive_report_queryset(search_query=""):
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


def archive_reports_queryset(search_query=""):
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


def archive_summary(search_query=""):
    report_qs = archive_reports_queryset(search_query)
    query = str(search_query or "").strip()
    return {
        "folder_count": archive_report_queryset(query).count(),
        "report_count": report_qs.count(),
        "unassigned_count": report_qs.filter(folder__isnull=True).count(),
        "latest_report_at": report_qs.aggregate(last_saved=Max("created_at"))["last_saved"],
    }


def store_report_summary_session(request, report_summary):
    request.session[LAST_REPORT_SUMMARY_KEY] = {
        "columns": report_summary.get("columns", []),
        "rows": report_summary.get("rows", []),
        "warnings": report_summary.get("warnings", []),
        "count": report_summary.get("count", 0),
    }


def archive_folder_absolute_path(folder):
    return os.path.join(str(settings.MEDIA_ROOT), "report_archive", folder.folder_name)


def archive_report_storage_folder_name(report):
    if report.folder_id and getattr(report, "folder", None) is not None:
        return report.folder.folder_name
    return "non-assegnati"


def archive_report_absolute_path(report):
    try:
        return report.report_file.path
    except (NotImplementedError, ValueError):
        return os.path.join(
            str(settings.MEDIA_ROOT),
            "report_archive",
            archive_report_storage_folder_name(report),
            os.path.basename(report.report_file.name),
        )


def can_open_local_archive_path():
    if os.environ.get("RENDER_EXTERNAL_HOSTNAME"):
        return False
    if os.name == "nt":
        return True
    return bool(shutil.which("open") or shutil.which("xdg-open"))


def archive_storage_label():
    if os.environ.get("RENDER_EXTERNAL_HOSTNAME"):
        return "Cloud app (Render + disco persistente)"
    return "Computer locale"


def open_local_archive_path(path):
    if not can_open_local_archive_path():
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


def touch_archive_folder(folder):
    if folder is None:
        return
    CustomerArchiveFolder.objects.filter(pk=folder.pk).update(updated_at=timezone.now())


def archive_uploaded_report_title(filename):
    stem = os.path.splitext(display_upload_name(filename))[0]
    normalized = " ".join(stem.replace("_", " ").replace("-", " ").split())
    return normalized or "Report confronto"


def create_uploaded_archive_report(folder, form, user):
    uploaded_file = form.cleaned_data["report_file"]
    report = ComparisonReport(
        folder=folder,
        title=str(form.cleaned_data.get("title") or "").strip() or archive_uploaded_report_title(uploaded_file.name),
        notes=str(form.cleaned_data.get("notes") or "").strip(),
        original_filename=display_upload_name(uploaded_file.name),
        comparison_datetime=timezone.now(),
        created_by=user,
    )
    report.report_file.save(report.original_filename, uploaded_file, save=False)
    report.save()
    touch_archive_folder(folder)
    return report


def archive_unique_storage_name(folder, filename, storage, reserved_names=None):
    normalized_filename = display_upload_name(filename) or "report.xlsx"
    reserved_names = reserved_names if reserved_names is not None else set()
    temp_report = ComparisonReport(folder=folder)
    stem, ext = os.path.splitext(normalized_filename)
    suffix = 1
    while True:
        candidate_filename = normalized_filename if suffix == 1 else f"{stem}-{suffix}{ext}"
        candidate_name = archive_report_upload_to(temp_report, candidate_filename)
        if candidate_name not in reserved_names and not storage.exists(candidate_name):
            reserved_names.add(candidate_name)
            return candidate_name
        suffix += 1


def move_archive_report_to_folder(report, target_folder, reserved_names=None):
    old_folder = report.folder
    if old_folder is None and target_folder is None:
        return report.report_file.name
    if old_folder is not None and target_folder is not None and old_folder.pk == target_folder.pk:
        return report.report_file.name
    old_name = report.report_file.name
    storage = report.report_file.storage
    source_name = report.original_filename or os.path.basename(old_name) or f"report-{report.pk}.xlsx"
    target_name = archive_unique_storage_name(target_folder, source_name, storage, reserved_names=reserved_names)
    report.report_file.open("rb")
    try:
        content = report.report_file.read()
    finally:
        report.report_file.close()
    saved_name = storage.save(target_name, ContentFile(content))
    report.folder = target_folder
    report.report_file.name = saved_name
    report.save(update_fields=["folder", "report_file", "updated_at"])
    if old_name and old_name != saved_name:
        try:
            storage.delete(old_name)
        except OSError:
            logger.warning("Impossibile eliminare il vecchio file archivio %s", old_name, exc_info=True)
    touch_archive_folder(old_folder)
    if target_folder is not None and (old_folder is None or old_folder.pk != target_folder.pk):
        touch_archive_folder(target_folder)
    return saved_name


def merge_archive_folder_notes(target_notes, source_notes, source_folder_name):
    normalized_target = str(target_notes or "").strip()
    normalized_source = str(source_notes or "").strip()
    if not normalized_source:
        return normalized_target
    if not normalized_target:
        return normalized_source
    if normalized_source in normalized_target:
        return normalized_target
    return f"{normalized_target}\n\nDa cartella unita {source_folder_name}:\n{normalized_source}"


def merge_archive_folders(target_folder, source_folder):
    if target_folder.pk == source_folder.pk:
        raise ValueError("La cartella sorgente deve essere diversa da quella di destinazione.")
    source_folder_name = source_folder.folder_name
    source_folder_path = archive_folder_absolute_path(source_folder)
    reserved_names = set(target_folder.reports.values_list("report_file", flat=True))
    moved_reports = []
    source_reports = list(source_folder.reports.order_by("-comparison_datetime", "-created_at"))
    for report in source_reports:
        move_archive_report_to_folder(report, target_folder, reserved_names=reserved_names)
        moved_reports.append(report.pk)
    from django.core.files.base import ContentFile

    changed_fields = []
    if not target_folder.customer_email and source_folder.customer_email:
        target_folder.customer_email = source_folder.customer_email
        changed_fields.append("customer_email")
    if not target_folder.customer_phone and source_folder.customer_phone:
        target_folder.customer_phone = source_folder.customer_phone
        changed_fields.append("customer_phone")
    merged_notes = merge_archive_folder_notes(target_folder.notes, source_folder.notes, source_folder.folder_name)
    if merged_notes != target_folder.notes:
        target_folder.notes = merged_notes
        changed_fields.append("notes")
    if changed_fields:
        target_folder.save(update_fields=changed_fields + ["updated_at"])
    else:
        touch_archive_folder(target_folder)

    source_folder.delete()
    if os.path.isdir(source_folder_path):
        shutil.rmtree(source_folder_path, ignore_errors=True)
    return {
        "target_folder_name": target_folder.folder_name,
        "source_folder_name": source_folder_name,
        "moved_report_count": len(moved_reports),
    }


def replace_archive_report_file(report, uploaded_file):
    old_name = report.report_file.name
    report.report_file.save(display_upload_name(uploaded_file.name), uploaded_file, save=False)
    report.original_filename = display_upload_name(uploaded_file.name)
    if not report.title:
        report.title = archive_uploaded_report_title(uploaded_file.name)
    if report.comparison_datetime is None:
        report.comparison_datetime = timezone.now()
    report.save()
    if old_name and old_name != report.report_file.name:
        try:
            report.report_file.storage.delete(old_name)
        except OSError:
            logger.warning("Impossibile eliminare il vecchio file archivio %s", old_name, exc_info=True)
    touch_archive_folder(report.folder)


def delete_empty_archive_folder_if_needed(folder):
    if folder is None or folder.reports.exists():
        return
    folder_path = archive_folder_absolute_path(folder)
    folder.delete()
    if os.path.isdir(folder_path):
        shutil.rmtree(folder_path, ignore_errors=True)


def delete_archive_report(report):
    stored_name = report.report_file.name
    folder = report.folder
    report.delete()
    if stored_name:
        try:
            report.report_file.storage.delete(stored_name)
        except OSError:
            logger.warning("Impossibile eliminare il file archivio %s", stored_name, exc_info=True)
    touch_archive_folder(folder)


def delete_archive_folder(folder):
    reports = list(folder.reports.all())
    folder_name = folder.folder_name
    folder_path = archive_folder_absolute_path(folder)
    report_count = len(reports)
    reserved_names = set(ComparisonReport.objects.filter(folder__isnull=True).values_list("report_file", flat=True))
    for report in reports:
        move_archive_report_to_folder(report, None, reserved_names=reserved_names)
    folder.delete()
    if os.path.isdir(folder_path):
        shutil.rmtree(folder_path, ignore_errors=True)
    return {"folder_name": folder_name, "report_count": report_count}


def archive_report_move_message(target_folder):
    if target_folder is None:
        return "Report lasciato senza cartella."
    return f"Report spostato nella cartella {target_folder.customer_name}."


def build_reports_summary_from_archived_reports(reports):
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


def create_archived_comparison_report(data, prepared, user, folder=None):
    from django.core.files.base import ContentFile
    target_folder = folder
    content = build_excel_bytes(data, prepared)
    report = ComparisonReport(
        folder=target_folder,
        title=archive_report_title(data, prepared),
        commodity=data.get("commodity", ""),
        providers_label=prepared["calc"].get("providers_label", ""),
        bill_period_label=prepared["calc"].get("period_label", ""),
        comparison_datetime=parse_comparison_datetime(data.get("comparison_datetime")),
        original_filename=archive_report_filename(data),
        comparison_data=archive_report_payload(data, prepared),
        created_by=user,
    )
    report.report_file.save(report.original_filename, ContentFile(content), save=False)
    report.save()
    touch_archive_folder(target_folder)
    return report


def archive_report_source_data(report):
    comparison_data = report.comparison_data if isinstance(report.comparison_data, dict) else {}
    return {
        "nome_cliente": comparison_data.get("customer_name") or (report.folder.customer_name if report.folder_id else ""),
        "email_cliente": comparison_data.get("customer_email", ""),
        "telefono_cliente": comparison_data.get("customer_phone", ""),
    }


def selected_archive_folder_ids(values):
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


def selected_archive_folders_and_reports(folder_ids):
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


def archive_report_customer_name(report):
    comparison_data = report.comparison_data if isinstance(report.comparison_data, dict) else {}
    customer_name = str(comparison_data.get("customer_name") or "").strip()
    if customer_name:
        return customer_name
    if report.folder_id and getattr(report, "folder", None) is not None:
        return report.folder.customer_name
    return ""


def archive_report_pod_pdr(report):
    comparison_data = report.comparison_data if isinstance(report.comparison_data, dict) else {}
    return str(comparison_data.get("pod_pdr") or "").strip()


def archive_report_supply_address(report):
    comparison_data = report.comparison_data if isinstance(report.comparison_data, dict) else {}
    return str(comparison_data.get("supply_address") or "").strip()


def archive_report_commodity_label(report):
    comparison_data = report.comparison_data if isinstance(report.comparison_data, dict) else {}
    commodity = str(comparison_data.get("commodity") or report.commodity or "").strip().upper()
    if commodity == "GAS":
        return "Gas"
    if commodity == "EE":
        return "Luce"
    return "N.D."


def archive_report_reference_datetime(report):
    return report.comparison_datetime or report.created_at


def archive_report_entry(
    report,
    active_move_report_id=None,
    active_move_form=None,
    active_report_id=None,
    active_report_form=None,
    active_replace_report_id=None,
    active_replace_report_form=None,
):
    move_form = (
        active_move_form
        if active_move_report_id == report.pk and active_move_form is not None
        else ArchiveReportMoveForm(prefix=f"move-report-{report.pk}", current_folder=report.folder)
    )
    report_form = (
        active_report_form
        if active_report_id == report.pk and active_report_form is not None
        else ArchiveReportForm(instance=report, prefix=f"report-{report.pk}")
    )
    replace_form = (
        active_replace_report_form
        if active_replace_report_id == report.pk and active_replace_report_form is not None
        else ArchiveReportReplaceFileForm(prefix=f"replace-report-{report.pk}")
    )
    customer_name = archive_report_customer_name(report)
    pod_pdr = archive_report_pod_pdr(report)
    supply_address = archive_report_supply_address(report)
    commodity_label = archive_report_commodity_label(report)
    return {
        "report": report,
        "form": report_form,
        "replace_form": replace_form,
        "move_form": move_form,
        "absolute_path": archive_report_absolute_path(report),
        "display_title": report.title or "Report confronto",
        "customer_name": customer_name,
        "customer_name_display": customer_name or "Cliente non indicato",
        "pod_pdr": pod_pdr,
        "pod_pdr_display": pod_pdr or "N.D.",
        "supply_address": supply_address,
        "supply_address_display": supply_address or "N.D.",
        "commodity_label": commodity_label,
        "reference_datetime": archive_report_reference_datetime(report),
    }


def archive_context(request, selected_folder=None, selected_report=None, folder_form=None, create_folder_form=None, merge_form=None, add_report_form=None, selected_report_ids=None, selected_archive_folder_ids=None, selected_archive_folders=None, report_summary=None, report_summary_warnings=None, report_summary_scope="reports", active_report_id=None, active_report_form=None, active_replace_report_id=None, active_replace_report_form=None, active_move_report_id=None, active_move_report_form=None):
    search_query = str(request.GET.get("q") or request.POST.get("q") or "").strip()
    folders = list(archive_report_queryset(search_query))
    report_entries = []
    archived_report_entries = []
    selected_report_entry = None
    selected_report_ids = [int(report_id) for report_id in (selected_report_ids or [])]
    selected_archive_folder_ids = [int(folder_id) for folder_id in (selected_archive_folder_ids or [])]
    selected_archive_folders = list(selected_archive_folders or [])
    report_summary_warnings = list(report_summary_warnings or [])
    archive_folder_absolute_path = ""
    archive_local_open_available = can_open_local_archive_path()
    selected_report_lookup_id = getattr(selected_report, "pk", None)
    if selected_report_lookup_id is None:
        raw_selected_report_id = request.GET.get("report_id") or request.POST.get("selected_report_id")
        if str(raw_selected_report_id).isdigit():
            selected_report_lookup_id = int(raw_selected_report_id)
    if create_folder_form is None:
        create_folder_form = ArchiveFolderCreateForm(prefix="create-folder")
    for report in archive_reports_queryset(search_query):
        entry = archive_report_entry(
            report,
            active_move_report_id=active_move_report_id,
            active_move_form=active_move_report_form,
            active_report_id=active_report_id,
            active_report_form=active_report_form,
            active_replace_report_id=active_replace_report_id,
            active_replace_report_form=active_replace_report_form,
        )
        archived_report_entries.append(entry)
        if selected_folder is None and selected_report_lookup_id == report.pk:
            selected_report = report
            selected_report_entry = entry
    if selected_folder is not None:
        selected_folder = get_object_or_404(
            CustomerArchiveFolder.objects.prefetch_related("reports").annotate(report_count=Count("reports")),
            pk=selected_folder.pk,
        )
        archive_folder_absolute_path = archive_folder_absolute_path(selected_folder)
        reports = list(selected_folder.reports.order_by("-comparison_datetime", "-created_at"))
        if folder_form is None:
            folder_form = ArchiveFolderForm(instance=selected_folder)
        if merge_form is None:
            merge_form = ArchiveFolderMergeForm(prefix="merge-folder", target_folder=selected_folder)
        if add_report_form is None:
            add_report_form = ArchiveReportUploadForm(prefix="add-report")
        for report in reports:
            entry = archive_report_entry(
                report,
                active_move_report_id=active_move_report_id,
                active_move_form=active_move_report_form,
                active_report_id=active_report_id,
                active_report_form=active_report_form,
                active_replace_report_id=active_replace_report_id,
                active_replace_report_form=active_replace_report_form,
            )
            report_entries.append(entry)
            if selected_report_lookup_id == report.pk:
                selected_report = report
                selected_report_entry = entry
    elif selected_report is not None and selected_report_entry is None:
        selected_report_entry = archive_report_entry(
            selected_report,
            active_move_report_id=active_move_report_id,
            active_move_form=active_move_report_form,
            active_report_id=active_report_id,
            active_report_form=active_report_form,
            active_replace_report_id=active_replace_report_id,
            active_replace_report_form=active_replace_report_form,
        )
    return {
        "brand_home_url": reverse("confronto"),
        "page_title": "Archivio report",
        "admin_tabs": admin_tabs("archivio-report"),
        "archive_query": search_query,
        "archive_summary": archive_summary(search_query),
        "archive_folders": folders,
        "selected_folder": selected_folder,
        "selected_report": selected_report,
        "selected_report_entry": selected_report_entry,
        "folder_form": folder_form,
        "create_folder_form": create_folder_form,
        "merge_form": merge_form,
        "add_report_form": add_report_form,
        "archive_storage_label": archive_storage_label(),
        "archive_folder_absolute_path": archive_folder_absolute_path,
        "archive_local_open_available": archive_local_open_available,
        "selected_report_ids": selected_report_ids,
        "selected_archive_folder_ids": selected_archive_folder_ids,
        "selected_archive_folders": selected_archive_folders,
        "archive_report_summary": report_summary,
        "archive_report_summary_warnings": report_summary_warnings,
        "archive_report_summary_scope": report_summary_scope,
        "report_entries": report_entries,
        "archived_report_entries": archived_report_entries,
    }


def archive_redirect_url(search_query="", folder_id=None, report_id=None):
    if folder_id is not None:
        base_url = reverse("archivio_report_cartella", kwargs={"folder_id": folder_id})
    else:
        base_url = reverse("archivio_report")
    query_params = {}
    query = str(search_query or "").strip()
    if query:
        query_params["q"] = query
    if report_id is not None:
        query_params["report_id"] = report_id
    if not query_params:
        return base_url
    return f"{base_url}?{urlencode(query_params)}"


def public_access_page(request):
    from django.shortcuts import render
    return render(
        request,
        "registration/access_choice.html",
        {
            "brand_home_url": reverse("accesso_clienti"),
            "login_url": reverse("login"),
            "register_url": reverse("register"),
        },
    )


def home_tabs():
    root_url = reverse("confronto")
    login_url = reverse("login")
    return [
        {
            "key": "confronto",
            "label": "Esegui confronto",
            "href": f"{login_url}?{urlencode({'next': f'{root_url}?panel=confronto'})}",
        },
        {
            "key": "genera-codici",
            "label": "Genera Codice Invito",
            "href": f"{login_url}?{urlencode({'next': f'{root_url}?panel=genera-codici'})}",
        },
        {
            "key": "sunto-report",
            "label": "Sunto Report",
            "href": f"{login_url}?{urlencode({'next': f'{root_url}?panel=sunto-report'})}",
        },
        {
            "key": "clienti",
            "label": "Clienti",
            "subs": [
                {"label": "Stato Clienti", "href": f"{login_url}?{urlencode({'next': f'{root_url}?panel=status-clienti'})}"},
                {"label": "Apri Dashboard", "href": reverse("accesso_clienti")},
                {"label": "Apri Registrazione Cliente", "href": reverse("register")},
            ],
        },
    ]


def normalize_admin_focus_panel(panel):
    return panel if panel in {"confronto", "genera-codici", "status-clienti", "sunto-report", "archivio-report"} else "confronto"


def admin_tabs(active_panel, operator_mode=False):
    root_url = reverse("confronto")
    if operator_mode:
        return [{"key": "sunto-report", "label": "Sunto report", "href": f"{root_url}?panel=sunto-report#sunto-report", "active": active_panel == "sunto-report"}]
    return [
        {"key": "confronto", "label": "Esegui confronto", "href": f"{root_url}?panel=confronto#confronto-bollette-offerte", "active": active_panel == "confronto"},
        {"key": "genera-codici", "label": "Genera Codice Invito", "href": f"{root_url}?panel=genera-codici#genera-codici", "active": active_panel == "genera-codici"},
        {"key": "sunto-report", "label": "Sunto Report", "href": f"{root_url}?panel=sunto-report#sunto-report", "active": active_panel == "sunto-report"},
        {
            "key": "clienti",
            "label": "Clienti",
            "subs": [
                {"label": "Stato Clienti", "href": f"{root_url}?panel=status-clienti#stato-clienti"},
                {"label": "Apri Dashboard", "href": reverse("accesso_clienti")},
                {"label": "Apri Registrazione Cliente", "href": reverse("register")},
            ],
        },
        {"key": "archivio-report", "label": "Apri archivio file", "href": reverse("archivio_report"), "active": active_panel == "archivio-report"},
    ]


def home_page(request):
    from django.shortcuts import render
    return render(
        request,
        "confronti/home.html",
        {
            "brand_home_url": reverse("confronto"),
            "page_title": "Energia Solidale",
            "hide_header_brand": True,
            "admin_tabs": home_tabs(),
            "customer_dashboard_url": reverse("accesso_clienti"),
        },
    )
