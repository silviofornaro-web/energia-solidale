import logging
import os
import shutil
import subprocess

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.core.files.base import ContentFile
from django.db.models import Count, Q
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse
from django.utils import timezone

from ..forms import (
    ArchiveCurrentReportSaveForm,
    ArchiveFolderCreateForm,
    ArchiveFolderForm,
    ArchiveFolderMergeForm,
    ArchiveReportForm,
    ArchiveReportMoveForm,
    ArchiveReportReplaceFileForm,
    ArchiveReportUploadForm,
)
from ..models import ComparisonReport, CustomerArchiveFolder
from ..roles import is_internal_user
from ..services import prepare_comparison, session_to_service_data
from .helpers import (
    LAST_COMPARISON_KEY,
    LAST_ARCHIVED_REPORT_KEY,
    archive_context,
    archive_redirect_url,
    archive_report_source_data,
    can_open_local_archive_path,
    create_archive_folder,
    create_archive_folder_from_report_data,
    create_archived_comparison_report,
    delete_archive_folder,
    delete_archive_report,
    delete_empty_archive_folder_if_needed,
    is_archive_admin,
    merge_archive_folders,
    move_archive_report_to_folder,
    open_local_archive_path,
    replace_archive_report_file,
    redirect_for_non_archive_admin,
    selected_archive_folder_ids,
    selected_archive_folders_and_reports,
    touch_archive_folder,
    build_reports_summary_from_archived_reports,
    store_report_summary_session,
    archive_folder_absolute_path,
)


logger = logging.getLogger(__name__)


@login_required
def archivia_report_corrente(request):
    if not is_internal_user(request.user):
        return redirect("accesso_clienti")
    if request.method != "POST":
        return redirect("confronto")

    raw = request.session.get(LAST_COMPARISON_KEY)
    archived_report = None
    archived_report_id = request.session.get(LAST_ARCHIVED_REPORT_KEY)
    if archived_report_id:
        archived_report = ComparisonReport.objects.select_related("folder").filter(pk=archived_report_id).first()

    default_folder_name = ""
    if archived_report is not None:
        default_folder_name = archive_report_source_data(archived_report).get("nome_cliente", "")
    elif raw:
        default_folder_name = session_to_service_data(raw).get("nome_cliente", "")

    save_form = ArchiveCurrentReportSaveForm(request.POST, default_folder_name=default_folder_name)
    if not save_form.is_valid():
        messages.error(request, "Non sono riuscito a leggere la destinazione archivio del report.")
        return redirect(f"{reverse('confronto')}?panel=confronto")

    folder = save_form.cleaned_data.get("existing_folder")
    if save_form.cleaned_data.get("new_folder_name"):
        source_data = archive_report_source_data(archived_report) if archived_report is not None else {}
        if not source_data and raw:
            source_data = session_to_service_data(raw)
        folder = create_archive_folder_from_report_data(data=source_data, user=request.user, folder_name=save_form.cleaned_data["new_folder_name"])

    if archived_report is not None:
        previous_folder = archived_report.folder
        if previous_folder is None and folder is None:
            messages.success(request, "Report gia archiviato senza cartella.")
        elif previous_folder is not None and folder is not None and previous_folder.pk == folder.pk:
            messages.success(request, f"Report gia archiviato nella cartella {folder.customer_name}.")
        else:
            move_archive_report_to_folder(archived_report, folder)
            delete_empty_archive_folder_if_needed(previous_folder)
            if folder is not None:
                messages.success(request, f"Report archiviato nella cartella {folder.customer_name}.")
            else:
                messages.success(request, "Report archiviato senza cartella. Puoi assegnarlo dopo dall'archivio.")
        if is_archive_admin(request.user) and folder is not None:
            return redirect("archivio_report_cartella", folder_id=folder.pk)
        if is_archive_admin(request.user):
            return redirect("archivio_report")
        return redirect(f"{reverse('confronto')}?panel=confronto")

    if not raw:
        messages.error(request, "Non c'e nessun confronto interno pronto da archiviare.")
        return redirect(f"{reverse('confronto')}?panel=confronto")

    data = session_to_service_data(raw)
    prepared = prepare_comparison(data)
    report = create_archived_comparison_report(data, prepared, request.user, folder=folder)
    request.session[LAST_ARCHIVED_REPORT_KEY] = report.pk
    if folder is not None:
        messages.success(request, f"Report archiviato nella cartella {folder.customer_name}.")
    else:
        messages.success(request, "Report archiviato senza cartella. Puoi assegnarlo dopo dall'archivio.")
    if is_archive_admin(request.user) and report.folder_id is not None:
        return redirect("archivio_report_cartella", folder_id=report.folder_id)
    if is_archive_admin(request.user):
        return redirect("archivio_report")
    return redirect(f"{reverse('confronto')}?panel=confronto")


@login_required
def archivio_report(request):
    if not is_archive_admin(request.user):
        return redirect_for_non_archive_admin(request)

    if request.method == "POST":
        action = request.POST.get("action")
        search_query = request.POST.get("q")
        selected_report_id = request.POST.get("selected_report_id")

        if action == "create_archive_folder":
            selected_report = ComparisonReport.objects.select_related("folder").filter(pk=selected_report_id).first()
            create_folder_form = ArchiveFolderCreateForm(request.POST, prefix="create-folder")
            if create_folder_form.is_valid():
                folder = create_folder_form.save(commit=False)
                folder.created_by = request.user
                folder.save()
                messages.success(request, f"Cartella {folder.customer_name} creata.")
                if selected_report is not None:
                    if selected_report.folder_id is not None:
                        return redirect(archive_redirect_url(search_query, folder_id=selected_report.folder_id, report_id=selected_report.pk))
                    return redirect(archive_redirect_url(search_query, report_id=selected_report.pk))
                return redirect(archive_redirect_url(search_query, folder_id=folder.pk))
            return render(request, "confronti/archive.html", archive_context(request, create_folder_form=create_folder_form, selected_folder=selected_report.folder if selected_report and selected_report.folder_id else None, selected_report=selected_report))

        if action == "delete_archive_report_global":
            report = get_object_or_404(ComparisonReport.objects.select_related("folder"), pk=request.POST.get("report_id"))
            folder_label = report.folder.customer_name if report.folder else "senza cartella"
            delete_archive_report(report)
            messages.success(request, f"Report eliminato dall'archivio ({folder_label}).")
            return redirect(archive_redirect_url(search_query))

        if action == "update_archive_report_global":
            report = get_object_or_404(ComparisonReport.objects.select_related("folder"), pk=request.POST.get("report_id"), folder__isnull=True)
            report_form = ArchiveReportForm(request.POST, instance=report, prefix=f"report-{report.pk}")
            if report_form.is_valid():
                report_form.save()
                messages.success(request, "File archivio aggiornato.")
                return redirect(archive_redirect_url(search_query, report_id=report.pk))
            return render(request, "confronti/archive.html", archive_context(request, selected_report=report, active_report_id=report.pk, active_report_form=report_form))

        if action == "replace_archive_report_file_global":
            report = get_object_or_404(ComparisonReport.objects.select_related("folder"), pk=request.POST.get("report_id"), folder__isnull=True)
            replace_form = ArchiveReportReplaceFileForm(request.POST, request.FILES, prefix=f"replace-report-{report.pk}")
            if replace_form.is_valid():
                replace_archive_report_file(report, replace_form.cleaned_data["report_file"])
                messages.success(request, "File report sostituito.")
                return redirect(archive_redirect_url(search_query, report_id=report.pk))
            return render(request, "confronti/archive.html", archive_context(request, selected_report=report, active_replace_report_id=report.pk, active_replace_report_form=replace_form))

        if action == "move_archive_report_global":
            report = get_object_or_404(ComparisonReport.objects.select_related("folder"), pk=request.POST.get("report_id"))
            move_form = ArchiveReportMoveForm(request.POST, prefix=f"move-report-{report.pk}", current_folder=report.folder)
            if move_form.is_valid():
                target_folder = move_form.cleaned_data["destination_folder"]
                move_archive_report_to_folder(report, target_folder)
                messages.success(request, _move_message(target_folder))
                if target_folder is not None:
                    return redirect(archive_redirect_url(search_query, folder_id=target_folder.pk))
                return redirect(archive_redirect_url(search_query, report_id=report.pk))
            return render(request, "confronti/archive.html", archive_context(request, selected_report=report if report.folder_id is None else None, active_move_report_id=report.pk, active_move_report_form=move_form))

        if action == "delete_archive_folder":
            folder = get_object_or_404(CustomerArchiveFolder, pk=request.POST.get("folder_id"))
            deleted = delete_archive_folder(folder)
            messages.success(request, f"Cartella {deleted['folder_name']} eliminata. {deleted['report_count']} report restano senza cartella.")
            return redirect(archive_redirect_url(search_query, report_id=selected_report_id or None))

        if action == "merge_selected_archive_folders":
            selected_ids = selected_archive_folder_ids(request.POST.getlist("selected_folder_ids"))
            if len(selected_ids) != 2:
                messages.error(request, "Seleziona esattamente 2 cartelle cliente da unire.")
                return render(request, "confronti/archive.html", archive_context(request, selected_archive_folder_ids=selected_ids))
            try:
                target_folder_id = int(request.POST.get("merge_target_folder_id"))
            except (TypeError, ValueError):
                target_folder_id = 0
            if target_folder_id not in selected_ids:
                messages.error(request, "Scegli quale delle 2 cartelle selezionate deve restare attiva.")
                return render(request, "confronti/archive.html", archive_context(request, selected_archive_folder_ids=selected_ids))
            folders_by_id = CustomerArchiveFolder.objects.in_bulk(selected_ids)
            if len(folders_by_id) != 2:
                messages.error(request, "Le cartelle selezionate non sono piu disponibili nell'archivio.")
                return render(request, "confronti/archive.html", archive_context(request, selected_archive_folder_ids=selected_ids))
            target_folder = folders_by_id[target_folder_id]
            source_folder_id = next(fid for fid in selected_ids if fid != target_folder_id)
            source_folder = folders_by_id[source_folder_id]
            merged = merge_archive_folders(target_folder, source_folder)
            messages.success(request, f"Cartella {merged['source_folder_name']} unita in {merged['target_folder_name']} con {merged['moved_report_count']} report spostati.")
            return redirect(archive_redirect_url(search_query, folder_id=target_folder.pk))

        if action == "build_archive_folder_summary":
            selected_ids = selected_archive_folder_ids(request.POST.getlist("selected_folder_ids"))
            if not selected_ids:
                messages.error(request, "Seleziona almeno una cartella cliente da includere nel sunto.")
                return render(request, "confronti/archive.html", archive_context(request, selected_archive_folder_ids=selected_ids, report_summary_scope="folders"))
            selected_folders, selected_reports, empty_folders = selected_archive_folders_and_reports(selected_ids)
            if not selected_folders:
                messages.error(request, "Le cartelle selezionate non sono disponibili nell'archivio.")
                return render(request, "confronti/archive.html", archive_context(request, selected_archive_folder_ids=selected_ids, report_summary_scope="folders"))
            if not selected_reports:
                messages.error(request, "Le cartelle selezionate non contengono report da usare nel sunto.")
                return render(request, "confronti/archive.html", archive_context(request, selected_archive_folder_ids=selected_ids, selected_archive_folders=selected_folders, report_summary_scope="folders"))
            report_summary = build_reports_summary_from_archived_reports(selected_reports)
            report_summary_warnings = report_summary.get("warnings", [])
            if empty_folders:
                folder_labels = ", ".join(f.customer_name for f in empty_folders)
                report_summary_warnings = [f"Queste cartelle non contenevano report e sono state saltate: {folder_labels}.", *report_summary_warnings]
            store_report_summary_session(request, report_summary)
            messages.success(request, f"Sunto creato per {report_summary.get('count', 0)} report da {len(selected_folders)} cartelle.")
            return render(request, "confronti/archive.html", archive_context(request, selected_archive_folder_ids=selected_ids, selected_archive_folders=selected_folders, report_summary=report_summary, report_summary_warnings=report_summary_warnings, report_summary_scope="folders"))

    return render(request, "confronti/archive.html", archive_context(request))


def _move_message(target_folder):
    if target_folder is None:
        return "Report lasciato senza cartella."
    return f"Report spostato nella cartella {target_folder.customer_name}."


@login_required
def archivio_report_cartella(request, folder_id):
    if not is_archive_admin(request.user):
        return redirect_for_non_archive_admin(request)

    folder = get_object_or_404(CustomerArchiveFolder, pk=folder_id)

    if request.method == "POST":
        action = request.POST.get("action")
        search_query = request.POST.get("q")
        selected_report_id = request.POST.get("selected_report_id")

        if action == "update_archive_folder":
            folder_form = ArchiveFolderForm(request.POST, instance=folder)
            if folder_form.is_valid():
                folder_form.save()
                messages.success(request, "Dati cartella archivio aggiornati.")
                return redirect(archive_redirect_url(search_query, folder_id=folder.pk, report_id=selected_report_id or None))
            return render(request, "confronti/archive.html", archive_context(request, selected_folder=folder, folder_form=folder_form))

        if action == "merge_archive_folder":
            merge_form = ArchiveFolderMergeForm(request.POST, prefix="merge-folder", target_folder=folder)
            if merge_form.is_valid():
                source_folder = merge_form.cleaned_data["source_folder"]
                merged = merge_archive_folders(folder, source_folder)
                messages.success(request, f"Cartella {merged['source_folder_name']} unita in {merged['target_folder_name']} con {merged['moved_report_count']} report spostati.")
                return redirect(archive_redirect_url(search_query, folder_id=folder.pk, report_id=selected_report_id or None))
            return render(request, "confronti/archive.html", archive_context(request, selected_folder=folder, merge_form=merge_form))

        if action == "add_archive_report":
            add_report_form = ArchiveReportUploadForm(request.POST, request.FILES, prefix="add-report")
            if add_report_form.is_valid():
                from .helpers import create_uploaded_archive_report
                create_uploaded_archive_report(folder, add_report_form, request.user)
                messages.success(request, f"File aggiunto nella cartella {folder.folder_name}.")
                return redirect(archive_redirect_url(search_query, folder_id=folder.pk, report_id=selected_report_id or None))
            return render(request, "confronti/archive.html", archive_context(request, selected_folder=folder, add_report_form=add_report_form))

        if action == "build_archive_report_summary":
            selected_report_ids = [int(v) for v in request.POST.getlist("selected_reports") if str(v).isdigit()]
            if not selected_report_ids:
                messages.error(request, "Seleziona almeno un report da includere nel sunto.")
                return render(request, "confronti/archive.html", archive_context(request, selected_folder=folder, selected_report_ids=selected_report_ids))
            selected_reports = list(folder.reports.filter(pk__in=selected_report_ids).order_by("-comparison_datetime", "-created_at"))
            if not selected_reports:
                messages.error(request, "I report selezionati non sono disponibili in questa cartella.")
                return render(request, "confronti/archive.html", archive_context(request, selected_folder=folder))
            report_summary = build_reports_summary_from_archived_reports(selected_reports)
            report_summary_warnings = report_summary.get("warnings", [])
            store_report_summary_session(request, report_summary)
            messages.success(request, f"Sunto creato per {report_summary.get('count', 0)} report selezionati.")
            return render(request, "confronti/archive.html", archive_context(request, selected_folder=folder, selected_report_ids=[r.pk for r in selected_reports], report_summary=report_summary, report_summary_warnings=report_summary_warnings))

        if action == "update_archive_report":
            report = get_object_or_404(ComparisonReport, pk=request.POST.get("report_id"), folder=folder)
            report_form = ArchiveReportForm(request.POST, instance=report, prefix=f"report-{report.pk}")
            if report_form.is_valid():
                report_form.save()
                touch_archive_folder(folder)
                messages.success(request, "Report archiviato aggiornato.")
                return redirect(archive_redirect_url(search_query, folder_id=folder.pk, report_id=report.pk))
            return render(request, "confronti/archive.html", archive_context(request, selected_folder=folder, selected_report=report, active_report_id=report.pk, active_report_form=report_form))

        if action == "replace_archive_report_file":
            report = get_object_or_404(ComparisonReport, pk=request.POST.get("report_id"), folder=folder)
            replace_form = ArchiveReportReplaceFileForm(request.POST, request.FILES, prefix=f"replace-report-{report.pk}")
            if replace_form.is_valid():
                replace_archive_report_file(report, replace_form.cleaned_data["report_file"])
                messages.success(request, "File report sostituito.")
                return redirect(archive_redirect_url(search_query, folder_id=folder.pk, report_id=report.pk))
            return render(request, "confronti/archive.html", archive_context(request, selected_folder=folder, selected_report=report, active_replace_report_id=report.pk, active_replace_report_form=replace_form))

        if action == "move_archive_report":
            report = get_object_or_404(ComparisonReport, pk=request.POST.get("report_id"), folder=folder)
            move_form = ArchiveReportMoveForm(request.POST, prefix=f"move-report-{report.pk}", current_folder=folder)
            if move_form.is_valid():
                target_folder = move_form.cleaned_data["destination_folder"]
                move_archive_report_to_folder(report, target_folder)
                messages.success(request, _move_message(target_folder))
                if target_folder is None:
                    return redirect(archive_redirect_url(search_query, report_id=report.pk))
                return redirect(archive_redirect_url(search_query, folder_id=target_folder.pk, report_id=report.pk))
            return render(request, "confronti/archive.html", archive_context(request, selected_folder=folder, selected_report=report, active_move_report_id=report.pk, active_move_report_form=move_form))

        if action == "delete_archive_report":
            report = get_object_or_404(ComparisonReport, pk=request.POST.get("report_id"), folder=folder)
            delete_archive_report(report)
            messages.success(request, "Report archiviato eliminato.")
            return redirect(archive_redirect_url(search_query, folder_id=folder.pk))

        if action == "delete_archive_folder":
            deleted = delete_archive_folder(folder)
            messages.success(request, f"Cartella {deleted['folder_name']} eliminata. {deleted['report_count']} report restano senza cartella.")
            return redirect(archive_redirect_url(search_query))

    return render(request, "confronti/archive.html", archive_context(request, selected_folder=folder))


@login_required
def apri_cartella_archivio_locale(request, folder_id):
    if not is_archive_admin(request.user):
        return redirect_for_non_archive_admin(request)
    folder = get_object_or_404(CustomerArchiveFolder, pk=folder_id)
    if request.method != "POST":
        return redirect("archivio_report_cartella", folder_id=folder.pk)
    folder_path = archive_folder_absolute_path(folder)
    try:
        open_local_archive_path(folder_path)
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
