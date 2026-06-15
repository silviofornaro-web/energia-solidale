import logging
import os
from urllib.parse import quote, urljoin

from django.contrib.auth import get_user_model
from django.contrib.auth.decorators import login_required
from django.http import HttpResponse
from django.shortcuts import redirect, render
from django.urls import reverse
from django.utils import timezone

from .bill_parser import parse_uploaded_bill
from .forms import BillUploadForm, ConfrontoForm, CustomerInviteForm, session_to_service_data
from .models import InviteCode
from .services import build_excel_bytes, offer_options_payload, prepare_comparison, provider_label, safe_download_filename


logger = logging.getLogger(__name__)
LAST_UPLOADED_BILL_NAME_KEY = "last_uploaded_bill_name"
LAST_UPLOADED_BILL_NAME_CLIENT_KEY = "last_uploaded_bill_name_cliente_illumia"
LAST_COMPARISON_KEY = "last_confronto"
LAST_COMPARISON_CLIENT_KEY = "last_confronto_cliente_illumia"
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


def _mode_config(customer_mode=False):
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
    return {
        "comparison_session_key": LAST_COMPARISON_KEY,
        "upload_name_session_key": LAST_UPLOADED_BILL_NAME_KEY,
        "page_title": "Confronto bollette",
        "page_intro": "",
        "download_url_name": "scarica_excel",
    }


def _force_customer_mode_data(data):
    payload = dict(data)
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


def _comparison_form(*args, customer_mode=False, **kwargs):
    return ConfrontoForm(*args, customer_mode=customer_mode, **kwargs)


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


def _is_internal_user(user):
    return bool(user.is_authenticated and (user.is_staff or user.is_superuser))


def _confronto_page(request, *, customer_mode=False):
    mode = _mode_config(customer_mode)
    prepared = None
    rows = None
    extraction_warnings = []
    extraction_count = 0
    customer_invite_form = CustomerInviteForm()
    customer_invite_result = None
    customer_status = None
    uploaded_bill_name = request.session.get(mode["upload_name_session_key"], "")
    upload_form = BillUploadForm()
    if request.method == "POST":
        action = request.POST.get("action")
        if action == "send_customer_invite" and not customer_mode:
            customer_invite_form = CustomerInviteForm(request.POST)
            form = _comparison_form(customer_mode=customer_mode)
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
        elif action == "show_customer_status" and not customer_mode:
            customer_status = _customer_status_snapshot()
            form = _comparison_form(customer_mode=customer_mode)
        elif action == "extract_bill":
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
                    form = _comparison_form(customer_mode=customer_mode)
                    extraction_warnings = [
                        "Non sono riuscito a leggere questo PDF. Inserisci manualmente i valori della bolletta."
                    ]
                else:
                    initial_data = parsed.values
                    if customer_mode:
                        initial_data = _force_customer_mode_data(initial_data)
                    form = _comparison_form(initial=initial_data, customer_mode=customer_mode)
                    extraction_warnings = parsed.warnings
                    extraction_count = len(parsed.values)
                    upload_form = BillUploadForm()
            else:
                form = _comparison_form(customer_mode=customer_mode)
        elif action == "reset_bill":
            request.session.pop(mode["comparison_session_key"], None)
            request.session.pop(mode["upload_name_session_key"], None)
            uploaded_bill_name = ""
            providers = ["ILLUMIA"] if customer_mode else (request.POST.getlist("providers") or [request.POST.get("provider") or ""])
            initial_data = {
                "nome_cliente": request.POST.get("nome_cliente") or "",
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
            form = _comparison_form(initial=initial_data, customer_mode=customer_mode)
        else:
            form = _comparison_form(request.POST, customer_mode=customer_mode)
            if form.is_valid():
                data = form.service_data()
                if customer_mode:
                    data = _force_customer_mode_data(data)
                data["comparison_datetime"] = timezone.localtime().isoformat(timespec="seconds")
                prepared = prepare_comparison(data)
                rows = prepared["rows"]
                session_data = form.session_data()
                if customer_mode:
                    session_data = _force_customer_mode_data(session_data)
                session_data["comparison_datetime"] = data["comparison_datetime"]
                request.session[mode["comparison_session_key"]] = session_data
                form = _comparison_form(initial=data, customer_mode=customer_mode)
    else:
        initial = (
            {
                "providers": ["ILLUMIA"],
                "tariff_selection_mode": "LATEST",
                "pod_pdr": "",
                "email_cliente": "",
                "telefono_cliente": "",
            }
            if customer_mode
            else None
        )
        form = _comparison_form(initial=initial, customer_mode=customer_mode)

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
            "customer_optional_open": _customer_optional_section_open(form),
            "page_title": mode["page_title"],
            "page_intro": mode["page_intro"],
            "download_url_name": mode["download_url_name"],
            "customer_dashboard_local_url": reverse("confronto_cliente_illumia"),
            "customer_register_url": _customer_registration_url(),
            "customer_invite_form": customer_invite_form,
            "customer_invite_result": customer_invite_result,
            "customer_status": customer_status,
            "whatsapp_sender_number": WHATSAPP_SENDER_NUMBER,
            "whatsapp_web_home_url": WHATSAPP_WEB_HOME_URL,
        },
    )


def _scarica_excel(request, *, customer_mode=False):
    mode = _mode_config(customer_mode)
    raw = request.session.get(mode["comparison_session_key"])
    if not raw:
        return HttpResponse("Nessun confronto pronto. Torna alla pagina principale e calcola il confronto.", status=400)
    data = session_to_service_data(raw)
    if customer_mode:
        data = _force_customer_mode_data(data)
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


def confronto(request):
    if not request.user.is_authenticated:
        return redirect("register")
    if not _is_internal_user(request.user):
        return redirect("confronto_cliente_illumia")
    return _confronto_page(request)


@login_required
def confronto_cliente_illumia(request):
    return _confronto_page(request, customer_mode=True)


@login_required
def scarica_excel(request):
    if not _is_internal_user(request.user):
        return redirect("confronto_cliente_illumia")
    return _scarica_excel(request)


@login_required
def scarica_excel_cliente_illumia(request):
    return _scarica_excel(request, customer_mode=True)
