import logging
from threading import Lock

from django.contrib.auth import login
from django.contrib.auth.views import LoginView
from django.core.management import call_command
from django.shortcuts import redirect, render
from django.urls import reverse
from django.utils.http import url_has_allowed_host_and_scheme

from .forms import ClientRegistrationForm
from .models import InviteCode
from .roles import is_internal_user


logger = logging.getLogger(__name__)
_database_ready = False
_database_ready_lock = Lock()


def ensure_login_database_ready():
    global _database_ready
    if _database_ready:
        return
    with _database_ready_lock:
        if _database_ready:
            return
        call_command("migrate", interactive=False, verbosity=0)
        call_command("create_initial_superuser", verbosity=0)
        _database_ready = True


class ResilientLoginView(LoginView):
    def dispatch(self, request, *args, **kwargs):
        try:
            ensure_login_database_ready()
        except Exception:
            logger.exception("Login database setup failed before rendering the login page.")
        return super().dispatch(request, *args, **kwargs)

    def post(self, request, *args, **kwargs):
        try:
            return super().post(request, *args, **kwargs)
        except Exception:
            logger.exception("Login failed because the database is not ready; running migrations once.")
            ensure_login_database_ready()
            return super().post(request, *args, **kwargs)

    def get_success_url(self):
        next_url = self.get_redirect_url()
        if next_url:
            return next_url
        if is_internal_user(self.request.user):
            return reverse("confronto")
        return reverse("accesso_clienti")


def _safe_next_url(request, fallback_name="accesso_clienti"):
    fallback = reverse(fallback_name)
    candidate = request.POST.get("next") or request.GET.get("next") or fallback
    if url_has_allowed_host_and_scheme(candidate, allowed_hosts={request.get_host()}, require_https=request.is_secure()):
        return candidate
    return fallback


def register(request):
    try:
        ensure_login_database_ready()
    except Exception:
        logger.exception("Registration database setup failed before rendering the signup page.")
    next_url = _safe_next_url(request)
    internal_user = is_internal_user(request.user)
    registration_success = ""
    if request.user.is_authenticated and not internal_user:
        return redirect(next_url)
    if request.method == "POST":
        form = ClientRegistrationForm(request.POST)
        if form.is_valid():
            try:
                user = form.save()
            except ValueError:
                form.add_error("invite_code", "Questo codice invito non e piu disponibile.")
            else:
                if internal_user:
                    registration_success = (
                        f"Account cliente creato per {user.email}. "
                        "La tua sessione admin e rimasta attiva."
                    )
                    form = ClientRegistrationForm()
                    return render(
                        request,
                        "registration/register.html",
                        {
                            "form": form,
                            "next": next_url,
                            "registration_success": registration_success,
                        },
                    )
                login(request, user)
                return redirect(next_url)
    else:
        invite_code = InviteCode.normalize_code(request.GET.get("invite_code"))
        initial = {"invite_code": invite_code} if invite_code else None
        form = ClientRegistrationForm(initial=initial)
    return render(
        request,
        "registration/register.html",
        {
            "form": form,
            "next": next_url,
            "registration_success": registration_success,
        },
    )
