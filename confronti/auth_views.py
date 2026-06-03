import logging
from threading import Lock

from django.contrib.auth.views import LoginView
from django.core.management import call_command


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
