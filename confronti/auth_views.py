import logging

from django.contrib.auth.views import LoginView
from django.core.management import call_command


logger = logging.getLogger(__name__)


class ResilientLoginView(LoginView):
    def post(self, request, *args, **kwargs):
        try:
            return super().post(request, *args, **kwargs)
        except Exception:
            logger.exception("Login failed because the database is not ready; running migrations once.")
            call_command("migrate", interactive=False, verbosity=0)
            call_command("create_initial_superuser", verbosity=0)
            return super().post(request, *args, **kwargs)
