import os

from django.core.wsgi import get_wsgi_application


os.environ.setdefault("DJANGO_SETTINGS_MODULE", "energia_solidale_django.settings")


def run_render_startup_maintenance():
    if not os.environ.get("RENDER_EXTERNAL_HOSTNAME"):
        return
    if os.environ.get("SKIP_RENDER_STARTUP_MAINTENANCE", "").lower() in {"1", "true", "yes", "on"}:
        return

    from django.core.management import call_command

    call_command("migrate", interactive=False, verbosity=0)
    call_command("create_initial_superuser", verbosity=0)


application = get_wsgi_application()
run_render_startup_maintenance()
