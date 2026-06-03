from django.core.management import call_command
from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt


@csrf_exempt
def db_health(request):
    try:
        if request.GET.get("migrate") == "1":
            call_command("migrate", interactive=False, verbosity=0)
            call_command("create_initial_superuser", verbosity=0)

        from django.contrib.auth import get_user_model
        from django.db import connection

        with connection.cursor() as cursor:
            cursor.execute("SELECT 1")
            cursor.fetchone()

        tables = set(connection.introspection.table_names())
        user_count = get_user_model().objects.count() if "auth_user" in tables else None
        return JsonResponse(
            {
                "ok": True,
                "vendor": connection.vendor,
                "auth_user": "auth_user" in tables,
                "django_session": "django_session" in tables,
                "user_count": user_count,
            }
        )
    except Exception as exc:
        return JsonResponse(
            {
                "ok": False,
                "error": type(exc).__name__,
                "message": str(exc)[:1000],
            },
            status=500,
        )
