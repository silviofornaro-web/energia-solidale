#!/usr/bin/env python
import os
import sys


def main():
    os.environ.setdefault("DJANGO_SETTINGS_MODULE", "energia_solidale_django.settings")
    if (
        os.environ.get("RENDER_EXTERNAL_HOSTNAME")
        and len(sys.argv) > 1
        and sys.argv[1] in {"migrate", "create_initial_superuser"}
        and os.environ.get("RUN_RENDER_STARTUP_MAINTENANCE", "").lower() not in {"1", "true", "yes", "on"}
    ):
        print(f"Skipping {sys.argv[1]} during Render startup.")
        return
    from django.core.management import execute_from_command_line

    execute_from_command_line(sys.argv)


if __name__ == "__main__":
    main()
