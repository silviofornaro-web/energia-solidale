import os

from django.contrib.auth import get_user_model
from django.contrib.auth.models import Group
from django.core.management.base import BaseCommand

from confronti.roles import ILLUMIA_OPERATOR_GROUP


class Command(BaseCommand):
    help = "Create the initial admin and Illumia operator users from environment variables if they do not exist."

    def handle(self, *args, **options):
        User = get_user_model()
        self._create_superuser(User)
        self._create_illumia_operator(User)

    def _create_superuser(self, User):
        username = os.environ.get("DJANGO_SUPERUSER_USERNAME", "").strip()
        email = os.environ.get("DJANGO_SUPERUSER_EMAIL", "").strip()
        password = os.environ.get("DJANGO_SUPERUSER_PASSWORD", "")

        if not username or not email or not password:
            self.stdout.write("Initial superuser skipped: missing environment variables.")
            return

        user, created = User.objects.get_or_create(
            username=username,
            defaults={"email": email, "is_staff": True, "is_superuser": True},
        )
        if created:
            user.set_password(password)
            user.save(update_fields=["password"])
            self.stdout.write(self.style.SUCCESS(f"Initial superuser created: {username}"))
            return

        changed = False
        if not user.is_staff or not user.is_superuser:
            user.is_staff = True
            user.is_superuser = True
            changed = True
        if email and user.email != email:
            user.email = email
            changed = True
        if changed:
            user.save(update_fields=["email", "is_staff", "is_superuser"])
        self.stdout.write(f"Initial superuser already exists: {username}")

    def _create_illumia_operator(self, User):
        username = os.environ.get("DJANGO_OPERATOR_USERNAME", "").strip().lower()
        email = os.environ.get("DJANGO_OPERATOR_EMAIL", username).strip().lower()
        password = os.environ.get("DJANGO_OPERATOR_PASSWORD", "")
        first_name = os.environ.get("DJANGO_OPERATOR_FIRST_NAME", "").strip()
        last_name = os.environ.get("DJANGO_OPERATOR_LAST_NAME", "").strip()

        if not username or not password:
            self.stdout.write("Initial Illumia operator skipped: missing environment variables.")
            return

        group, _created_group = Group.objects.get_or_create(name=ILLUMIA_OPERATOR_GROUP)
        user, created = User.objects.get_or_create(
            username=username,
            defaults={
                "email": email or username,
                "first_name": first_name,
                "last_name": last_name,
                "is_staff": True,
                "is_superuser": False,
            },
        )
        if created:
            user.set_password(password)
            user.save(update_fields=["password"])
            self.stdout.write(self.style.SUCCESS(f"Initial Illumia operator created: {username}"))
        else:
            changed_fields = []
            if user.email != (email or username):
                user.email = email or username
                changed_fields.append("email")
            if first_name and user.first_name != first_name:
                user.first_name = first_name
                changed_fields.append("first_name")
            if last_name and user.last_name != last_name:
                user.last_name = last_name
                changed_fields.append("last_name")
            if not user.is_staff:
                user.is_staff = True
                changed_fields.append("is_staff")
            if user.is_superuser:
                user.is_superuser = False
                changed_fields.append("is_superuser")
            if not user.check_password(password):
                user.set_password(password)
                changed_fields.append("password")
            if changed_fields:
                user.save(update_fields=changed_fields)
            self.stdout.write(f"Initial Illumia operator already exists: {username}")

        user.groups.add(group)
