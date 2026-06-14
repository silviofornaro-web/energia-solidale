from django.contrib.auth import get_user_model
from django.core.management.base import BaseCommand

from confronti.models import InviteCode


class Command(BaseCommand):
    help = "Mostra clienti registrati e stato dei codici invito."

    def handle(self, *args, **options):
        User = get_user_model()
        users = list(User.objects.filter(is_staff=False, is_superuser=False).order_by("-id"))
        available = list(
            InviteCode.objects.filter(is_active=True, used_at__isnull=True, used_by__isnull=True).order_by("-created_at")
        )
        used = list(
            InviteCode.objects.exclude(used_at__isnull=True, used_by__isnull=True).order_by("-used_at", "-created_at")
        )

        self.stdout.write("=== CLIENTI REGISTRATI ===")
        for user in users:
            full_name = f"{user.first_name} {user.last_name}".strip()
            self.stdout.write(
                f"{user.id} | {user.username} | {user.email} | {full_name or 'N.D.'} | attivo={user.is_active}"
            )
        self.stdout.write(f"Totale clienti registrati: {len(users)}")

        self.stdout.write("")
        self.stdout.write("=== CODICI DISPONIBILI ===")
        for invite in available:
            self.stdout.write(f"{invite.code} | {invite.label} | attivo={invite.is_active}")
        self.stdout.write(f"Totale codici disponibili: {len(available)}")

        self.stdout.write("")
        self.stdout.write("=== CODICI USATI ===")
        for invite in used:
            used_by = invite.used_by.username if invite.used_by else "N.D."
            self.stdout.write(f"{invite.code} | {invite.label} | usato_da={used_by} | used_at={invite.used_at}")
        self.stdout.write(f"Totale codici usati: {len(used)}")
