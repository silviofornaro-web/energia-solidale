import secrets
import string

from django.conf import settings
from django.db import models
from django.utils import timezone


class InviteCode(models.Model):
    code = models.CharField("Codice invito", max_length=24, unique=True, blank=True)
    label = models.CharField("Etichetta", max_length=120, blank=True)
    note = models.TextField("Note", blank=True)
    is_active = models.BooleanField("Attivo", default=True)
    created_at = models.DateTimeField("Creato il", auto_now_add=True)
    used_at = models.DateTimeField("Usato il", null=True, blank=True)
    used_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="invite_codes_used",
        verbose_name="Usato da",
    )

    class Meta:
        ordering = ("-created_at",)
        verbose_name = "Codice invito"
        verbose_name_plural = "Codici invito"

    def __str__(self):
        return self.code

    @staticmethod
    def normalize_code(value):
        cleaned = "".join(ch for ch in str(value or "").upper() if ch.isalnum())
        return cleaned

    @classmethod
    def generate_code(cls, length=10):
        alphabet = string.ascii_uppercase + string.digits
        return "".join(secrets.choice(alphabet) for _ in range(length))

    @property
    def is_available(self):
        return self.is_active and not self.used_at and self.used_by_id is None

    def save(self, *args, **kwargs):
        self.code = self.normalize_code(self.code)
        if not self.code:
            candidate = self.generate_code()
            while type(self).objects.filter(code=candidate).exists():
                candidate = self.generate_code()
            self.code = candidate
        super().save(*args, **kwargs)

    def mark_used(self, user):
        self.used_by = user
        self.used_at = timezone.now()
        self.is_active = False
        self.save(update_fields=["used_by", "used_at", "is_active"])
