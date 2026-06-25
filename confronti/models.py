import secrets
import string
from pathlib import Path

from django.conf import settings
from django.db import models
from django.utils import timezone
from django.utils.text import slugify


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


def _archive_storage_filename(filename):
    cleaned = Path(str(filename or "report.xlsx")).name
    normalized = "".join(ch if ch.isalnum() or ch in {"-", "_", "."} else "_" for ch in cleaned)
    return normalized or "report.xlsx"


def archive_report_upload_to(instance, filename):
    return f"report_archive/{instance.folder.folder_name}/{_archive_storage_filename(filename)}"


class CustomerArchiveFolder(models.Model):
    customer_name = models.CharField("Nome cliente", max_length=160)
    customer_email = models.EmailField("Email cliente", blank=True)
    customer_phone = models.CharField("Telefono cliente", max_length=40, blank=True)
    folder_name = models.CharField("Cartella archivio", max_length=180, unique=True, editable=False)
    notes = models.TextField("Note", blank=True)
    created_at = models.DateTimeField("Creata il", auto_now_add=True)
    updated_at = models.DateTimeField("Aggiornata il", auto_now=True)
    created_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="archive_folders_created",
        verbose_name="Creata da",
    )

    class Meta:
        ordering = ("-updated_at", "-created_at")
        verbose_name = "Cartella archivio cliente"
        verbose_name_plural = "Cartelle archivio clienti"

    def __str__(self):
        return f"{self.customer_name} ({self.folder_name})"

    @staticmethod
    def normalize_customer_name(value):
        return " ".join(str(value or "").strip().split())

    @staticmethod
    def normalize_customer_email(value):
        return str(value or "").strip().lower()

    @staticmethod
    def normalize_customer_phone(value):
        return "".join(ch for ch in str(value or "") if ch.isdigit())

    @classmethod
    def build_folder_name_candidate(cls, customer_name, *, current_date=None):
        date_label = (current_date or timezone.localdate()).isoformat()
        slug = slugify(cls.normalize_customer_name(customer_name)) or "cliente"
        return f"{date_label}__{slug[:120]}"

    @classmethod
    def generate_unique_folder_name(cls, customer_name):
        base_name = cls.build_folder_name_candidate(customer_name)
        candidate = base_name
        suffix = 2
        while cls.objects.filter(folder_name=candidate).exists():
            candidate = f"{base_name}-{suffix}"
            suffix += 1
        return candidate

    def save(self, *args, **kwargs):
        self.customer_name = self.normalize_customer_name(self.customer_name) or "Cliente senza nome"
        self.customer_email = self.normalize_customer_email(self.customer_email)
        self.customer_phone = self.normalize_customer_phone(self.customer_phone)
        if not self.folder_name:
            self.folder_name = type(self).generate_unique_folder_name(self.customer_name)
        super().save(*args, **kwargs)


class ComparisonReport(models.Model):
    folder = models.ForeignKey(
        CustomerArchiveFolder,
        on_delete=models.CASCADE,
        related_name="reports",
        verbose_name="Cartella cliente",
    )
    title = models.CharField("Titolo report", max_length=180, blank=True)
    commodity = models.CharField("Fornitura", max_length=12, blank=True)
    providers_label = models.CharField("Fornitori confronto", max_length=120, blank=True)
    bill_period_label = models.CharField("Periodo bolletta", max_length=120, blank=True)
    comparison_datetime = models.DateTimeField("Confronto eseguito il", null=True, blank=True)
    report_file = models.FileField("File report", upload_to=archive_report_upload_to, max_length=255)
    original_filename = models.CharField("Nome file originale", max_length=255, blank=True)
    notes = models.TextField("Note", blank=True)
    comparison_data = models.JSONField("Dati confronto", default=dict, blank=True)
    created_at = models.DateTimeField("Creato il", auto_now_add=True)
    updated_at = models.DateTimeField("Aggiornato il", auto_now=True)
    created_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="comparison_reports_created",
        verbose_name="Creato da",
    )

    class Meta:
        ordering = ("-comparison_datetime", "-created_at")
        verbose_name = "Report confronto archiviato"
        verbose_name_plural = "Report confronti archiviati"

    def __str__(self):
        return self.title or self.original_filename or Path(self.report_file.name).name
