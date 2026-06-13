from django.conf import settings
from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    initial = True

    dependencies = [
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
    ]

    operations = [
        migrations.CreateModel(
            name="InviteCode",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("code", models.CharField(blank=True, max_length=24, unique=True, verbose_name="Codice invito")),
                ("label", models.CharField(blank=True, max_length=120, verbose_name="Etichetta")),
                ("note", models.TextField(blank=True, verbose_name="Note")),
                ("is_active", models.BooleanField(default=True, verbose_name="Attivo")),
                ("created_at", models.DateTimeField(auto_now_add=True, verbose_name="Creato il")),
                ("used_at", models.DateTimeField(blank=True, null=True, verbose_name="Usato il")),
                (
                    "used_by",
                    models.ForeignKey(
                        blank=True,
                        null=True,
                        on_delete=django.db.models.deletion.SET_NULL,
                        related_name="invite_codes_used",
                        to=settings.AUTH_USER_MODEL,
                        verbose_name="Usato da",
                    ),
                ),
            ],
            options={
                "verbose_name": "Codice invito",
                "verbose_name_plural": "Codici invito",
                "ordering": ("-created_at",),
            },
        ),
    ]
