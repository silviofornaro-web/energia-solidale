from django.core.management.base import BaseCommand, CommandError

from confronti.models import InviteCode


class Command(BaseCommand):
    help = "Crea uno o piu codici invito per l'area clienti."

    def add_arguments(self, parser):
        parser.add_argument("--label", default="", help="Etichetta da associare al codice invito.")
        parser.add_argument("--note", default="", help="Nota interna opzionale.")
        parser.add_argument(
            "--count",
            type=int,
            default=1,
            help="Numero di codici da creare. Default: 1.",
        )
        parser.add_argument(
            "--code",
            default="",
            help="Codice esplicito da usare. Valido solo se --count=1.",
        )

    def handle(self, *args, **options):
        count = int(options["count"] or 1)
        label = str(options["label"] or "").strip()
        note = str(options["note"] or "").strip()
        explicit_code = str(options["code"] or "").strip()

        if count < 1:
            raise CommandError("--count deve essere almeno 1.")
        if explicit_code and count != 1:
            raise CommandError("--code puo essere usato solo con --count=1.")

        created_codes = []
        for index in range(count):
            current_label = label
            if label and count > 1:
                current_label = f"{label} {index + 1}"
            invite = InviteCode.objects.create(
                code=explicit_code if index == 0 else "",
                label=current_label,
                note=note,
            )
            created_codes.append(invite)

        for invite in created_codes:
            suffix = f" | {invite.label}" if invite.label else ""
            self.stdout.write(self.style.SUCCESS(f"{invite.code}{suffix}"))
