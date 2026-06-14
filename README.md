# Energia Solidale

Web app Django per confrontare una bolletta luce/gas con le tariffe Illumia o E.ON e generare un file Excel di confronto.

## Avvio locale

```bash
source .venv-django/bin/activate
python manage.py runserver 127.0.0.1:8000
```

Poi apri:

```text
http://127.0.0.1:8000/
```

## Pubblicazione web

La pubblicazione e configurata per Render tramite `render.yaml`.

Render usa:

- `build.sh` per installare le dipendenze e raccogliere gli statici;
- `requirements-django.txt` per le dipendenze Python;
- `gunicorn energia_solidale_django.wsgi:application` per avviare Django;
- PostgreSQL configurato da `DATABASE_URL`.

## File dati necessari

- `manage.py`
- `energia_solidale_django/`
- `confronti/`
- `requirements-django.txt`
- `requirements.txt`
- `render.yaml`
- `build.sh`
- `esempio_confronto_corretto.xlsx`
- `indici_pun_psv_2025_2026.xlsx`
- `tariffe/`
- `estrazioni_tariffe/`

## Accesso

L'app richiede login Django. Gli utenti si gestiscono dal pannello admin:

```text
/admin/
```

Su Render il superutente iniziale viene creato con le variabili ambiente:

- `DJANGO_SUPERUSER_USERNAME`
- `DJANGO_SUPERUSER_EMAIL`
- `DJANGO_SUPERUSER_PASSWORD`
