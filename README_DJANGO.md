# Energia Solidale - Versione Django

Questa versione affianca l'app Streamlit e usa Django per una web app più stabile:

- login utenti tramite Django
- form bolletta senza rerun continui
- confronto a video in tabella HTML
- download Excel come normale risposta HTTP

## Avvio locale

```bash
cd /Users/silviofornaro/Desktop/EnergiaSolidaleLast
python3 -m venv .venv-django
source .venv-django/bin/activate
pip install -r requirements-django.txt
python manage.py migrate
python manage.py createsuperuser
python manage.py runserver
```

Poi apri:

```text
http://127.0.0.1:8000/
```

## Utenti

Gli utenti si gestiscono dall'admin Django:

```text
http://127.0.0.1:8000/admin/
```

Puoi creare utenti con username uguale alla loro email.

## Pubblicazione su Render

Il repository contiene `render.yaml`, quindi su Render puoi creare un Blueprint collegato al repository GitHub.

Durante la creazione ti verranno chieste queste variabili:

- `DJANGO_SUPERUSER_USERNAME`: username amministratore iniziale
- `DJANGO_SUPERUSER_EMAIL`: email amministratore iniziale
- `DJANGO_SUPERUSER_PASSWORD`: password amministratore iniziale

Render creerà anche un database PostgreSQL e avvierà l'app con:

```text
gunicorn energia_solidale_django.wsgi:application
```

Dopo il primo deploy puoi entrare in:

```text
https://indirizzo-render/admin/
```

Da lì puoi creare gli utenti autorizzati all'uso dell'app.

Per aggiungere un dominio personalizzato come `app.energiasolidale.org`, aggiungi il dominio su Render e poi imposta nel DNS del dominio il record richiesto da Render.

## Note

I file usati sono gli stessi della versione Streamlit:

- `tariffe/`
- `indici_pun_psv_2025_2026.xlsx`
- `esempio_confronto_corretto.xlsx`

La versione Streamlit resta disponibile e non viene rimossa.
