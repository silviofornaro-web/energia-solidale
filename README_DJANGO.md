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

## Note

I file usati sono gli stessi della versione Streamlit:

- `tariffe/`
- `indici_pun_psv_2025_2026.xlsx`
- `esempio_confronto_corretto.xlsx`

La versione Streamlit resta disponibile e non viene rimossa.
