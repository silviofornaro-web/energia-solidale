# Energia Solidale

Web app Streamlit per confrontare una bolletta luce/gas con le tariffe Illumia o E-ON e generare un file Excel di confronto.

## Avvio locale

```bash
python -m streamlit run app_rest.py
```

Su Mac puoi anche usare:

```bash
./avvia_energia_solidale.command
```

## Pubblicazione come web app

Per Streamlit Community Cloud:

1. Crea un repository GitHub.
2. Carica questi file/cartelle:
   - `streamlit_app.py`
   - `app_rest.py`
   - `requirements.txt`
   - `esempio_confronto_corretto.xlsx`
   - `indici_pun_psv_2025_2026.xlsx`
   - `tariffe/`
   - `estrazioni_tariffe/`
   - `.streamlit/config.toml`
3. Su Streamlit Cloud seleziona il repository.
4. Come main file usa `streamlit_app.py`.

## File dati

- `esempio_confronto_corretto.xlsx`: template per l'export del confronto.
- `indici_pun_psv_2025_2026.xlsx`: indici PUN/PSV.
- `tariffe/`: tariffe Illumia per segmento e mese.
- `estrazioni_tariffe/`: tariffe E-ON estratte dai PDF.

## Accesso con utenti autorizzati

Il login interno si abilita dai Secrets di Streamlit Cloud. Le password non vanno mai inserite in chiaro: genera prima un hash con:

```bash
python generate_password_hash.py
```

Poi in Streamlit Cloud, nella sezione Secrets, inserisci:

```toml
[auth]
enabled = true

[auth.users]
"utente@example.com" = "HASH_GENERATO"

[auth.names]
"utente@example.com" = "Nome Cognome"
```

Se `auth.enabled` manca o vale `false`, l'app resta aperta senza login.
