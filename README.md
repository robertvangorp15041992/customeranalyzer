# customeranalyzer

Klanten omzetanalyse dashboard (Streamlit).

## Lokaal draaien

```bash
pip install -r requirements.txt
export DASHBOARD_PASSWORD="jouw_wachtwoord"
streamlit run dashboard.py
```

## Deploy op Railway

1. Deploy vanuit GitHub
2. Voeg in **Variables** de omgevingsvariabele toe:
   - `DASHBOARD_PASSWORD` = jouw wachtwoord