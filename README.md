# ChatMD Streamlit

ChatMD web versija, skirta paleisti per Streamlit Community Cloud.

## Lokalus paleidimas

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## API raktai

Raktai nelaikomi kode. Streamlit Cloud nustatymuose įrašykite Secrets:

```toml
GOOGLE_API_KEY = "..."
OPENAI_API_KEY = "..."
```

Jeigu naudojate tik Gemini, užtenka `GOOGLE_API_KEY`.
