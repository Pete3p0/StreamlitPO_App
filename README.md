# Streamlit PO PDF Converter

Upload multiple purchase-order PDFs and download one combined Excel sheet.

## Files
- `app.py` - Streamlit app
- `requirements.txt` - Python deps
- `Procfile` - Optional process file
- `setup.sh` - Optional Streamlit config bootstrap
- `.streamlit/config.toml` - Streamlit settings

## Run locally
```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\\Scripts\\activate
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Cloud
1. Push this folder to GitHub repo.
2. In Streamlit Cloud, create app and point to `app.py`.
3. It will install `requirements.txt` automatically.

> Note: Streamlit Cloud does not require `Procfile`, but included if you need Heroku-style deploys too.
