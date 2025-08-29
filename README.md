PPT Generator

Simple web app that turns a large block of text into a PowerPoint presentation using a provided template.

Quick start

1. Create virtualenv and install:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

2. Run locally:

```bash
python app.py
```

3. Open http://127.0.0.1:5000

Notes

- This POC does not call an LLM yet — it splits text into slide-sized chunks and maps them to slides using the uploaded template when provided.
- Do not store user API keys in logs or files. The current app accepts an API key in the form but does not persist it.
