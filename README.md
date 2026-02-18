# 📊 Report Automation — Streamlit Web App

Automates daily CRM report generation: reads a CSV export, filters and categorizes invoices, uploads to Google Sheets with formatting, and shares the result.

## Project Structure

```
report-automation/
├── app.py                    ← Main app (this is the only code file)
├── requirements.txt          ← Python dependencies
├── secrets.toml.example      ← Template for your Google credentials
├── .streamlit/
│   └── config.toml           ← Streamlit theme + settings
└── .gitignore                ← Keeps secrets out of git
```

## What changed from the desktop (.app) version

| Aspect | Old (tkinter) | New (Streamlit) |
|---|---|---|
| Platform | macOS .app only | Any browser, any device |
| Auth | JSON file bundled in app | Secrets stored securely in Streamlit Cloud |
| UI freezing | Main thread blocked during run | Streamlit handles this natively |
| Error handling | Crash → messagebox | Inline error messages in the UI |
| Sharing | Build .exe / .app per platform | Send a URL |
| Google auth | Runs on every app launch | Cached with `@st.cache_resource` |
