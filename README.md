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

## Setup Instructions

### Step 1 — Create a GitHub repo

1. Go to [github.com/new](https://github.com/new)
2. Name it `report-automation` (private repo recommended)
3. Don't initialize with README (you already have files)

### Step 2 — Push the files

Open Terminal in the `report-automation` folder and run:

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/report-automation.git
git push -u origin main
```

### Step 3 — Deploy on Streamlit Cloud (free)

1. Go to [share.streamlit.io](https://share.streamlit.io)
2. Sign in with your GitHub account
3. Click **"New app"**
4. Select your `report-automation` repo, branch `main`, file `app.py`
5. **Before clicking Deploy** → click **"Advanced settings"**
6. In the **Secrets** box, paste this (with your REAL JSON key):

```toml
GOOGLE_SERVICE_ACCOUNT_JSON = '< paste entire contents of your report-automation-455720-f26e907a0c3e.json file here as a single line >'
```

> **How to get the single-line JSON:** Open your `.json` key file, select all, copy.
> Then in the Secrets box, type `GOOGLE_SERVICE_ACCOUNT_JSON = '` then paste, then close with `'`
> Make sure the entire JSON is on ONE line between the single quotes.

7. Click **Deploy**

Your app will be live at `https://YOUR_APP_NAME.streamlit.app` in ~2 minutes.

### Step 4 — Share with your colleague

Just send them the URL. They open it in any browser, upload CSV, click RUN. Done.

## Local Development (optional)

```bash
# Create virtual environment
python3 -m venv venv
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Set up secrets for local run
cp secrets.toml.example .streamlit/secrets.toml
# Edit .streamlit/secrets.toml with your real JSON key

# Run
streamlit run app.py
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
