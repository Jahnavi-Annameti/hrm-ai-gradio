# HRM AI Business Automation Agent — Gradio Edition

This is a **single-process** rebuild of your original FastAPI + React project.
There is no separate backend server to configure — all the logic (resume
parsing, AI scoring, SQLite storage, Excel export, JD/interview generation,
email sending) runs inside the same Python process as the Gradio UI. This is
why the original "backend not working after upload" problem goes away: there's
nothing external to fail to connect to.

## Features (parity with the original app)

- **Resume Screening** — upload a PDF/DOCX/TXT/image resume, set role, skills,
  requirements, qualifications, experience, and get an AI-style confidence
  score (45-94) plus a Selected / Waiting List / Rejected decision, matched &
  missing skills, detected projects/links, and a draft outcome email.
- **Applicants & Pipeline** — a live dashboard (totals, selection/rejection
  rates, average score) plus a filterable table of every candidate. Manage
  status (Select / Waitlist / Reject / Delete) by candidate ID.
- **Excel Export** — one-click download of Selected / Waiting List / Rejected
  / All candidates as styled `.xlsx` files (same layout as the original).
- **Job Description Generator** — rule-based JD generator (same template/
  sections as the original `/ai/generate` route).
- **Interview Questions Generator** — same template as the original
  `/ai/interview-questions` route.
- **Send Outcome Email** — sends via Gmail SMTP using an **App Password**
  (recommended over a normal password since Gmail blocks basic auth).

## Run locally

```bash
pip install -r requirements.txt
python app.py
```

Then open the printed local URL (default `http://0.0.0.0:7860`).

## Deploying on Hugging Face Spaces (Gradio)

1. Create a new Space → SDK: **Gradio**.
2. Upload `app.py` and `requirements.txt` (this folder) to the Space.
3. Spaces auto-detects `app.py` as the entry point and installs
   `requirements.txt`. No extra configuration needed.
4. The `data/` folder (SQLite DB + generated Excel files) will be created
   automatically on first run. Note: on most free hosting tiers this storage
   is **ephemeral** — it resets on redeploy/restart. For persistent data,
   use a Space with a persistent disk, or swap `DB_PATH`/`DATA_DIR` to point
   at a mounted volume.

## Notes on parity / differences from the original

- The original React app had a fake `localStorage`-based login/signup with no
  real auth (the backend's `/auth/login` always returned `{"ok": True}`).
  Since it added no real security, it has been omitted here — everyone using
  the Space shares the same candidate pipeline. If you need per-user
  isolation or real authentication, that's a separate feature to add.
- The AI scoring/decision logic, Excel formatting, JD template, and
  interview-question template are copied over **verbatim** from your
  original `backend/app/main.py`, so results will match exactly.
- Image resumes (`.jpg/.png`) still require `pytesseract` + the `tesseract`
  binary installed on the host to OCR; without it you'll get a friendly
  fallback message, same as the original.
