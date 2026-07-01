# MailFlow — Hosting Guide

## ⚠️ Read this first — why your Vercel site can't send email

The app has **two parts**:

| Part | What it is | Can it live on Vercel? |
|------|-----------|------------------------|
| **Frontend** (`static/`) | The dashboard UI — HTML/CSS/JS | ✅ Yes (it's just static files) |
| **Backend** (`app.py`) | The engine that actually sends email over SMTP | ❌ **No** |

**A Vercel deployment shows the dashboard but cannot send a single email**, because:
1. Vercel is *serverless* — functions stop after 10–60s and can't hold a live SMTP session or stream progress for a minutes-long campaign.
2. Serverless/PaaS free tiers commonly **block outbound SMTP ports** (25/465/587) to stop spam.

So the backend must run somewhere **always-on that allows outbound SMTP**. Your options, honestly ranked:

- **Local machine** (best/simplest) — `python app.py`, use `http://localhost:8000`. Sends immediately, zero cost. ← use this to actually send.
- **A small VPS** (Hostinger VPS, DigitalOcean, Fly.io) — for a public URL that can send. You control the ports.
- **Render/Railway free** — runs the app, **but free tiers block SMTP**, so sending will fail. Only useful as a public demo of the UI.

👉 **To send email right now: run it locally.** The Vercel URL is only a UI showcase.

---

## Option A — Everything local (simplest)

```bash
pip install -r requirements.txt
python app.py
```
Opens `http://localhost:8000` automatically. Leave `static/config.js` empty.

---

## Option B — Frontend on Vercel + Backend on Render (free, public)

### 1. Push this folder to GitHub.

### 2. Deploy the backend to Render
1. Go to [render.com](https://render.com) → **New +** → **Blueprint** → pick your repo.
2. It reads `render.yaml` and deploys `app.py`. You get a URL like
   `https://mailflow-api.onrender.com`.
3. (Optional, after step 3) set the `ALLOWED_ORIGINS` env var to your Vercel URL.

### 3. Deploy the frontend to Vercel
1. Edit **`static/config.js`** → set your Render URL, then commit & push:
   ```js
   window.MAILFLOW_API_BASE = "https://mailflow-api.onrender.com";
   ```
2. Go to [vercel.com](https://vercel.com) → **Add New** → **Project** → import
   `github.com/ankit-datatrainer/MailFlow`.
3. In the import screen set:
   - **Framework Preset:** `Other`
   - **Root Directory:** `static`   ← important, this makes Vercel serve the UI folder
   - Build Command / Output Directory: leave blank
4. Click **Deploy**. Your dashboard goes live at `https://<project>.vercel.app`.
5. Back on Render, set `ALLOWED_ORIGINS` to that Vercel URL so the API accepts it.

---

## Notes
- **Free SMTP limits:** Gmail ~500/day; keep the per-email delay at 5s+ to avoid throttling.
- **Render free tier sleeps** after 15 min idle — the first request wakes it (~30s cold start). Fine for occasional campaigns.
- **Templates** live in `templates/<name>/` (`subject.txt` + `body.html`) and are read by the backend, so add them in the repo the backend is deployed from.
