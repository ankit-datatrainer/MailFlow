# MailFlow — Hosting Guide

The app has **two parts**:

| Part | What it is | Where it can be hosted free |
|------|-----------|-----------------------------|
| **Frontend** (`static/`) | The dashboard UI — pure HTML/CSS/JS | ✅ **Vercel** (static, instant) |
| **Backend** (`app.py`) | The sending engine — runs campaigns for minutes, streams live progress, sends SMTP | ✅ **Render** / Railway free tier |

> ⚠️ **Why not the backend on Vercel?** Vercel is *serverless* — functions stop after 10–60 seconds and can't run background threads or live SSE streams. A campaign with delays runs for minutes and must hold state, so it needs an **always-on** host. Render's free tier does this. Vercel is perfect for the static frontend.

You can also just run everything **locally** (`python app.py`) — no hosting needed.

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
