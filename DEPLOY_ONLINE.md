# 🚀 Deploy Online — Get a Shareable Link (Free, ~10 mins)

## What you need
- Free GitHub account → https://github.com
- Free Streamlit Cloud account → https://share.streamlit.io

---

## STEP 1 — Create a GitHub account
Go to https://github.com → Sign up (any email)

---

## STEP 2 — Create a new repository
1. Click **+** (top right) → **New repository**
2. Name it: `whatsapp-roi-dashboard`
3. Set to **Public** → click **Create repository**

---

## STEP 3 — Upload your 2 files
1. In your new repo click **"uploading an existing file"**
2. Drag and drop: `app.py` and `requirements.txt`
3. Click **Commit changes**

---

## STEP 4 — Deploy on Streamlit Cloud
1. Go to https://share.streamlit.io → sign in with GitHub
2. Click **New app**
3. Select repo: `whatsapp-roi-dashboard`
4. Main file: `app.py`
5. Click **Deploy!**

Wait ~2 mins → you get a live link like:
```
https://yourname-whatsapp-roi-dashboard.streamlit.app
```
Share this link with your team. No install needed for them.

---

## STEP 5 — Add email secrets (one time)
In Streamlit Cloud → your app → **Settings → Secrets**, paste:
```toml
EMAIL_SENDER   = "asifpowerking786@gmail.com"
EMAIL_PASSWORD = "your-16-char-app-password"
```
Click **Save** → app restarts automatically.

> ⚠️ Your Gmail is used to SEND emails, but recipients will see
> "WhatsApp ROI Dashboard <noreply@classplus.co>" as the sender name.
> To truly send from noreply@classplus.co, your IT team needs to
> configure that domain's SMTP credentials instead.

---

## STEP 6 — Done!
Share the `.streamlit.app` link with teammates.
No login, no install — just open and use. ✅
