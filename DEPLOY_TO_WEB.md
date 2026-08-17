# Turning this into a Website (GitHub + Streamlit Cloud)

This makes the tool available as a real web page — you visit a link, upload
your Excel, click a button, download a zip. No installing anything, on any
device, including your phone.

Total time: about 10 minutes, and it's free.

---

## Step 1 — Create a GitHub account (skip if you have one)

Go to **github.com** → Sign up. It's free.

## Step 2 — Create a new repository

1. Click the **+** icon top-right → **New repository**
2. Name it something like `leave-document-generator`
3. Set it to **Private** (recommended — this contains your company's real forms and employee data structure) or Public, your choice
4. Click **Create repository**

## Step 3 — Upload the tool's files

1. On your new repository's page, click **Add file → Upload files**
2. Open the `tool` folder on your computer (the one from the zip I gave you)
3. Drag in **everything inside** the `tool` folder — all files and the
   `templates` and `scripts` subfolders — directly into the upload box
   (make sure you're uploading the *contents* of the `tool` folder, not the
   folder itself, so `streamlit_app.py` ends up at the top level of the repo)
4. Scroll down, click **Commit changes**

Your repo should now show files like `streamlit_app.py`, `requirements.txt`,
`packages.txt`, `templates/`, `scripts/` sitting at the top level.

## Step 4 — Deploy on Streamlit Community Cloud

1. Go to **share.streamlit.io**
2. Sign in with your GitHub account (click "Continue with GitHub")
3. Click **Create app** (or **New app**)
4. Choose:
   - **Repository:** the one you just created
   - **Branch:** `main`
   - **Main file path:** `streamlit_app.py`
5. Click **Deploy**

It will take a few minutes the first time (it's installing LibreOffice
in the background via `packages.txt`). Once done, you'll get a live link like:

```
https://your-app-name.streamlit.app
```

That link is your website. Bookmark it, share it with coworkers if useful —
anyone with the link can open it and generate documents.

## Step 5 — Using it

- Open your link
- Click **Browse files**, select your filled employee Excel sheet
- Click **Generate Leave Documents**
- Watch the log, then click **⬇️ Download all documents (zip)**
- Unzip on your computer — same folder structure as before (FMCO / Mehan / Al Falak)

## Updating it later

If I (or you) change the code — for example adding the Witness Undertaking
Form — you just upload the changed file(s) to the same GitHub repo
(**Add file → Upload files** again, or drag the new version over the old
one) and Streamlit Cloud automatically redeploys within a minute or two.
No need to repeat Step 4.

## A few things worth knowing

- **Private data:** your Excel upload is processed in memory on Streamlit's
  server for that session and isn't permanently stored, but if this handles
  sensitive employee data, keeping the GitHub repo **Private** and being
  mindful of what a free public hosting tier means for your company's data
  policy is worth a quick check with whoever handles IT/compliance at FMCO.
- **Free tier limits:** Streamlit Community Cloud is free but apps can go to
  sleep after inactivity — the first visit after a while might take ~30
  seconds to "wake up." That's normal.
- **If deployment fails:** click **Manage app** → **Logs** on Streamlit Cloud
  to see the exact error. The most common cause is a file ending up in the
  wrong folder — double check `streamlit_app.py`, `requirements.txt`, and
  `packages.txt` are all sitting at the top level of the repo, not nested
  inside another folder.
