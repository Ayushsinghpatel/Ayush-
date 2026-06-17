# PP Auto · PO Management Dashboard

Internal Purchase Order management dashboard for PP Auto Innovators Pvt. Ltd.
Built with React + Vite + Tailwind CSS + Recharts.

⚠️ **This contains an app-level password gate, not a real authentication
system.** The password is checked in browser JavaScript, so a technically
determined person could bypass it via dev tools. It is meant to keep casual/
accidental access out — not to be the sole protection for genuinely sensitive
data. If supplier rates, pricing, or other confidential business data will
live here long-term, plan to move to proper backend auth or internal/VPN
hosting.

## 1. Local development

```bash
npm install
cp .env.example .env.local
# edit .env.local and set a real password
npm run dev
```

Visit the printed localhost URL, enter the password you set in `.env.local`.

## 2. Push to GitHub

```bash
git init
git add .
git commit -m "Initial commit: PO management dashboard"
git branch -M main
git remote add origin https://github.com/<your-username>/<repo-name>.git
git push -u origin main
```

Create the empty repo on GitHub first (github.com → New repository), then run
the commands above with that repo's URL.

**Important:** `.env.local` is in `.gitignore` and will NOT be pushed — your
real password stays off GitHub. Only `.env.example` (a placeholder) goes up.

## 3. Deploy on Vercel (recommended, free)

1. Go to vercel.com → sign in with GitHub.
2. Click "Add New Project" → select this repo.
3. Framework preset: Vite (auto-detected).
4. Before clicking Deploy, open "Environment Variables" and add:
   - Key: `VITE_APP_PASSWORD`
   - Value: your chosen password
5. Click Deploy.
6. Every future `git push` to `main` auto-redeploys.

## 3b. Or deploy on Netlify (alternative, free)

1. Go to netlify.com → sign in with GitHub.
2. "Add new site" → "Import an existing project" → pick this repo.
3. Build command: `npm run build` · Publish directory: `dist`
4. Site settings → Environment variables → add `VITE_APP_PASSWORD`.
5. Deploy site.

## 4. Sharing access

Share the deployed URL + the password (out-of-band, e.g. verbally or via a
secure channel — not in the same email as the link) with whoever needs
access. Anyone with the URL can see the login screen; only people with the
password get past it.

## 5. Changing the password later

Update the `VITE_APP_PASSWORD` environment variable in the Vercel/Netlify
dashboard and trigger a redeploy (or just push any commit — it redeploys
automatically). Existing logged-in sessions expire automatically after 12
hours (see `SESSION_TTL_HOURS` in `src/AuthGate.jsx`).

## Notes

- All data in `POManagement.jsx` is currently mock/sample data hardcoded in
  the file. Use the "Import Excel" tab in the app to load real data at
  runtime — it stays in browser memory only and is never sent anywhere
  (no backend exists in this app).
- Refreshing the page or closing the tab clears any uploaded data, since
  there's no persistent storage. If you need data to persist across
  sessions, that's a separate feature to add (e.g. a real backend/database).
