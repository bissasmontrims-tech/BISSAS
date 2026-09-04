# PFL Production Dashboard

A Vite + React dashboard for Printed Fabric Label (PFL) factory production and operator performance, built with Tailwind CSS, Recharts, and lucide-react.

## Tech stack

- **React 18** + **Vite 5** — app shell and dev/build tooling
- **Tailwind CSS 3** — styling (utility classes used throughout `pfl-dashboard.jsx`)
- **Recharts** — all charts (line, bar, area)
- **xlsx (SheetJS)** — CSV/Excel data import on the Import Data page
- **lucide-react** — icons

No backend, database, or environment variables are required — the app is fully client-side and ships with embedded sample production data that can be replaced via the in-app Import Data page.

## Project structure

```
pfl-dashboard/
├── index.html
├── package.json
├── vite.config.js
├── tailwind.config.js
├── postcss.config.js
├── .gitignore
├── README.md
├── src/
│   ├── main.jsx
│   ├── pfl-dashboard.jsx
│   └── index.css
└── public/
```

## Run locally

```bash
npm install
npm run dev
```

Then open the URL Vite prints (default `http://localhost:5173`).

## Build for production

```bash
npm run build
```

This outputs a static production build to `dist/`. To sanity-check the build locally:

```bash
npm run preview
```

## Upload to GitHub

```bash
git init
git add .
git commit -m "Initial commit: PFL production dashboard"
git branch -M main
git remote add origin https://github.com/<your-username>/<your-repo>.git
git push -u origin main
```

## Deploy to Vercel

**Option A — Vercel dashboard (no CLI):**
1. Push this repo to GitHub (see above).
2. Go to [vercel.com/new](https://vercel.com/new) and import the repository.
3. Vercel auto-detects the Vite framework preset. Confirm these settings (they're the defaults):
   - **Build Command:** `npm run build`
   - **Output Directory:** `dist`
   - **Install Command:** `npm install`
4. Click **Deploy**.

**Option B — Vercel CLI:**
```bash
npm install -g vercel
vercel login
vercel        # first deploy, follow prompts
vercel --prod # promote to production
```

## Environment variables

None. The dashboard runs entirely in the browser with no external API calls, so there is nothing to configure in Vercel's Environment Variables settings.

## Notes

- The dashboard's calculations, filters, tables, charts, operator drill-down, import/export, and settings logic are unchanged from the original `pfl-dashboard.jsx` — this project only adds the surrounding Vite/Tailwind scaffolding needed to run and deploy it.
- Tailwind is configured to scan `index.html` and everything under `src/`, so the utility classes already used in `pfl-dashboard.jsx` work without further setup.
