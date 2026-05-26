# UGC Production Tracker

Internal dashboard for tracking UGC show production across Scripting, ER Review, and Production stages.

## Setup (one time)

1. Create a new GitHub repository
2. Upload all files from this zip maintaining the folder structure
3. Go to Vercel → New Project → Import the GitHub repo → Deploy
4. In Vercel project settings → Deployment Protection → disable everything

## Daily workflow

1. Update `data.xlsx` and upload it to the GitHub repo root
2. GitHub Actions automatically runs (~60 sec)
3. Reads xlsx → updates `index.html` with fresh data → pushes back
4. Vercel redeploys → dashboard live

## Repo structure

```
ugc-tracker/
├── index.html                     ← Dashboard (data baked in, no fetch)
├── data.xlsx                      ← Your Excel tracker
├── scripts/
│   └── update_data.py             ← Reads xlsx, updates index.html
└── .github/
    └── workflows/
        └── update-data.yml        ← Auto-runs on every push to main
```
