# Daily Data Analyst Paid Search Job Search

Runs at * AM Toronto EDT every day via GitHub Actions.
Emails you an Excel report with 170+ Data Engineer roles from the last 24 hours.

100% free. No AI. No resume. Just jobs.

## Setup (10 minutes)

### 1 — Create GitHub repo and push these files to main branch

### 2 — Get Adzuna API keys (free)
- Go to developer.adzuna.com → Register → Dashboard → Create App
- Copy your App ID and App Key

### 3 — Get Gmail App Password
- myaccount.google.com/security → enable 2-Step Verification
- Search "App passwords" → Create → copy the 16-char password

### 4 — Add 4 GitHub Secrets
Repo → Settings → Secrets and variables → Actions

| Secret             | Value                  |
|--------------------|------------------------|
| ADZUNA_APP_ID      | Your Adzuna App ID     |
| ADZUNA_APP_KEY     | Your Adzuna App Key    |
| GMAIL_ADDRESS      | yourname@gmail.com     |
| GMAIL_APP_PASSWORD | 16-char app password   |

### 5 — Personalise job_search.py (top of file)
- SKILLS — your skillset
- SEARCH_QUERIES — your target job titles

### 6 — Run
Actions → Daily Data Engineer Job Search → Run workflow

## Cost
Everything is free. No paid APIs or subscriptions.

## Excel Output
- All roles sorted by skill match score
- Green rows = strongest matches
- Clickable Apply → links
- Columns: Title, Company, Location, Remote, Salary, Posted, Skills Match, Score, Qualifications
