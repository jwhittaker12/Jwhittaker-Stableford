# ⛳ Frends Stableford Championship App

A web app to track your year-long Stableford competition.
Log rounds, view live standings, and export to Excel — no more manual updates.

---

## Setup & Deployment (step by step)

### What you need
- A free [GitHub](https://github.com) account
- A free [Vercel](https://vercel.com) account (sign up with GitHub)

---

### Step 1 — Create a GitHub repository
1. Go to [github.com/new](https://github.com/new)
2. Name it `stableford-app` (or anything you like)
3. Set it to **Private**
4. Click **Create repository**

---

### Step 2 — Upload the files
On the new repo page, click **uploading an existing file**, then drag and drop all these files:
- `app.py`
- `requirements.txt`
- `vercel.json`
- `.gitignore`
- The `templates/` folder (containing `index.html`)

Click **Commit changes**.

---

### Step 3 — Deploy to Vercel
1. Go to [vercel.com](https://vercel.com) and sign in with GitHub
2. Click **Add New → Project**
3. Find your `stableford-app` repo and click **Import**
4. Leave all settings as default and click **Deploy**
5. Wait ~60 seconds — Vercel will give you a live URL like `https://stableford-app.vercel.app`

---

### Step 4 — Open your app
Visit the URL Vercel gave you. The app will:
- Load all existing round data automatically
- Show the live leaderboard
- Let you log new rounds from any device

---

## Using the app

### Logging a round
1. Tap **➕ Log Round** in the nav
2. Enter the date, select the course, and choose 9 or 18 holes
3. Enter each player's Stableford points and gross score (leave blank for players who didn't play)
4. Hit **Submit Round** — standings update instantly

### Adding a new course
When logging a round, select **+ Add new course...** from the dropdown.
Enter the course name and par (per 9 holes). It'll be saved for future rounds.

### Updating a handicap
Go to **⚙️ Settings** and update the player's handicap. Hit Save.

### Exporting to Excel
Click the **⬇ Export Excel** button at the top right to download a formatted spreadsheet.

---

## Important note on data storage

Vercel's free tier uses **serverless functions** which means the SQLite database
resets on each deployment. To keep your data permanently, you have two options:

**Option A (easiest):** Use [PlanetScale](https://planetscale.com) or [Supabase](https://supabase.com)
as a free hosted database. Ask Claude to help you switch the DB connection.

**Option B:** Export to Excel regularly as your backup — all data is in the export.

---

## Questions?
Just ask Claude — the full build script (`build_golf_v5.py`) contains all the round
data and can regenerate your Excel file at any time.
