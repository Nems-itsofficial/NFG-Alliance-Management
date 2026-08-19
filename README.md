# Alliance Command — standalone web app

This is the free-hosted, Claude-independent version of the alliance tool. It's a normal
React app (Vite) talking to a Supabase database. Nothing here depends on Anthropic or
Claude in any way — once it's deployed, it runs on its own forever, for $0/month on the
free tiers of Supabase + Vercel.

**Important:** I can't create accounts or click "deploy" on your behalf — those steps
need to happen in your own browser, logged into your own accounts. Everything below is
copy-paste / point-and-click, no coding required. Budget about 20 minutes the first time.

---

## 1. Create the database (Supabase — free)

1. Go to **supabase.com** → sign up → **New project**. Pick any name/region, set a database password (save it somewhere).
2. Once the project is ready, go to **SQL Editor** → **New query**.
3. Open `schema.sql` from this folder, copy all of it, paste into the editor, click **Run**.
   This creates all five tables and locks them down so only signed-in officers can read/write.
4. Go to **Settings → API**. Copy the **Project URL** and the **anon public** key — you'll need both in step 3.
5. Go to **Authentication → Users → Add user**. Create one account per officer (email + password), with **Auto Confirm User** checked. There's no public sign-up page — accounts only exist if you create them here.

## 2. Put the code on GitHub

1. Go to **github.com** → sign in (or sign up, free) → **New repository** → name it anything (e.g. `alliance-command`) → Create.
2. On your computer, unzip the project folder you downloaded from this chat, then either:
   - Use GitHub Desktop / your usual git workflow to push it, **or**
   - On the new repo's page, use **"uploading an existing file"** and drag in every file/folder from the unzipped project (keep the `src/` folder structure intact).

## 3. Deploy (Vercel — free)

1. Go to **vercel.com** → sign up with your GitHub account.
2. **Add New → Project** → select the repo you just created → Vercel auto-detects Vite. Don't click Deploy yet.
3. Expand **Environment Variables** and add:
   - `VITE_SUPABASE_URL` → the Project URL from step 1.4
   - `VITE_SUPABASE_ANON_KEY` → the anon public key from step 1.4
4. Click **Deploy**. In ~1 minute you'll get a live URL like `alliance-command.vercel.app` — that's your app, shareable with officers.
5. Every officer signs in with the email/password you created for them in step 1.5.

That's it — no domain purchase required. (You can attach a custom domain later in Vercel's project settings if you ever want one, but the free `.vercel.app` URL works indefinitely.)

## Local development (optional)

If you want to run it on your own computer first:

```
npm install
cp .env.example .env      # then fill in your Supabase URL + anon key
npm run dev
```

## Migrating data from the Claude version

In the Claude artifact version, use **Settings → Export backup (.json)** to download a
complete snapshot. In this app, sign in, open **Settings → Import backup**, and upload
that same file. It clears whatever's currently in the database and re-loads everything —
roster, growth profiles, events, and participation — with fresh IDs, fully wired up.

## Notes

- **Free tier limits:** Supabase free tier includes 500MB database and pauses a project
  after a week of no API activity (an officer visiting the site wakes it back up in a
  few seconds). Vercel's free tier has no realistic bandwidth ceiling for a tool this size.
- **This is your export-safe backup, not a live mirror of the Hostinger version** (or
  vice versa, whichever you make primary) — see the chat for why I'd recommend against
  running both live at once.
