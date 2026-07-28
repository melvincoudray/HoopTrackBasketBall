# Suivi Indiv — Deploy to Netlify + Supabase

## What's in this folder

- `src/App.jsx` — your full app, now wired to Supabase for permanent storage
- `src/main.jsx` — small file that starts the app
- `package.json`, `vite.config.js`, `index.html` — standard project setup
- `netlify.toml` — tells Netlify how to build the site
- `supabase-schema.sql` — the one database table this app needs
- `.env.example` — template for your Supabase keys

## Steps

### 1. Supabase (the database)

1. Create a free account at [supabase.com](https://supabase.com) → "New Project".
2. Once created, go to **SQL Editor** → paste the contents of `supabase-schema.sql` → Run.
3. Go to **Project Settings → API** and note down two values: **Project URL** and **anon public key**.

### 2. GitHub

1. Create a repository (e.g. `suivi-indiv`) at github.com.
2. Upload this folder's contents, keeping the `src` folder structure (don't upload `.env` if you created one locally — only `.env.example`).

### 3. Netlify (hosting)

1. Create a free account at netlify.com.
2. "Add new site" → "Import an existing project" → connect GitHub → pick your repository.
3. Build command `npm run build` and publish directory `dist` should be detected automatically.
4. Before deploying, go to **Site configuration → Environment variables** and add:
   - `VITE_SUPABASE_URL` = the Project URL from step 1
   - `VITE_SUPABASE_ANON_KEY` = the anon public key from step 1
5. Click **Deploy**.

### 4. Test it

Open the live URL. Try creating a team, adding a player, importing a match —
then refresh the page. If your data is still there after refreshing, Supabase
is correctly connected and your data is now permanently saved in the cloud.

## If the build fails

Copy the error message from Netlify's "Deploy log" and send it back — it's
almost always a small, fixable detail (missing dependency, typo in a path,
or a missing environment variable).
