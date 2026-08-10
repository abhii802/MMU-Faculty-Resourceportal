# Faculty Resource Console — Generator Backend

Wraps `research_paper_generator.py`, `cs_pptx_generator.py`, and
`cs_notes_generator.py` behind three private HTTP endpoints so the
GitHub Pages front end can call them. The `.py` files are never sent
to the browser — only the generated `.docx` / `.pptx` comes back.

## Files
- `app.py` — Flask server, three routes
- `research_paper_generator.py`, `cs_pptx_generator.py`, `cs_notes_generator.py` — your scripts, each with one small callable wrapper added (`generate_paper`, `generate_ppt`) so they can be called with a topic string instead of `input()`
- `requirements.txt`

## Deploy (Render.com — free tier, easiest)
1. Push this `backend/` folder to a **private** GitHub repo (keep the scripts private — that repo does not need to be public).
2. On [render.com](https://render.com): **New → Web Service** → connect that repo.
3. Build command: `pip install -r requirements.txt`
4. Start command: `gunicorn app:app`
5. Environment variables (Render dashboard → Environment):
   - `ALLOWED_ORIGIN` = `https://YOUR-GITHUB-USERNAME.github.io`
   - `OPENAI_API_KEY` = *(optional — enables AI-enhanced content; without it, scripts fall back to template mode)*
6. Deploy. You'll get a URL like `https://your-service.onrender.com`.

Render's free tier sleeps after inactivity — the first request after a break takes ~30–50s to wake up. Fine for occasional faculty use; upgrade to a paid instance if that delay is a problem.

Railway, Fly.io, or PythonAnywhere work the same way — same three files, same environment variables.

## Test it once deployed
```bash
curl -X POST https://your-service.onrender.com/api/generate/notes \
  -H "Content-Type: application/json" \
  -d '{"topic":"Binary Search Trees"}' \
  --output test.docx
```

## Connect the webpage
In `faculty-resource-portal.html`, near the bottom `<script>`, set:
```js
const GENERATOR_ENDPOINTS = {
  paper: "https://your-service.onrender.com/api/generate/paper",
  ppt:   "https://your-service.onrender.com/api/generate/ppt",
  notes: "https://your-service.onrender.com/api/generate/notes"
};
```
Commit and push — GitHub Pages updates automatically. The three generator buttons go live with no other change.

## Notes
- Paper generation is the slowest (fetches references, builds figures/tables) — expect 1–3 minutes on a free-tier instance. Consider a longer client-side timeout or a "check back in a minute" message if you see request timeouts.
- Without `OPENAI_API_KEY`, the PPT and notes generators run in template mode (still produces a file, just without AI-written content).
