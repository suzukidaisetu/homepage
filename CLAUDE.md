# Working on this site with Claude

This repo is Aga Shunsei's archive site. Content lives in a Google Drive folder; the design lives here.

- `site/index.html` — the whole page (markup, styles, script). Design changes happen here.
- `scripts/build.mjs` — reads the Drive folder (secrets `DRIVE_FOLDER_ID`, `GOOGLE_API_KEY`) or `content/`, writes `dist/`.
- `.github/workflows/deploy.yml` — builds hourly and on push to `main`, publishes to GitHub Pages
  (https://suzukidaisetu.github.io/homepage/) and saves the built site to the `snapshot` branch.

## Design loop (preview first, then publish)
1. `git fetch --depth 1 origin +snapshot:refs/remotes/origin/snapshot` — latest real content.
2. Edit `site/index.html`.
3. `node scripts/make-preview.mjs` — writes `preview/` (current page + real content).
4. Publish `preview/index.html` with the files under `preview/` as an Artifact for the owner to review.
   The existing preview artifact is https://claude.ai/artifact/Xoi3x9U1fajzoEkksaQwxG (update it by url).
5. When the owner approves, commit `site/index.html` to `main` and push. The site updates in a few minutes.

The owner writes in Japanese and prefers to see changes quickly in the preview before they go live.
Content format is described in `content/README.txt` (same structure as the Drive folder).
