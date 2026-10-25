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
   The preview artifact is https://claude.ai/artifact/HjioDupJi4uTvMNWc4ftgu (private; update it by url).
   The older one (Xoi3x9U1fajzoEkksaQwxG) is public and can't be republished from a session.
   The snapshot may predate the small image sizes: if data/site.json has no "sizes", make img/s (480px) and img/m (1200px)
   copies locally and add "sizes": ["s","m"] so the preview matches the live site.
   Wrap preview/index.html with <!DOCTYPE html>, charset and viewport before publishing (make-preview strips them).
5. When the owner approves, commit `site/index.html` to `main` and push. The site updates in a few minutes.

The owner writes in Japanese and prefers to see changes quickly in the preview before they go live.
Content format is described in `content/README.txt` (same structure as the Drive folder).
References are a spreadsheet (Google Sheets, read as .xlsx; old references.txt still works) and texts are Google Docs (read as .docx;
old .txt still works). `scripts/office.mjs` reads both without dependencies. `scripts/make-templates.py <site.json> <dir>` turns a built
site.json into the sheet + docs (used for the move to this format). The sheet's Apps Script menu is `drive/リファレンス表のスクリプト.gs`.
