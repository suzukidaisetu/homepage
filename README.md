# Aga Shunsei — archive

Google Drive のフォルダに素材を入れると、1時間以内にサイトが自動で更新される仕組みです。

- `site/` サイトの見た目と動き(デザインを変えるときはここ)
- `scripts/build.mjs` Drive のフォルダを読んで、サイトを組み立てるプログラム
- `content/` 見本の素材。Drive がまだつながっていないときは、これでサイトが作られます
- `.github/workflows/deploy.yml` 毎時間、自動で組み立てて公開する設定

素材の入れ方は `content/README.txt` を見てください。

## 手元で確認する
```
npm install
npm run build      # content/ から dist/ を作る
npm run preview    # http://localhost:8080 で表示
```
Drive から作るときは `DRIVE_FOLDER_ID` と `GOOGLE_API_KEY` を環境変数に入れて `npm run build`。
