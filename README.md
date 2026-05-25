# Apple · Strategic Account · Cursor

Reveal.js presentation for the Apple/Cursor account thesis deck.

## View online

After GitHub Pages is enabled: **https://davidreinfeld.github.io/cursor-panel/**

## Local

Open `index.html` in a browser, or serve the folder:

```bash
python3 -m http.server 8080
```

Then visit http://localhost:8080

## Deploy

Push `main` to GitHub. The [Pages workflow](.github/workflows/pages.yml) publishes the site automatically.

```bash
git push origin main
```

First-time setup in the repo: **Settings → Pages → Build and deployment → Source: GitHub Actions**.

## PPTX (optional)

```bash
npm install
node add_momentum_slide.js
```

Generates `momentum_slide_test.pptx` (gitignored).
