# Excel → EJS

A tiny single-page browser helper: upload a spreadsheet, pick the table range,
write an [EJS](https://ejs.co/) template, and render it against the rows. Every
data row becomes an `item`, the first row of the range supplies the (snake_case)
keys, and the result is shown both as a live HTML preview and as raw HTML.

> [!NOTE]
> **The UI is in Georgian (ქართული).** This README is in English, but every
> label, hint, and the built-in EJS instructions inside the app are written in
> Georgian.

> [!WARNING]
> This is **not a real product** and nothing worth depending on. It's a
> vibe-coded toy — built fast, for fun, no guarantees, no roadmap, no support.
> Don't ship it anywhere that matters.

## What it does

- Reads `.xlsx`, `.xls`, and `.csv` entirely in the browser — nothing is
  uploaded anywhere.
- Auto-detects the used range and lets you adjust the start/end cells (e.g.
  `A1` → `X32`).
- Treats the first row of the range as headers, normalized to English
  `snake_case` keys (blank headings fall back to `col_1`, `col_2`, …).
- Exposes a single `items` variable to your EJS template
  (`items = [{ key: value }, …]`).
- Shows the output as a sandboxed HTML **preview** or as raw **HTML**, with a
  copy button.

## Tech

- [React 19](https://react.dev/) + [Vite](https://vite.dev/) + TypeScript
- [SheetJS (`xlsx`)](https://sheetjs.com/) for spreadsheet parsing
- [EJS](https://ejs.co/) for templating (runs client-side)

## Develop

```bash
npm install
npm run dev      # start the dev server
npm run build    # type-check + production build to dist/
npm run lint     # eslint
```

## Deploy

Pushing to `main` builds the app and deploys it to GitHub Pages via the official
GitHub Actions workflow (`.github/workflows/deploy.yml`). One-time setup:
**Settings → Pages → Source: "GitHub Actions"**.
