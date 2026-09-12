# Catalog Compliance Checker

A browser-only TMF620 product catalog auditor. Upload a JSON/CSV/XLSX
catalog export and get a severity-ranked compliance report — no data
leaves the browser. No sample data, either: nothing is sent to a server,
so there's nothing to try if you don't have a real file handy — use the
built-in "Try the sample catalog" link on the upload screen, which loads
a deliberately imperfect demo catalog exercising every rule below.

## Tech stack

- **React 18 + Vite 7** — no backend, builds to static assets
- **Tailwind CSS** for styling
- **PapaParse** (CSV) and **xlsx (SheetJS)** for spreadsheet parsing
- **jsPDF + jspdf-autotable** for the PDF export
- **lucide-react** for icons

All parsing, validation, and report generation happens client-side in
[src/App.jsx](src/App.jsx).

## How scoring works

Each entity contributes a fixed number of checks; every non-info issue
found (`critical` or `warning`) counts against the total. The compliance
score is `100 - (issues / total checks) * 100 * 4`, clamped to `0–100`.
`info`-level findings (unused specs/prices/categories, duplicate names,
uncategorized offerings) are surfaced in the report but don't affect the
score.

## What it checks

**Structural**
- Missing ID / name / lifecycle status
- Duplicate IDs across entities
- Unrecognized lifecycle status values (against the standard TMF620 set)

**Referential integrity**
- Broken spec / price / category / bundle-child references
- Active offering built on a Retired specification
- Category hierarchy cycles (A → parent B → parent A)
- Bundle cycles (a bundle that (indirectly) contains itself)

**Date / lifecycle sanity**
- Invalid date formats in validFrom / validTo
- validFrom after validTo
- Marked Active but validTo has already passed
- Marked Active but validFrom is still in the future

**Hygiene (informational, doesn't affect the score)**
- Uncategorized offerings
- Unused specs / prices / categories (defined but never referenced)
- Duplicate offering names

## Export

From the report view you can download:
- **CSV** — flat list of every issue (severity, entity, rule, message)
- **PDF** — formatted report with score summary and issue table

## Local development

```bash
npm install
npm run dev
```

Opens at `http://localhost:5173`.

## Build

```bash
npm run build
```

Output goes to `dist/`. Preview it locally with:

```bash
npm run preview
```

## Deployment

The build output is fully static (no server/API needed), so any static
host works.

### Vercel
```bash
npm i -g vercel
vercel
```
Vercel auto-detects Vite. Accept the defaults (build command `npm run
build`, output directory `dist`).

### Netlify
```bash
npm i -g netlify-cli
netlify deploy --build
```
Or connect the repo in the Netlify dashboard with build command `npm run
build` and publish directory `dist`.

### GitHub Pages
1. In `vite.config.js`, set `base: "/your-repo-name/"` (already stubbed
   with a comment in this project).
2. Build: `npm run build`
3. Push the contents of `dist/` to a `gh-pages` branch, e.g. using the
   `gh-pages` npm package:
   ```bash
   npm i -D gh-pages
   npx gh-pages -d dist
   ```
4. Enable Pages in the repo settings, pointing at the `gh-pages` branch.

### Any other static host (S3, Cloudflare Pages, Render, etc.)
Just upload the contents of `dist/` after running `npm run build`.

## File format reference

**JSON** — array or object of entities. Accepts:
- TMF620-nested shape (each offering has `productSpecification`,
  `category`, `productOfferingPrice` as nested objects/arrays, and
  `validFor.startDateTime` / `endDateTime`)
- Flat shape with explicit `specId` / `priceId` / `categoryId` /
  `childOfferingIds` fields
- Singular or plural bucket key names (`productOffering` or
  `productOfferings`, `category` or `categories`, etc.)
- A single bare entity object (no wrapper) is treated as one offering

**CSV / Excel** — columns: `id`, `type`, `name`, `lifecycleStatus`,
`validFrom`, `validTo`, `specId`, `priceId`, `categoryId`,
`childOfferingIds`.

## Known limitations

This is a client-side structural/referential checker. It does not
cross-check against a live source system (Hansen EPC, Amdocs, BRM,
etc.) — that would require a backend connector with credentials to
those systems.
