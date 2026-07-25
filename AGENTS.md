# Martin's Card Prototyper — Agent Instructions

## Architecture

Single-page static web app. Three files, no build step:

- `index.html` — shell, all UI markup, CDN script imports (pdf-lib, html2canvas)
- `app.js` — all logic. State lives in a global `state` object. `render()` clears and rebuilds the card stage, selection inspector, layer list, and card list.
- `styles.css` — shared PnP-tools design tokens plus editor styles; light/dark state is mirrored on `html[data-theme]` and `body[data-theme]`

No package manager, build step, automated tests, or CI. Runtime dependencies and Google Fonts load from CDNs, so local use still needs network access unless those resources are already cached.

## Running

```bash
python3 -m http.server 8080
```

Open `http://localhost:8080`. No dev server, no hot reload.

## Key patterns in app.js

- **State**: single `state` object with `cards[]`, `activeCardId`, `design`, `selectedTileId`. All cards in a session are independent.
- **Active card**: `state.design` is the working copy of the active card. `syncActiveCardFromState()` copies it back into the matching `cards[]` record; `activateCard()` saves the old working copy before loading another.
- **Naming**: despite the `#projectName` label, `state.design.name` names the active card and supplies its export filename. The version 2 project JSON has no separate project-level name.
- **Render**: `render()` is the single source of truth — it clears `#cardStage`, re-renders visible tiles and supporting panels, syncs the active card record, and updates `updatedAt`. Call it after state mutations that affect rendered UI.
- **Tile types**: `title`, `effect`, `flavor`, `main-image`, `card-background`, `icon`. Each has a preset size/position from `createTile()` and default style from `getDefaultTextStyle()`.
- **Layer order**: `design.tiles` renders from back to front. The Layers panel displays the same array in reverse, so its first item is visually on top.
- **Units**: internal state uses inches. Display uses pixels via `toPixels()`/`toInches()` which derive scale from `#cardStage.clientWidth`.
- **Drag/resize**: `beginDrag()` → `onPointerMove()` → `onPointerUp()`. Snap uses `state.design.gridStepIn`.
- **Inspector**: selecting a tile opens the floating palette. Title and icon tiles expose a basic/advanced toggle; other tile types show their applicable controls directly.
- **Application frame**: the header and design tokens intentionally match BoardSplitter, Card Formatter, and Card Extractor. Preserve the shared blue/violet palette, translucent surfaces, branded header, Related Sites menu, and responsive behavior when changing UI.
- **Theme boot**: the inline script in `index.html` selects the saved or system theme before CSS loads to prevent a flash. `applyTheme()` mirrors the value to both the root element and body, updates the browser theme-color meta tag, and leaves the SVG toggle icons intact.
- **Workspace layout**: desktop uses a left controls panel, center stage, and fixed right inspector. At `1180px` and below the layout becomes one column and the inspector docks along the viewport bottom.
- **Canvas sizing**: `.stage-panel` must keep an explicit `grid-template-columns: minmax(0, 1fr)` so the grid does not collapse to max-content. `#cardStageWrap` uses a definite responsive width (`500–640px` on desktop, with smaller viewport-based limits at the responsive breakpoints). Avoid percentage-only sizing on this grid item; it previously collapsed the card to an illegible width.
- **Export**: PNG via html2canvas (screenshot of `#cardStage`). PDF via pdf-lib (tiled print sheets with corner cut guides). All-card PNG queues one browser download per card; all-card PDF places each card once and ignores `#pdfCopies`.
- **Project files**: `buildProjectPayload()` writes version 2 multi-card JSON. `loadDesignFromData()` also accepts the legacy single-design shape. Uploaded images are stored inline as data URLs.
- **Persistence**: projects are not auto-saved or restored. `persistDraft()` currently only syncs in-memory state, and `loadDraft()` always returns `false`. Durable saves require JSON download/upload.
- **Local storage**: only theme (`martins_card_prototyper.theme`) and tutorial dismissal (`martins_card_prototyper.tutorial_seen`) are persisted.
- **Reset**: `Start Over` creates one starter card while preserving the active card size, snap, grid-step, and safe-border settings.

## Gotchas

- **No `npm`/`yarn`** — editing `index.html` to add scripts means managing CDN versions manually.
- **CSS cache busting**: `index.html` loads `styles.css` with a version query. Increment it when a deployed CSS change must invalidate browser/CDN caches.
- **`#cardStage` is the export canvas** — anything with `.export-hide` class is hidden during PNG/PDF export.
- **Safe border**: `SAFE_BORDER_IN = 0.125` inches. Card-background tiles are exempt from safe border constraints.
- **Image upload limits**: 2.5 MB max, 3000x3000 max dimensions (constants at top of `app.js`).
- **Image formats**: validation accepts PNG and JPEG MIME types only. The same validator is used for tile images and title-banner images.
- **Fonts**: loaded via Google Fonts `<link>` in `index.html` head. Adding new fonts requires editing that link.
- **Cloudflare analytics** token is embedded in `index.html` — do not expose or rotate without updating.
- **Manual verification**: use `node --check app.js` for syntax, then exercise affected behavior through the local HTTP server in a browser.

## File structure

```
index.html    — markup + CDN imports
app.js        — all application logic
styles.css    — all styles
test card/    — sample assets (not part of the app)
CNAME        — GitHub Pages custom domain
```
