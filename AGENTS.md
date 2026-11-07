# Martin's Card Prototyper — Agent Instructions

## Architecture

Static browser app with no package manager or build step:

- `index.html` — shell, editor markup, CDN imports, dialogs, and tutorial layer.
- `app.js` — browser state, rendering, placement/transforms, inspector, assets, persistence, tutorial, and exports.
- `editor-core.js` — pure templates, migration, geometry, snapping helpers, and history shared by browser code and Node tests.
- `styles.css` — shared PnP-tools visual system and editor/canvas styles.
- `icon-manifest.js` — generated 327-entry Fractal Symbols metadata file.
- `scripts/build-icon-manifest.mjs` — regenerates the icon manifest from the bundled asset folders.
- `tests/editor-core.test.js` — dependency-free Node checks.

Runtime dependencies and Google Fonts load from CDNs. Local browser use normally needs network access for those resources.

## Running and verification

```bash
python3 -m http.server 8080
node --check app.js
node --check editor-core.js
node --test tests/*.test.js
```

Open `http://localhost:8080`. There is no hot reload.

## State model

- `state.cards[]` stores independent card records.
- `state.design` is the working copy of the active card.
- `state.activeCardId` identifies the active record.
- `state.selectedTileId` identifies the selected element.
- `state.history` stores bounded command snapshots for undo/redo.
- `syncActive()` copies the working design back into its card record.
- Mutations should go through `mutate()` so rendering, history, and draft scheduling stay consistent.

The editor uses inches for design geometry. Screen positions are converted through the rendered card width. New typography uses physical `fontSizePt`; legacy pixel font sizes are migrated by `editor-core.js`.

## Interaction invariants

- Add is a click-or-drag placement workflow. Escape cancels placement.
- Single-click selects/moves. Double-click enters text editing. Escape or blur commits text.
- Pointer transforms update the current DOM element during the gesture and commit one history command on pointer-up.
- Resize handles support Shift aspect-ratio locking. Rotation supports free degrees, Shift 15° snapping, and strong cardinal snapping.
- Smart guides use the printable safe border, card center, and nearby visible elements.
- Selection overlays, handles, guides, and safe-area helpers must never appear in PNG/PDF output.
- Background elements fill the card and stay at the back of the layer stack.

## Templates and components

`PrototyperCore.makeTemplate()` defines Ability, Character, Item, and Blank layouts. Components are Background, Title, Rules, Flavor, Artwork, and Icon/value. Keep new defaults attractive at the physical card scale and prefer contextual presets over exposing every setting immediately.

## Assets and licensing

Fractal Symbols are bundled under `Fractal-Symbols-Complete-Pack/` with selected legacy-only assets under `FractalSymbols Icons Set/`. Regenerate `icon-manifest.js` with:

```bash
node scripts/build-icon-manifest.mjs
```

Selected SVGs are embedded as data URLs and carry `sourceAsset` metadata. Fractal Symbols are CC BY 4.0. Keep the in-app license dialog, suggested attribution, footer attribution for relevant PDF sheets, and source links intact.

Uploaded images remain PNG/JPG only, max 2.5 MB and 3000 × 3000 pixels.

## Persistence and project compatibility

New JSON is version 3 with the existing multi-card structure. Loaders must continue accepting version-2 multi-card projects and legacy single-design files. Do not serialize undo history.

IndexedDB stores one debounced recoverable draft. JSON Save/Load is still the portable persistence path. Local storage is reserved for theme and tutorial preference keys.

## Export

PNG uses html2canvas on `#card-stage`. PDF uses pdf-lib with US Letter/A4, 3×3, 3×2, and Gutterfold layouts plus corner cut guides. All-card PNG exports one file per card; all-card PDF places each card once. PDF sheets containing Fractal artwork include the required attribution line.

## UI system

Preserve the shared PnP-tools blue/violet palette, translucent surfaces, branded sticky header, Related Sites menu, light/dark theme boot, responsive behavior, focus styles, and reduced-motion support. Increment the `styles.css?v=...` query in `index.html` when deployed CSS changes need cache invalidation.

## Manual checks

Exercise placement, move, resize, rotation, snapping, keyboard nudging, double-click text editing, style presets, library search/category selection, PNG/JPG upload, template creation, card switching, undo/redo, IndexedDB restore/discard, JSON migration/round trip, PNG/PDF export, attribution, theme switching, responsive layouts, and tutorial navigation.

## Registry & provenance rule (portfolio-wide)

Personal project — completely separate from Apple work.

- `.npmrc` pins `registry=https://registry.npmjs.org`. Never let Apple's internal npm mirror (`npm.apple.com`, `artifacts.apple.com/artifactory/...`) end up in `package-lock.json`; GitHub Actions can't reach it and CI fails with `ENOTFOUND`.
- Audit before pushing: `../scripts/check-registry.sh` (add `--fix` to normalize).
- No work credentials, internal hostnames, proprietary assets, or confidential material in this repo.
