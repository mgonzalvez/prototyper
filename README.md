# Martin's Card Prototyper

Martin's Card Prototyper is a static browser editor for quickly building tabletop card prototypes. It runs locally or from GitHub Pages with no build step.

## Current experience

- Canvas-first workspace with a slim tool rail for Add, Assets, Cards, Layers, Project, and Export.
- Click or drag to place components; drag to move, resize with eight handles, and rotate with the rotation handle.
- Smart snapping to the safe area, card center, nearby elements, and the configurable grid.
- Arrow-key nudging, Shift precision nudging, Delete, duplicate, zoom, Fit, undo, and redo.
- Contextual inspector with essential controls first and exact geometry/typography under More controls.
- Ability, Character, Item, and Blank starter templates.
- Multi-card sessions with independent designs, duplication, deletion, and switching.
- Searchable Fractal Symbols library with 327 deduplicated CC BY 4.0 symbols for Icon and Artwork elements.
- PNG/JPG upload support with 2.5 MB and 3000 × 3000 validation.
- IndexedDB draft recovery plus portable version-3 JSON project files.
- PNG preview/export and tiled PDF export for US Letter or A4, including cut guides and Fractal Symbols attribution when needed.
- Shared blue/violet PnP-tools header, light/dark themes, Related Sites menu, responsive desktop/tablet editing, and mobile-friendly access.

## Components

The built-in component types are Background, Title, Rules, Flavor, Artwork, and Icon/value. New components receive opinionated visual defaults and can be refined through style presets or the contextual inspector.

Library symbols are embedded as inline SVG data URLs when selected, so saved projects remain portable. Fractal Symbols are licensed under CC BY 4.0; the app provides license details and adds a discreet attribution line to PDF sheets that contain library artwork.

## Quick start

1. Start the app and choose Ability, Character, Item, or Blank.
2. Open Add, choose a component, then click or drag on the card.
3. Select Artwork or Icon/value to open Assets and choose a Fractal symbol or upload PNG/JPG artwork.
4. Select an element to use presets and essential styling. Open More controls for exact dimensions, rotation, font, frame, and transparency.
5. Use Preview, then export PNG or a print-ready PDF from Export.
6. Use Cards to create additional cards in the same session.
7. Save Project JSON when you need a portable copy. Browser drafts are recovered automatically after refresh.

## Project files and browser storage

New downloads use the version-3 multi-card shape:

```json
{
  "version": 3,
  "format": "martins-card-prototyper-project",
  "activeCardId": "...",
  "cards": []
}
```

The loader accepts the prior version-2 multi-card format and legacy single-design files. New optional tile fields include physical `fontSizePt`, `sourceAsset` metadata, and icon value/tint settings. Images and selected library SVGs are embedded inline.

IndexedDB stores one debounced recoverable draft. JSON Save/Load remains the portable persistence mechanism. Local storage stores only theme and tutorial preferences.

## Fractal Symbols manifest

The bundled assets are in `Fractal-Symbols-Complete-Pack/` plus selected legacy-only symbols in `FractalSymbols Icons Set/`. Regenerate the browser manifest after changing those folders:

```bash
node scripts/build-icon-manifest.mjs
```

The generated `icon-manifest.js` is intentionally checked in for static hosting.

## Run locally

```bash
python3 -m http.server 8080
```

Open <http://localhost:8080>. Google Fonts, pdf-lib, and html2canvas load from CDNs, so browser use normally needs an internet connection.

## Verification

```bash
node --check app.js
node --check editor-core.js
node --test tests/*.test.js
```

The tests cover template defaults, legacy migration, geometry constraints, history, icon-manifest integrity, DOM IDs, and local resource references. Manual browser checks should cover placement, transforms, text editing, library selection, draft restore, JSON round trips, PNG/PDF output, themes, and the tutorial.

## Deployment

The repository is configured for GitHub Pages with custom domain `prototyper.gonzhome.us` in `CNAME`.
