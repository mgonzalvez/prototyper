# Prototyper — Development Guide

Martin's Card Prototyper is a **single-page static web app** for designing tabletop card mockups and exporting print-ready sheets. No build step, no framework — vanilla HTML/CSS/JS with two CDN libraries.

## File Structure

```
prototyper/
├── index.html   (~520 lines)  — App shell, all UI panels, modals, and CDN script imports
├── app.js       (~2526 lines) — All application logic: state, rendering, drag/drop, export
├── styles.css   (~933 lines) — Design tokens, layout, component styles, dark mode
├── README.md    — User-facing documentation and quick-start guide
└── CNAME        — Custom domain for GitHub Pages
```

## Architecture Overview

### Three-file, zero-dependency architecture

The entire app lives in three files. There is no build system, no bundler, no framework. Two external libraries are loaded via CDN:

- **pdf-lib** (unpkg) — PDF document creation and PNG embedding for print-ready exports
- **html2canvas** (jsDelivr) — DOM-to-canvas rasterization for PNG preview/export

Everything else is vanilla JavaScript with direct DOM manipulation.

### State model

All mutable app state lives in a single `state` object:

```js
state = {
  projectId: null,          // JSON project file ID (null = fresh session)
  design: {                 // Current card design (tiles, card size, settings)
    name: "Card 1",
    card: { preset: "poker", widthIn: 2.5, heightIn: 3.5 },
    snapEnabled: true,
    gridStepIn: 0.005,
    showSafeBorder: true,
    tiles: [...],           // Array of tile objects (see below)
    updatedAt: "ISO string",
  },
  cards: [                  // Multi-card session: each card has its own design
    { id, name, design, selectedTileId },
  ],
  activeCardId: "...",      // Currently active card ID
  selectedTileId: "...",    // Currently selected tile ID (null = nothing selected)
  liveEditable: null,       // Currently live-editing tile ID (null = none)
  suspendDraftSync: false,  // Pause localStorage draft auto-save
}
```

### Tile types

Six tile types, each with distinct rendering and control panels:

| Type | Key properties | Controls |
|------|---------------|----------|
| `title` | `text`, `style`, `titleBanner` | Font, size, alignment, banner mode (color/gradient/image), banner padding |
| `effect` | `text`, `style` | Font, size, alignment, line height, letter spacing, color |
| `flavor` | `text`, `style` | Same as effect but defaults to Cormorant Garamond |
| `main-image` | `imageFill` (mode: image/solid/gradient, imageDataUrl) | Image upload, fill mode, gradient presets |
| `card-background` | `imageFill` | Same as main-image, background layer |
| `icon` | `imageFill`, `iconValue` (value, color, size, outline, nudgeX, nudgeY) | Numeric value overlay on icon image |

Each tile object has: `id`, `type`, `name`, `x`, `y`, `width`, `height`, `rotationDeg`, `hidden`, `showOutline`, `transparentBg`, `style`, `imageFill`, `titleBanner`, `iconValue`, `text` (for text tiles).

### Rendering pipeline

1. **`render()`** — Main render function called on every state change
2. Clears and rebuilds `#cardStage` from `state.design.tiles`
3. Each tile becomes a DOM element with position, rotation, and content
4. Text tiles use `contenteditable` for live editing
5. Image tiles render uploaded images via `<img>` or CSS background
6. Layer panel (`#layersList`) rebuilds from `state.design.tiles`

### Drag-and-drop system

Tiles on `#cardStage` support:
- **Drag to move** — mousedown on non-text tile body, track mouse movement, update `x`/`y` in state
- **Resize** — drag the resize handle (bottom-right corner) to adjust `width`/`height`
- **Snap-to-grid** — when enabled, positions snap to `gridStepIn` increments
- **Safe border** — 0.125" border guardrails visible during editing (toggleable)
- **Rotation** — 0°, 90°, 180°, 270° via dropdown (applies CSS transform)

### Multi-card session

- Cards are stored in `state.cards[]`, each with independent design
- `state.activeCardId` tracks which card is currently being edited
- Switching cards swaps `state.design` from the card's design clone
- Duplicate creates a new card with a copy of the current card's design
- Delete removes a card from the array (minimum 1 card always)

### Export system

**PNG export:**
- Uses `html2canvas` to rasterize `#cardStage`
- Scales up by 2-4x for 300 DPI output
- Hides UI chrome (safe border, selection handles, delete buttons) via `.export-mode` class
- Active card: single PNG download
- All cards: separate PNG file per card, triggered sequentially

**PDF export:**
- Uses `pdf-lib` to create print-ready sheets
- Layouts: 3x3 grid, 2x3 grid, Gutterfold (rotated for booklet)
- Pages: US Letter or A4
- Copies: configurable (1-300) for single-card mode
- All-cards mode: each card placed once across pages
- Corner cut guides drawn as lines above each card position
- Cards are rotated clockwise for Gutterfold layout

**Preview:**
- Renders the current card via `html2canvas` into a modal
- Shows a scaled preview before export
- Download button saves the preview PNG

### Project file workflow

- **Save** — Serializes full `state` to JSON, triggers download
- **Load** — Reads JSON file, normalizes data (handles legacy single-card format), rebuilds state
- **Start Over** — Resets to fresh project, optionally preserving card size/grid settings
- **Draft auto-save** — Saves current state to `localStorage` on every change, restores on page load

### UI panels

**Left panel** (collapsible sections):
- **Project** — Name, card size preset, custom dimensions, snap/grid/safe border toggles, grid step
- **Cards** — Create, duplicate, delete, switch cards; card count display
- **Tiles** — Tile palette to add new tiles (drag or click to add)
- **Layers** — Layer list with reorder (up/down), visibility toggle, background transparency toggle
- **Export** — Scope selector, PNG/PDF export buttons, PDF page size/layout/copies controls
- **Project** — Save JSON, Load JSON, Start Over buttons

**Floating palette** (appears when a tile is selected):
- **Basic mode** — Tile name, transparent bg, show outline, position/size inputs, rotation, delete
- **Advanced mode** — Text controls (font, size, alignment, line height, letter spacing, color, bg color), title banner controls, icon value controls

**Top bar:**
- Project name display
- Feedback button (mailto link)
- Tutorial button
- Theme toggle (light/dark)
- Tools menu (links to companion PnP sites)

**Modals:**
- Tutorial prompt (View/Skip on first load)
- Tutorial guide (multi-step walkthrough with Back/Next/Exit)
- Preview card (shows rendered card, download button)

### Constants

```js
MAX_IMAGE_BYTES = 2.5 MB
MAX_IMAGE_DIMENSION = 3000px
SAFE_BORDER_IN = 0.125"
DPI_EXPORT = 300
POINTS_PER_IN = 72
```

### Card sizes

- **Poker:** 2.5" x 3.5" (default)
- **Tarot:** 2.75" x 4.75"
- **Mini:** 1.75" x 2.5"
- **Custom:** user-defined (1-10" width, 1-14" height)

### Page sizes

- **US Letter:** 8.5" x 11"
- **A4:** 8.27" x 11.69"

### Typography

38 Google Fonts loaded via single stylesheet link. Default presets:
- Title: Fjalla One, 82px, centered, all caps
- Effect: Fjalla One, 42px
- Flavor: Cormorant Garamond, 42px

## Key Functions (app.js)

| Function | Purpose |
|----------|---------|
| `bootstrap()` | Entry point — loads draft, sets up event bindings, theme, tooltips |
| `render()` | Rebuilds card stage from `state.design.tiles` |
| `createTile(type)` | Creates a new tile object with defaults |
| `getDefaultTextStyle(type)` | Returns default font/size for tile type |
| `getDefaultImageFill(type)` | Returns default fill for image tile type |
| `activateCard(cardId)` | Switches active card, updates `state.design` |
| `duplicateCard()` | Creates new card with copied design |
| `deleteCard(cardId)` | Removes card from session |
| `addTile(type)` | Creates and adds a new tile to current card |
| `selectTile(tileId)` | Selects tile, shows floating palette |
| `deleteTile(tileId)` | Removes tile from current card |
| `captureCardPngDataUrl()` | Rasterizes card stage via html2canvas |
| `captureExportCards(scope, options)` | Captures one or all cards for export |
| `exportPng()` | Triggers PNG download |
| `exportPdf()` | Creates and downloads print-ready PDF |
| `saveProject()` | Serializes state to JSON file |
| `loadProject(file)` | Reads JSON, normalizes, rebuilds state |
| `resetProject(options)` | Resets to fresh project |
| `syncInputsFromState()` | Syncs DOM inputs to current state |
| `syncStateFromInputs()` | Syncs state from DOM inputs |
| `normalizeDesignData(data)` | Normalizes imported JSON (handles legacy format) |
| `buildStarterDesignFromCurrent(name, settings)` | Creates default starter design |
| `uuid()` | Generates unique IDs for cards and tiles |
| `clamp(value, min, max)` | Utility for value clamping |
| `normalizeRotation(deg)` | Normalizes rotation to 0/90/180/270 |
| `rotateDataUrlClockwise(dataUrl)` | Rotates PNG data URL 90° clockwise |
| `getLayoutConfig(layoutKey)` | Returns layout dimensions for PDF export |
| `getPositions(layout, pageW, pageH, cardW, cardH)` | Calculates card positions on page |
| `drawCutGuides(page, box)` | Draws corner cut guides on PDF page |
| `safeFile(name)` | Sanitizes filename for downloads |
| `validateImageUpload(file)` | Validates PNG/JPG: type, size, dimensions |
| `triggerDownload(url, filename)` | Triggers browser file download |

## Styling Conventions (styles.css)

- CSS custom properties for theming (light/dark via `body[data-theme="dark"]`)
- Light theme: warm paper-like background (`#f2eee6`) with radial gradient highlights
- Dark theme: cool slate background (`#232932`)
- Accent color: blue (`#4a9dff` / `#6eb3ff`)
- Danger color: red (`#d73a49`)
- Border radius: 14px for panels, 12px for modals
- Panels use semi-transparent backgrounds with backdrop blur
- Card stage uses aspect-ratio for card proportions
- Tiles use absolute positioning within the card stage
- Resize handles: gradient blue, bottom-right corner
- Export mode hides UI chrome via `.export-mode` class
- Tooltips: custom CSS-based tooltips on interactive elements
- Responsive: adapts to different viewport sizes

## Running Locally

```bash
python3 -m http.server 8080
# Then open http://localhost:8080
```

No build step required. Just serve the directory and open in a browser.

## Deployment

- **GitHub Pages** — Push to repo, enable Pages from settings
- **Custom domain** — `CNAME` file is included
- **Cloudflare Web Analytics** — Beacon script included in `index.html` with token `e42a535ee3274b5daaa809c5231225df`

## Companion Sites

The Tools menu links to:
- PnPFinder: http://pnpfinder.com
- PnPTools: https://pnptools.gonzhome.us
- PnP Launchpad: https://launchpad.gonzhome.us
- Card Formatter: https://formatter.gonzhome.us
- Card Extractor: https://extractor.gonzhome.us

## Important Implementation Notes

1. **No framework** — Direct DOM manipulation throughout. No virtual DOM, no reactivity system.
2. **localStorage draft** — Auto-saves state on every change, restores on page load. Key: `martins_card_prototyper.draft`.
3. **Tutorial persistence** — Tutorial seen flag stored in localStorage: `martins_card_prototyper.tutorial_seen`.
4. **Theme persistence** — Theme preference stored in localStorage: `martins_card_prototyper.theme`.
5. **Legacy format support** — `normalizeDesignData()` handles both current multi-card JSON and legacy single-card JSON files.
6. **Image upload validation** — Strict limits: PNG/JPG only, max 2.5 MB, max 3000x3000px.
7. **PDF rotation** — Gutterfold layout requires clockwise rotation of card images for proper booklet orientation.
8. **html2canvas config** — `ignoreElements` hides UI chrome during export; `backgroundColor: null` preserves transparency.
9. **Safe border** — 0.125" inset from card edges, visible during editing, hidden during export.
10. **Snap-to-grid** — Default step is 0.005" (adjustable 0.005"-0.5"), can be disabled for freeform placement.
