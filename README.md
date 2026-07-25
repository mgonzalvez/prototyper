# Martin's Card Prototyper

Martin's Card Prototyper is a static web app for quickly building tabletop card mockups in the browser.

No build step is required. It is designed to run locally or from a static host such as GitHub Pages.
An internet connection is needed to load Google Fonts, `pdf-lib`, and `html2canvas` from their CDNs.

## Current Features

- Drag-and-drop tile-based card editor.
- Shared PnP-tools interface aligned with BoardSplitter, Card Formatter, and Card Extractor:
  - Sticky branded application header
  - Blue/violet light and dark themes
  - Related Sites navigation
  - Responsive glass-panel workspace
- Responsive card canvas sized for legible editing on desktop, tablet, and mobile.
- Multi-card session workflow:
  - Create, duplicate, delete, and switch cards in one project session
  - Each card keeps independent tile/content/layer edits
- Poker, tarot, mini, and custom card sizes.
- Built-in tile types:
  - Card Background Image
  - Title
  - Effect Text
  - Flavor Text
  - Main Image
  - Icon
- Default starter layout on new project:
  - Title tile at top (centered, all caps)
  - Main image tile in middle
  - Effect text tile in lower third
- Default text style presets:
  - Title: Fjalla One, 82 px, centered
  - Effect: Fjalla One, 42 px
  - Flavor: Cormorant Garamond, 42 px
- Card-safe border guardrails and snap-to-grid placement (with optional freeform).
- Rich text controls for text tiles:
  - Bold, italic, underline
  - Horizontal alignment (dropdown)
  - Vertical alignment
  - Font family/size
  - Line height and letter spacing
  - Text color and tile background color
  - Text flow mode (rotated flow or vertical writing)
- Title banner options:
  - Solid color banner
  - Gradient banner
  - Uploaded image banner (auto-fits title tile)
  - Banner padding control (default 0.5 mm)
- Icon value controls:
  - Styled numeric value overlay
  - Color, size, outline, X/Y nudge
- Layer management panel:
  - Reorder layers up/down
  - Toggle layer visibility
  - Toggle layer background transparency
- Per-tile controls:
  - Basic and advanced inspector modes for title and icon tiles
  - Editable layer label
  - Position and size
  - Rotation (0°, 90°, 180°, 270°)
  - Tile background transparency
  - Tile outline visibility
- Image-tile fill controls (Card Background, Main Image, Icon):
  - Fill mode: Image, Solid Color, Gradient
  - Light gradient presets
  - Uploaded images always render in the foreground
- Upload validation for PNG/JPG images:
  - Max file size: 2.5 MB
  - Max dimensions: 3000x3000
- Export options:
  - Export scope: Active Card or All Cards in Session
  - Preview Card before export
  - PNG:
    - Active card export
    - All-cards export as separate PNG files
  - PDF (print-ready tiled layout)
    - Active card (uses Cards in PDF copy count)
    - All cards in session (each card placed once, in order)
  - US Letter or A4
  - 3x3, 2x3, or Gutterfold layouts
  - Corner cut guides drawn above card images
  - Rotation-aware text rendering in preview, PNG, and PDF
  - Export uses per-tile settings for outline/background visibility
- Project file workflow:
  - Save project as JSON (includes all cards in the session)
  - Load project JSON (supports both current multi-card and legacy single-card files)
  - Start Over reset (keeps the active card size and grid/safe-border settings)
- In-app onboarding:
  - Tutorial prompt with View/Skip on first load
  - Multi-step guided walkthrough with Back/Next/Exit, anchored to relevant UI
  - Reopen tutorial from top bar
- Header utilities:
  - Feedback button (`mailto:help@pnpfinder.com` with prefilled subject)
  - Theme toggle with system-theme fallback
  - Related Sites menu linking to companion PnP tools
- Cloudflare Web Analytics snippet included in `index.html`.
- Mouseover tooltips on major UI controls.

## Project and Browser Storage

- The **Project Name** field names the active card. It is used for that card's entry and export filenames.
- **Save Project JSON** is the durable save mechanism. The browser does not auto-save the card project or restore it after a refresh.
- The browser stores only the selected theme and whether the introductory tutorial has been dismissed.
- Uploaded images are embedded as data URLs in saved project JSON files, so image-heavy projects can produce large files.

## Quick Start

1. Start with the default tiles already on the card:
   - **Title**
   - **Main Image**
   - **Effect Text**
2. Type your card title directly into the title tile.
3. Upload a main image.
   - `game-icons.net` is a great source of free prototype assets.
4. Edit your effect text directly in the effect tile.
5. Optionally add:
   - **Icon Tiles** with styled numeric values
   - **Flavor Text Tile** for supporting text
   - **Card Background Tile** with image, solid color, or gradient fill
6. Use the **Cards** section to:
   - create new cards
   - duplicate existing cards
   - switch between cards in the same session
7. Use **Preview Card** to verify output.
8. Choose **Export Scope**:
   - **Active Card** for one card, or
   - **All Cards in Session** for batch export
9. Export as:
   - **PNG** (single or batch by scope), or
   - **print-ready PDF** (single-card copies or full-session sheet export)
10. Optionally:
   - **Save JSON** to keep your full session project file
   - **Load JSON** to continue an in-progress project

Save the project JSON before refreshing, closing the tab, or using **Start Over**.

## In-App Tutorial Walkthrough

When you open the site, you can choose:

- **View Tutorial**: shows a short guided walkthrough
- **Skip**: closes onboarding and opens the editor directly

You can reopen the walkthrough anytime using the **Tutorial** button in the top bar.

## Run Locally

From the project folder:

```bash
python3 -m http.server 8080
```

Then open:

- <http://localhost:8080>

There is no package install, build command, test suite, or hot reload.
The stylesheet URL in `index.html` includes a cache-busting query value; increment it when deployed CSS changes need to bypass browser or CDN caches.

## Deployment

The repository is configured for GitHub Pages with the custom domain <https://prototyper.gonzhome.us> in `CNAME`.
