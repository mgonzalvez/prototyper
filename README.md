# Martin's Card Prototyper

Martin's Card Prototyper is a static web app for quickly building tabletop card mockups in the browser.

No build step is required. It is designed to run locally or on GitHub Pages.

## Current Features

- Drag-and-drop tile-based card editor.
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
- Card-safe border guardrails and snap-to-grid placement (with optional freeform).
- Rich text controls for text tiles:
  - Bold, italic, underline
  - Horizontal alignment
  - Vertical alignment
  - Font family/size
  - Line height and letter spacing
  - Text color and tile background color
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
  - Preview Card before export
  - PNG (single card)
  - PDF (print-ready tiled layout)
  - US Letter or A4
  - 3x3, 2x3, or Gutterfold layouts
  - Corner cut guides drawn above card images
  - Export uses per-tile settings for outline/background visibility
- Project file workflow:
  - Save project as JSON
  - Load project JSON
  - Start Over reset
- In-app onboarding:
  - Tutorial prompt with View/Skip on first load
  - Multi-step guided walkthrough with Next/Back, anchored to relevant UI
  - Reopen tutorial from top bar
- Top bar utilities:
  - Feedback button (`mailto:help@pnpfinder.com` with prefilled subject)
  - Theme toggle (light/dark)
  - Tools menu linking to companion PnP sites
- Cloudflare Web Analytics snippet included in `index.html`.
- Mouseover tooltips on major UI controls.

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
6. Use **Preview Card** to verify output.
7. Export your card as:
   - an individual **PNG**, or
   - a tiled **print-ready PDF**
8. Optionally:
   - **Save JSON** to keep your project file
   - **Load JSON** to continue an in-progress project

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
