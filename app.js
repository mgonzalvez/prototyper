/* global PDFLib, html2canvas */

const MAX_IMAGE_BYTES = 2.5 * 1024 * 1024;
const MAX_IMAGE_DIMENSION = 3000;
const SAFE_BORDER_IN = 0.125;
const DPI_EXPORT = 300;
const POINTS_PER_IN = 72;
const TUTORIAL_SEEN_KEY = "martins_card_prototyper.tutorial_seen";
const THEME_KEY = "martins_card_prototyper.theme";

const cardSizes = {
  poker: { w: 2.5, h: 3.5 },
  tarot: { w: 2.75, h: 4.75 },
  mini: { w: 1.75, h: 2.5 },
};

const pageSizes = {
  letter: { w: 8.5, h: 11 },
  a4: { w: 8.27, h: 11.69 },
};

const els = {
  projectName: document.querySelector("#projectName"),
  cardSizeSelect: document.querySelector("#cardSizeSelect"),
  customSizeWrap: document.querySelector("#customSizeWrap"),
  customWidth: document.querySelector("#customWidth"),
  customHeight: document.querySelector("#customHeight"),
  snapToggle: document.querySelector("#snapToggle"),
  gridStepInput: document.querySelector("#gridStepInput"),
  showSafeBorder: document.querySelector("#showSafeBorder"),
  feedbackBtn: document.querySelector("#feedbackBtn"),
  toolsMenuBtn: document.querySelector("#toolsMenuBtn"),
  toolsMenu: document.querySelector("#toolsMenu"),
  themeToggleBtn: document.querySelector("#themeToggleBtn"),
  openTutorialBtn: document.querySelector("#openTutorialBtn"),
  tutorialModal: document.querySelector("#tutorialModal"),
  tutorialChoiceRow: document.querySelector("#tutorialChoiceRow"),
  tutorialViewBtn: document.querySelector("#tutorialViewBtn"),
  tutorialSkipBtn: document.querySelector("#tutorialSkipBtn"),
  tutorialGuide: document.querySelector("#tutorialGuide"),
  tutorialGuideStep: document.querySelector("#tutorialGuideStep"),
  tutorialGuideText: document.querySelector("#tutorialGuideText"),
  tutorialPrevBtn: document.querySelector("#tutorialPrevBtn"),
  tutorialNextBtn: document.querySelector("#tutorialNextBtn"),
  tutorialGuideSkipBtn: document.querySelector("#tutorialGuideSkipBtn"),
  tilePalette: document.querySelector("#tilePalette"),
  layersList: document.querySelector("#layersList"),
  saveProjectBtn: document.querySelector("#saveProjectBtn"),
  loadProjectBtn: document.querySelector("#loadProjectBtn"),
  projectFileInput: document.querySelector("#projectFileInput"),
  resetProjectBtn: document.querySelector("#resetProjectBtn"),
  previewCardBtn: document.querySelector("#previewCardBtn"),
  exportPngBtn: document.querySelector("#exportPngBtn"),
  pdfPageSize: document.querySelector("#pdfPageSize"),
  pdfLayout: document.querySelector("#pdfLayout"),
  pdfCopies: document.querySelector("#pdfCopies"),
  exportPdfBtn: document.querySelector("#exportPdfBtn"),
  previewModal: document.querySelector("#previewModal"),
  previewImage: document.querySelector("#previewImage"),
  previewDownloadBtn: document.querySelector("#previewDownloadBtn"),
  previewCloseBtn: document.querySelector("#previewCloseBtn"),
  statusText: document.querySelector("#statusText"),
  cardStageWrap: document.querySelector("#cardStageWrap"),
  cardStage: document.querySelector("#cardStage"),
  floatingPalette: document.querySelector("#floatingPalette"),
  tileControls: document.querySelector("#tileControls"),
  textControlsWrap: document.querySelector("#textControlsWrap"),
  tileNameInput: document.querySelector("#tileNameInput"),
  tileTransparentBgInput: document.querySelector("#tileTransparentBgInput"),
  tileShowOutlineInput: document.querySelector("#tileShowOutlineInput"),
  tileXInput: document.querySelector("#tileXInput"),
  tileYInput: document.querySelector("#tileYInput"),
  tileWInput: document.querySelector("#tileWInput"),
  tileHInput: document.querySelector("#tileHInput"),
  tileRotationSelect: document.querySelector("#tileRotationSelect"),
  deleteTileBtn: document.querySelector("#deleteTileBtn"),
  fontFamilySelect: document.querySelector("#fontFamilySelect"),
  fontSizeInput: document.querySelector("#fontSizeInput"),
  lineHeightInput: document.querySelector("#lineHeightInput"),
  textFlowModeSelect: document.querySelector("#textFlowModeSelect"),
  textVAlignSelect: document.querySelector("#textVAlignSelect"),
  letterSpacingInput: document.querySelector("#letterSpacingInput"),
  textColorInput: document.querySelector("#textColorInput"),
  bgColorInput: document.querySelector("#bgColorInput"),
  titleControls: document.querySelector("#titleControls"),
  titleBannerEnabled: document.querySelector("#titleBannerEnabled"),
  titleBannerMode: document.querySelector("#titleBannerMode"),
  titleBannerColorRow: document.querySelector("#titleBannerColorRow"),
  titleBannerColor: document.querySelector("#titleBannerColor"),
  titleBannerGradientRow: document.querySelector("#titleBannerGradientRow"),
  titleBannerGradient: document.querySelector("#titleBannerGradient"),
  titleBannerImageRow: document.querySelector("#titleBannerImageRow"),
  titleBannerUploadBtn: document.querySelector("#titleBannerUploadBtn"),
  titleBannerFileInput: document.querySelector("#titleBannerFileInput"),
  titleBannerPaddingMm: document.querySelector("#titleBannerPaddingMm"),
  imageControls: document.querySelector("#imageControls"),
  imageFillMode: document.querySelector("#imageFillMode"),
  imageFillColorRow: document.querySelector("#imageFillColorRow"),
  imageFillColor: document.querySelector("#imageFillColor"),
  imageFillGradientRow: document.querySelector("#imageFillGradientRow"),
  imageFillGradient: document.querySelector("#imageFillGradient"),
  uploadImageBtn: document.querySelector("#uploadImageBtn"),
  imageFileInput: document.querySelector("#imageFileInput"),
  iconControls: document.querySelector("#iconControls"),
  iconValueInput: document.querySelector("#iconValueInput"),
  iconValueColor: document.querySelector("#iconValueColor"),
  iconValueSize: document.querySelector("#iconValueSize"),
  iconOutlineColor: document.querySelector("#iconOutlineColor"),
  iconOutlineWidth: document.querySelector("#iconOutlineWidth"),
  iconOffsetX: document.querySelector("#iconOffsetX"),
  iconOffsetY: document.querySelector("#iconOffsetY"),
  toolbarButtons: [...document.querySelectorAll(".toolbar button")],
};

const state = {
  projectId: null,
  selectedTileId: null,
  dragTileType: null,
  liveEditable: null,
  dragAction: null,
  tutorial: {
    active: false,
    index: 0,
    highlightedSelector: "",
  },
  design: {
    name: "Untitled Design",
    card: { preset: "poker", widthIn: 2.5, heightIn: 3.5 },
    snapEnabled: true,
    gridStepIn: 0.005,
    showSafeBorder: true,
    tiles: [],
    updatedAt: new Date().toISOString(),
  },
};

function buildDefaultDesign(overrides = {}) {
  const defaults = {
    name: "Untitled Design",
    card: { preset: "poker", widthIn: 2.5, heightIn: 3.5 },
    snapEnabled: true,
    gridStepIn: 0.005,
    showSafeBorder: true,
  };
  const next = {
    ...defaults,
    ...overrides,
    card: {
      ...defaults.card,
      ...(overrides.card || {}),
    },
  };
  return {
    name: next.name,
    card: next.card,
    snapEnabled: next.snapEnabled,
    gridStepIn: next.gridStepIn,
    showSafeBorder: next.showSafeBorder,
    tiles: [],
    updatedAt: new Date().toISOString(),
  };
}

function setStatus(text) {
  els.statusText.textContent = text;
}

function applyTheme(theme) {
  const next = theme === "dark" ? "dark" : "light";
  document.body.setAttribute("data-theme", next);
  if (els.themeToggleBtn) {
    els.themeToggleBtn.textContent = next === "dark" ? "Light Mode" : "Dark Mode";
    els.themeToggleBtn.title = next === "dark" ? "Switch to light mode." : "Switch to dark mode.";
  }
}

function initTheme() {
  const saved = localStorage.getItem(THEME_KEY);
  if (saved === "dark" || saved === "light") {
    applyTheme(saved);
    return;
  }
  applyTheme("light");
}

function applyTooltips() {
  const tips = {
    projectName: "Name your current card project.",
    cardSizeSelect: "Choose a preset card size or Custom.",
    customWidth: "Set custom card width.",
    customHeight: "Set custom card height.",
    snapToggle: "Snap tile movement/resizing to grid increments.",
    gridStepInput: "Grid increment size used when snapping.",
    showSafeBorder: "Show or hide the safe print border guide.",
    feedbackBtn: "Send feedback by email.",
    saveProjectBtn: "Download this project as a JSON file.",
    loadProjectBtn: "Load a previously saved project JSON file.",
    resetProjectBtn: "Start over with the default starter tiles.",
    previewCardBtn: "Preview the card image before exporting.",
    exportPngBtn: "Export the current card as a PNG image.",
    pdfPageSize: "Choose paper size for PDF sheet export.",
    pdfLayout: "Choose tiled sheet arrangement for PDF export.",
    pdfCopies: "Number of card copies to place into the PDF export.",
    exportPdfBtn: "Export the card as a tiled print-ready PDF.",
    tileNameInput: "Rename the selected tile.",
    tileTransparentBgInput: "Make selected tile background transparent.",
    tileShowOutlineInput: "Show or hide the selected tile outline.",
    tileXInput: "Set selected tile X position.",
    tileYInput: "Set selected tile Y position.",
    tileWInput: "Set selected tile width.",
    tileHInput: "Set selected tile height.",
    tileRotationSelect: "Rotate selected tile content in 90-degree steps.",
    deleteTileBtn: "Delete the currently selected tile.",
    fontFamilySelect: "Choose font family for selected text tile.",
    fontSizeInput: "Set font size for selected text tile.",
    lineHeightInput: "Set line spacing for selected text tile.",
    textFlowModeSelect: "Choose rotated text flow or vertical writing mode.",
    textVAlignSelect: "Set vertical alignment inside the text tile.",
    letterSpacingInput: "Set spacing between letters.",
    textColorInput: "Set text color.",
    bgColorInput: "Set text tile background color.",
    titleBannerEnabled: "Toggle a banner behind title text.",
    titleBannerMode: "Choose banner style: color, gradient, or image.",
    titleBannerColor: "Choose title banner solid color.",
    titleBannerGradient: "Choose title banner gradient style.",
    titleBannerUploadBtn: "Upload an image for the title banner.",
    titleBannerPaddingMm: "Set title text padding inside the banner (mm).",
    imageFillMode: "Choose image, solid color, or gradient fill for selected image tile.",
    imageFillColor: "Choose a solid fill color for selected image tile.",
    imageFillGradient: "Choose a gradient fill style for selected image tile.",
    uploadImageBtn: "Upload image for selected image/icon/background tile.",
    iconValueInput: "Set icon numeric value text.",
    iconValueColor: "Set icon value text color.",
    iconValueSize: "Set icon value font size.",
    iconOutlineColor: "Set icon value outline color.",
    iconOutlineWidth: "Set icon value outline width.",
    iconOffsetX: "Nudge icon value horizontally.",
    iconOffsetY: "Nudge icon value vertically.",
    openTutorialBtn: "Open the quick walkthrough tutorial.",
    toolsMenuBtn: "Open links to your other PnP sites.",
    themeToggleBtn: "Toggle between light and dark themes.",
  };

  Object.entries(tips).forEach(([id, text]) => {
    const el = document.getElementById(id);
    if (el) el.title = text;
  });

  els.tilePalette?.querySelectorAll("button").forEach((btn) => {
    const type = btn.dataset.tileType || "tile";
    btn.title = `Drag this ${type.replace("-", " ")} tile onto the card.`;
  });

  els.toolbarButtons?.forEach((btn) => {
    const cmd = btn.dataset.cmd || "format";
    btn.title = `Apply ${cmd} formatting to selected text.`;
  });
}

function closePreviewModal() {
  if (!els.previewModal) return;
  els.previewModal.hidden = true;
}

function sortFontFamilyOptions() {
  if (!els.fontFamilySelect) return;
  const selectedValue = els.fontFamilySelect.value;
  const options = Array.from(els.fontFamilySelect.options);
  options.sort((a, b) =>
    a.textContent.trim().localeCompare(b.textContent.trim(), undefined, { sensitivity: "base" })
  );
  els.fontFamilySelect.innerHTML = "";
  options.forEach((option) => {
    els.fontFamilySelect.appendChild(option);
  });
  if (options.some((option) => option.value === selectedValue)) {
    els.fontFamilySelect.value = selectedValue;
  }
}

function openTutorialPrompt(force = false) {
  if (!els.tutorialModal) return;
  if (!force && sessionStorage.getItem(TUTORIAL_SEEN_KEY) === "1") return;
  els.tutorialChoiceRow.hidden = false;
  els.tutorialModal.hidden = false;
}

function getTutorialSteps() {
  return [
    {
      selector: "#projectName",
      text: "Start by naming your card project.",
    },
    {
      selector: "#cardStage .tile.title",
      text: "The Title tile is already on the card by default. Click it and type your card title.",
    },
    {
      selector: "#cardStage .tile.main-image",
      text: "The Main Image tile is also preloaded. Select it, then use Upload Image to add artwork (game-icons.net is a great free source).",
    },
    {
      selector: "#cardStage .tile.effect",
      text: "The Effect text tile is preloaded too. Click into it and enter your card rules text.",
    },
    {
      selector: "#tileRotationSelect",
      text: "Need vertical or rotated layouts? With a tile selected, use Rotation (0/90/180/270). For text tiles, Text Flow lets you choose Rotated Text or Vertical Writing.",
    },
    {
      selector: "#tilePalette",
      text: "After the core tiles are set, add optional tiles such as Icon, Flavor Text, or Card Background.",
    },
    {
      selector: "#layersList",
      text: "Use Layers to reorder tiles, hide layers, and make tile backgrounds transparent.",
    },
    {
      selector: "#exportPngBtn",
      text: "Export a single card image as PNG.",
    },
    {
      selector: "#exportPdfBtn",
      text: "Export print-ready tiled PDF sheets.",
    },
    {
      selector: "#saveProjectBtn",
      text: "Optionally save your work as JSON, then load that JSON later to continue.",
    },
  ];
}

function clearTutorialHighlight() {
  if (!state.tutorial.highlightedSelector) return;
  const prev = document.querySelector(state.tutorial.highlightedSelector);
  if (prev) prev.classList.remove("tutorial-target");
  state.tutorial.highlightedSelector = "";
}

function closeTutorial(markSeen = true) {
  if (els.tutorialModal) els.tutorialModal.hidden = true;
  if (els.tutorialGuide) els.tutorialGuide.hidden = true;
  clearTutorialHighlight();
  state.tutorial.active = false;
  if (markSeen) {
    sessionStorage.setItem(TUTORIAL_SEEN_KEY, "1");
  }
}

function positionTutorialGuide(target) {
  if (!els.tutorialGuide) return;
  const card = els.tutorialGuide.querySelector(".tutorial-guide-card");
  if (!card || !target) return;
  const rect = target.getBoundingClientRect();
  const cardRect = card.getBoundingClientRect();
  const margin = 12;

  let left = rect.right + margin;
  if (left + cardRect.width > window.innerWidth - 10) {
    left = rect.left - cardRect.width - margin;
  }
  left = clamp(left, 10, Math.max(10, window.innerWidth - cardRect.width - 10));

  let top = rect.top + (rect.height - cardRect.height) / 2;
  top = clamp(top, 72, Math.max(72, window.innerHeight - cardRect.height - 10));

  card.style.left = `${Math.round(left)}px`;
  card.style.top = `${Math.round(top)}px`;
}

function ensureTutorialTargetVisible(target) {
  if (!target) return;
  // Let the browser scroll any relevant container (including overflow panels)
  // so the target can be highlighted even when it's below the fold.
  target.scrollIntoView({
    block: "center",
    inline: "nearest",
    behavior: "auto",
  });
}

function renderTutorialGuide() {
  if (!state.tutorial.active || !els.tutorialGuide) return;
  const steps = getTutorialSteps();
  let index = clamp(state.tutorial.index, 0, steps.length - 1);
  let step = steps[index];
  let target = step ? document.querySelector(step.selector) : null;

  if (!target) {
    for (let i = 0; i < steps.length; i += 1) {
      const candidate = document.querySelector(steps[i].selector);
      if (candidate) {
        index = i;
        step = steps[i];
        target = candidate;
        break;
      }
    }
  }

  if (!target || !step) {
    closeTutorial(true);
    return;
  }

  state.tutorial.index = index;
  ensureTutorialTargetVisible(target);
  clearTutorialHighlight();
  target.classList.add("tutorial-target");
  state.tutorial.highlightedSelector = step.selector;

  els.tutorialGuideStep.textContent = `Step ${index + 1} of ${steps.length}`;
  els.tutorialGuideText.textContent = step.text;
  els.tutorialPrevBtn.disabled = index === 0;
  els.tutorialNextBtn.textContent = index === steps.length - 1 ? "Done" : "Next";
  els.tutorialGuide.hidden = false;
  positionTutorialGuide(target);
}

function startTutorialGuide() {
  state.tutorial.active = true;
  state.tutorial.index = 0;
  if (els.tutorialModal) els.tutorialModal.hidden = true;
  renderTutorialGuide();
}

function inchesToPoints(inches) {
  return inches * POINTS_PER_IN;
}

function mmToPx(mm) {
  return (Number(mm) * 96) / 25.4;
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function normalizeRotation(value) {
  const n = Number(value);
  if (!Number.isFinite(n)) return 0;
  const normalized = ((Math.round(n / 90) * 90) % 360 + 360) % 360;
  return [0, 90, 180, 270].includes(normalized) ? normalized : 0;
}

function uuid() {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function getCardSize() {
  return {
    w: state.design.card.widthIn,
    h: state.design.card.heightIn,
  };
}

function getPxPerIn() {
  const { w } = getCardSize();
  return els.cardStage.clientWidth / w;
}

function toInches(px) {
  return px / getPxPerIn();
}

function toPixels(inches) {
  return inches * getPxPerIn();
}

function snap(inches) {
  if (!state.design.snapEnabled) return inches;
  const step = Math.max(0.005, state.design.gridStepIn || 0.125);
  return Math.round(inches / step) * step;
}

function safeBounds() {
  const size = getCardSize();
  return {
    left: SAFE_BORDER_IN,
    top: SAFE_BORDER_IN,
    right: size.w - SAFE_BORDER_IN,
    bottom: size.h - SAFE_BORDER_IN,
  };
}

function constrainTile(tile) {
  const size = getCardSize();
  const bounds =
    tile.type === "card-background"
      ? { left: 0, top: 0, right: size.w, bottom: size.h }
      : safeBounds();
  const minW = 0.3;
  const minH = 0.3;
  tile.wIn = clamp(tile.wIn, minW, bounds.right - bounds.left);
  tile.hIn = clamp(tile.hIn, minH, bounds.bottom - bounds.top);
  tile.xIn = clamp(tile.xIn, bounds.left, bounds.right - tile.wIn);
  tile.yIn = clamp(tile.yIn, bounds.top, bounds.bottom - tile.hIn);
}

function getDefaultTextStyle(type) {
  if (type === "title") {
    return {
      fontFamily: "'Playfair Display', 'Times New Roman', serif",
      fontSize: 78,
      lineHeight: 1.2,
      textFlowMode: "rotated",
      verticalAlign: "middle",
      letterSpacing: 0,
      textAlign: "center",
      color: "#1f2430",
      background: "#ffffff",
    };
  }
  if (type === "flavor") {
    return {
      fontFamily: "'Cormorant Garamond', Georgia, serif",
      fontSize: 28,
      lineHeight: 1.35,
      textFlowMode: "rotated",
      verticalAlign: "middle",
      letterSpacing: 0,
      textAlign: "left",
      color: "#2e3440",
      background: "#ffffff",
    };
  }
  return {
    fontFamily: "'Manrope', 'Avenir Next', 'Helvetica Neue', Arial, sans-serif",
    fontSize: 36,
    lineHeight: 1.25,
    textFlowMode: "rotated",
    verticalAlign: "middle",
    letterSpacing: 0,
    textAlign: "left",
    color: "#1f2430",
    background: "#ffffff",
  };
}

function getDefaultImageFill(type) {
  if (type === "card-background") {
    return {
      mode: "gradient",
      color: "#f8f8f8",
      gradient: "linear-gradient(180deg, #ffffff, #e8f2ff)",
    };
  }
  return {
    mode: "image",
    color: "#f8f8f8",
    gradient: "linear-gradient(180deg, #ffffff, #e8f2ff)",
  };
}

function createTile(type, xIn, yIn) {
  const size = getCardSize();
  const isAutoCenter = xIn == null && yIn == null;
  const isImageType = type === "main-image" || type === "card-background";
  const presets = {
    "card-background": { w: size.w, h: size.h, x: 0, y: 0 },
    title: { w: 1.9, h: 0.7, y: SAFE_BORDER_IN },
    effect: { w: 1.9, h: 1.1, y: 2.0 },
    flavor: { w: 1.9, h: 0.75, y: 2.85 },
    "main-image": { w: 1.9, h: 1.45, y: 0.95 },
    icon: { w: 0.55, h: 0.55, y: SAFE_BORDER_IN + 0.05 },
  };
  const base = presets[type] || presets.effect;
  const centerX = (size.w - base.w) / 2;
  const centerY = (size.h - base.h) / 2;
  const bounds = safeBounds();
  const gap = 0.06;
  const titleBottomY = presets.title.y + presets.title.h;
  const effectY = bounds.bottom - presets.effect.h;
  const mainTop = titleBottomY + gap;
  const mainBottom = effectY - gap;
  const mainAvailable = Math.max(0.6, mainBottom - mainTop);
  const mainImageH = Math.min(presets["main-image"].h, mainAvailable);
  const mainImageY = mainTop + Math.max(0, (mainAvailable - mainImageH) / 2);
  const defaultY =
    type === "title"
      ? base.y
      : type === "card-background"
        ? 0
      : type === "main-image"
        ? mainImageY
        : type === "effect"
          ? effectY
          : centerY;
  const tile = {
    id: uuid(),
    type,
    label: type.replace("-", " "),
    xIn: xIn ?? (type === "card-background" ? 0 : centerX),
    yIn: yIn ?? defaultY,
    wIn: base.w,
    hIn: type === "main-image" ? mainImageH : base.h,
    rotationDeg: 0,
    textHtml: isImageType || type === "icon" ? "" : defaultHtml(type),
    titlePlacement: "top",
    hidden: false,
    transparentBg: type === "card-background" || type === "main-image" || type === "icon",
    showOutline: true,
    titleBanner: {
      enabled: false,
      mode: "color",
      color: "#1f2430",
      gradient: "linear-gradient(90deg, #1f2937, #3b82f6)",
      imageDataUrl: "",
      paddingMm: 0.5,
    },
    style: getDefaultTextStyle(type),
    imageDataUrl: "",
    imageFill: getDefaultImageFill(type),
    icon: {
      value: "1",
      position: "overlay",
      valueColor: "#ffffff",
      valueSize: 60,
      outlineColor: "#111111",
      outlineWidth: 2,
      offsetX: 0,
      offsetY: 0,
    },
  };

  constrainTile(tile);
  if (!isAutoCenter) {
    tile.xIn = snap(tile.xIn);
    tile.yIn = snap(tile.yIn);
  }
  return tile;
}

function defaultHtml(type) {
  if (type === "title") return "<div>CARD TITLE</div>";
  if (type === "effect") return "<div>Effect rules text goes here.</div>";
  if (type === "flavor") return "<div><em>Flavor text goes here.</em></div>";
  return "<div>Text</div>";
}

function updateStageGeometry() {
  const size = getCardSize();
  els.cardStageWrap.style.aspectRatio = `${size.w} / ${size.h}`;
  const left = `${(SAFE_BORDER_IN / size.w) * 100}%`;
  const right = `${(SAFE_BORDER_IN / size.w) * 100}%`;
  const top = `${(SAFE_BORDER_IN / size.h) * 100}%`;
  const bottom = `${(SAFE_BORDER_IN / size.h) * 100}%`;
  els.cardStage.style.setProperty("--safe-border-left", left);
  els.cardStage.style.setProperty("--safe-border-right", right);
  els.cardStage.style.setProperty("--safe-border-top", top);
  els.cardStage.style.setProperty("--safe-border-bottom", bottom);
  els.cardStage.style.setProperty("--safe-border-opacity", state.design.showSafeBorder ? "0.85" : "0");
}

function getSelectedTile() {
  return state.design.tiles.find((t) => t.id === state.selectedTileId) || null;
}

function syncSelectionVisuals() {
  const tileEls = els.cardStage.querySelectorAll(".tile");
  tileEls.forEach((node) => {
    const tileIndex = state.design.tiles.findIndex((t) => t.id === node.dataset.tileId);
    const baseZ = tileIndex >= 0 ? tileIndex + 1 : 1;
    const active = node.dataset.tileId === state.selectedTileId;
    node.classList.toggle("selected", active);
    node.style.zIndex = active ? String(baseZ + 100) : String(baseZ);
  });
}

function renderTile(tile) {
  const tileEl = document.createElement("div");
  tileEl.className = `tile ${tile.type}`;
  if (tile.showOutline === false) tileEl.classList.add("no-outline");
  tileEl.dataset.tileId = tile.id;
  if (state.selectedTileId === tile.id) tileEl.classList.add("selected");
  tileEl.style.left = `${toPixels(tile.xIn)}px`;
  tileEl.style.top = `${toPixels(tile.yIn)}px`;
  tileEl.style.width = `${toPixels(tile.wIn)}px`;
  tileEl.style.height = `${toPixels(tile.hIn)}px`;
  tileEl.style.background = tile.transparentBg ? "transparent" : tile.style.background || "#ffffff";
  const layerIndex = state.design.tiles.findIndex((t) => t.id === tile.id);
  const baseZ = layerIndex >= 0 ? layerIndex + 1 : 1;
  tileEl.style.zIndex = state.selectedTileId === tile.id ? String(baseZ + 100) : String(baseZ);

  const header = document.createElement("div");
  header.className = "tile-header export-hide";
  header.textContent = tile.label || tile.type;
  tileEl.appendChild(header);

  const deleteBtn = document.createElement("button");
  deleteBtn.className = "tile-delete export-hide";
  deleteBtn.type = "button";
  deleteBtn.setAttribute("aria-label", "Delete tile");
  deleteBtn.textContent = "X";
  tileEl.appendChild(deleteBtn);

  const content = document.createElement("div");
  content.className = "tile-content";
  content.style.transformOrigin = "50% 50%";

  if (isTextTile(tile)) {
    const editable = document.createElement("div");
    editable.className = "tile-text-editable";
    editable.dataset.tileContentId = tile.id;
    editable.contentEditable = "true";
    editable.spellcheck = false;
    editable.innerHTML = tile.textHtml || "";
    applyTextStyle(editable, tile);
    applyTextFlowLayout(editable, tile);
    editable.addEventListener("pointerdown", (e) => {
      e.stopPropagation();
      if (state.selectedTileId !== tile.id) {
        state.selectedTileId = tile.id;
        syncSelectionVisuals();
        renderSelectionPanel();
      }
    });
    editable.addEventListener("focus", () => {
      state.liveEditable = editable;
      if (state.selectedTileId !== tile.id) {
        state.selectedTileId = tile.id;
        syncSelectionVisuals();
      }
      renderSelectionPanel();
    });
    editable.addEventListener("blur", () => {
      if (state.liveEditable === editable) {
        state.liveEditable = null;
      }
    });
    editable.addEventListener("input", () => {
      tile.textHtml = editable.innerHTML;
      persistDraft();
    });
    content.appendChild(editable);
  } else {
    content.style.rotate = `${normalizeRotation(tile.rotationDeg)}deg`;
    const fill = {
      ...getDefaultImageFill(tile.type),
      ...(tile.imageFill || {}),
    };
    const hasImage = Boolean(tile.imageDataUrl);
    content.innerHTML = "";
    content.style.backgroundColor = "transparent";
    content.style.backgroundImage = "none";
    content.style.backgroundRepeat = "no-repeat";
    content.style.backgroundPosition = "center";
    content.style.backgroundSize = "cover";

    const layers = [];
    const layerSizes = [];
    const layerRepeats = [];
    const layerPositions = [];

    if (hasImage) {
      layers.push(`url(${tile.imageDataUrl})`);
      layerSizes.push(tile.type === "card-background" ? "cover" : "contain");
      layerRepeats.push("no-repeat");
      layerPositions.push("center");
    }
    if (fill.mode === "color") {
      content.style.backgroundColor = fill.color || "#f8f8f8";
    } else if (fill.mode === "gradient") {
      layers.push(fill.gradient || "linear-gradient(180deg, #ffffff, #e8f2ff)");
      layerSizes.push("cover");
      layerRepeats.push("no-repeat");
      layerPositions.push("center");
    }

    if (layers.length > 0) {
      content.style.backgroundImage = layers.join(", ");
      content.style.backgroundSize = layerSizes.join(", ");
      content.style.backgroundRepeat = layerRepeats.join(", ");
      content.style.backgroundPosition = layerPositions.join(", ");
    }

    if (!hasImage && fill.mode === "image") {
      content.innerHTML = "<div style='padding:8px;color:#67717d;font-size:12px;'>No image yet</div>";
    }
  }

  tileEl.appendChild(content);

  if (tile.type === "icon") {
    const value = document.createElement("div");
    value.className = "tile-icon-value overlay";
    value.textContent = tile.icon.value || "1";
    value.style.color = tile.icon.valueColor || "#fff";
    value.style.fontSize = `${tile.icon.valueSize || 24}px`;
    value.style.marginLeft = `${Number(tile.icon.offsetX || 0)}px`;
    value.style.marginTop = `${Number(tile.icon.offsetY || 0)}px`;
    value.style.rotate = `${normalizeRotation(tile.rotationDeg)}deg`;
    const outline = Number(tile.icon.outlineWidth || 0);
    if (outline > 0) {
      value.style.webkitTextStroke = `${outline}px ${tile.icon.outlineColor || "#111"}`;
      value.style.paintOrder = "stroke fill";
    } else {
      value.style.webkitTextStroke = "0";
    }
    tileEl.appendChild(value);
  }

  const handle = document.createElement("div");
  handle.className = "resize-handle export-hide";
  tileEl.appendChild(handle);

  header.addEventListener("pointerdown", (e) => beginDrag(e, tile.id, "move"));
  handle.addEventListener("pointerdown", (e) => beginDrag(e, tile.id, "resize"));
  deleteBtn.addEventListener("pointerdown", (e) => {
    // Prevent parent tile pointer handlers from re-rendering before click fires.
    e.stopPropagation();
  });
  deleteBtn.addEventListener("click", (e) => {
    e.preventDefault();
    e.stopPropagation();
    state.design.tiles = state.design.tiles.filter((t) => t.id !== tile.id);
    if (state.selectedTileId === tile.id) {
      state.selectedTileId = state.design.tiles[0]?.id || null;
    }
    render();
  });
  tileEl.addEventListener("pointerdown", (e) => {
    if (isTextTile(tile) && content.contains(e.target)) return;
    state.selectedTileId = tile.id;
    renderSelectionPanel();
    render();
  });

  return tileEl;
}

function isTextTile(tile) {
  return tile.type === "title" || tile.type === "effect" || tile.type === "flavor";
}

function applyTextStyle(element, tile) {
  const style = tile.style || {};
  const vAlignMap = {
    top: "flex-start",
    middle: "center",
    bottom: "flex-end",
  };
  element.style.display = "flex";
  element.style.flexDirection = "column";
  element.style.justifyContent = vAlignMap[style.verticalAlign] || "center";
  element.style.fontFamily =
    style.fontFamily || "'Manrope', 'Avenir Next', 'Helvetica Neue', Arial, sans-serif";
  element.style.fontSize = `${Number(style.fontSize || 16)}px`;
  element.style.lineHeight = String(style.lineHeight || 1.2);
  element.style.letterSpacing = `${Number(style.letterSpacing || 0)}px`;
  element.style.textAlign = style.textAlign || "left";
  element.style.color = style.color || "#1f2430";

  if (tile.type === "title") {
    const banner = tile.titleBanner || {};
    if (banner.enabled) {
      element.style.padding = `${mmToPx(banner.paddingMm ?? 0.5).toFixed(2)}px`;
      element.style.borderRadius = "4px";
      element.style.backgroundColor = "transparent";
      element.style.backgroundImage = "none";
      element.style.backgroundSize = "";
      element.style.backgroundRepeat = "";
      element.style.backgroundPosition = "";
      if (banner.mode === "color") {
        element.style.backgroundColor = banner.color || "#1f2430";
      } else if (banner.mode === "gradient") {
        element.style.backgroundImage = banner.gradient || "linear-gradient(90deg, #1f2937, #3b82f6)";
      } else if (banner.mode === "image" && banner.imageDataUrl) {
        element.style.backgroundImage = `url(${banner.imageDataUrl})`;
        element.style.backgroundSize = "100% 100%";
        element.style.backgroundRepeat = "no-repeat";
        element.style.backgroundPosition = "center";
      } else {
        element.style.backgroundColor = banner.color || "#1f2430";
      }
    } else {
      element.style.padding = "";
      element.style.borderRadius = "";
      element.style.backgroundColor = "";
      element.style.backgroundImage = "";
      element.style.backgroundSize = "";
      element.style.backgroundRepeat = "";
      element.style.backgroundPosition = "";
    }
  }
}

function applyTextFlowLayout(element, tile) {
  const rotation = normalizeRotation(tile.rotationDeg);
  const flowMode = tile.style?.textFlowMode === "vertical" ? "vertical" : "rotated";
  const contentW = Math.max(12, toPixels(tile.wIn) - 14);
  const contentH = Math.max(12, toPixels(tile.hIn) - 29);
  const quarterTurn = rotation === 90 || rotation === 270;

  element.style.position = "absolute";
  element.style.left = "50%";
  element.style.top = "50%";
  element.style.transformOrigin = "50% 50%";
  element.style.transform = `translate(-50%, -50%) rotate(${rotation}deg)`;
  element.style.overflow = "auto";
  element.style.whiteSpace = "pre-wrap";
  element.style.wordBreak = "break-word";

  if (flowMode === "vertical") {
    element.style.width = `${contentW}px`;
    element.style.height = `${contentH}px`;
    element.style.writingMode = "vertical-rl";
    element.style.textOrientation = "mixed";
    return;
  }

  if (quarterTurn) {
    // Swap layout box so line wrapping follows the Y-axis at 90/270.
    element.style.width = `${contentH}px`;
    element.style.height = `${contentW}px`;
  } else {
    element.style.width = `${contentW}px`;
    element.style.height = `${contentH}px`;
  }
  element.style.writingMode = "horizontal-tb";
  element.style.textOrientation = "";
}

function getEditableForSelectedTile() {
  const tile = getSelectedTile();
  if (!tile || !isTextTile(tile)) return null;
  const node = els.cardStage.querySelector(`[data-tile-content-id="${tile.id}"]`);
  return node || null;
}

function render() {
  updateStageGeometry();
  els.cardStage.innerHTML = "";
  state.design.tiles.forEach((tile) => {
    constrainTile(tile);
    if (tile.hidden) return;
    const node = renderTile(tile);
    els.cardStage.appendChild(node);
  });
  renderSelectionPanel();
  renderLayersPanel();
  if (state.tutorial.active) {
    renderTutorialGuide();
  }
  persistDraft();
}

function renderSelectionPanel() {
  const tile = getSelectedTile();
  els.tileControls.hidden = !tile;
  els.floatingPalette.hidden = !tile;
  if (!tile) return;

  els.tileNameInput.value = tile.label || "";
  els.tileTransparentBgInput.checked = !!tile.transparentBg;
  els.tileShowOutlineInput.checked = tile.showOutline !== false;
  els.tileXInput.value = tile.xIn.toFixed(2);
  els.tileYInput.value = tile.yIn.toFixed(2);
  els.tileWInput.value = tile.wIn.toFixed(2);
  els.tileHInput.value = tile.hIn.toFixed(2);
  els.tileRotationSelect.value = String(normalizeRotation(tile.rotationDeg));

  const style = tile.style || {};
  els.fontFamilySelect.value =
    style.fontFamily || "'Manrope', 'Avenir Next', 'Helvetica Neue', Arial, sans-serif";
  els.fontSizeInput.value = style.fontSize || 16;
  els.lineHeightInput.value = style.lineHeight || 1.2;
  els.textFlowModeSelect.value = style.textFlowMode || "rotated";
  els.textVAlignSelect.value = style.verticalAlign || "middle";
  els.letterSpacingInput.value = style.letterSpacing || 0;
  els.textColorInput.value = style.color || "#1f2430";
  els.bgColorInput.value = style.background || "#ffffff";

  const isText = isTextTile(tile);
  const isImage = tile.type === "main-image" || tile.type === "card-background";
  const isIcon = tile.type === "icon";
  els.textControlsWrap.hidden = !isText;
  els.titleControls.hidden = tile.type !== "title";
  els.imageControls.hidden = !(isImage || isIcon);
  els.iconControls.hidden = !isIcon;

  if (tile.type === "title") {
    const banner = tile.titleBanner || {};
    els.titleBannerEnabled.checked = !!banner.enabled;
    els.titleBannerMode.value = banner.mode || "color";
    els.titleBannerColor.value = banner.color || "#1f2430";
    els.titleBannerGradient.value = banner.gradient || "linear-gradient(90deg, #1f2937, #3b82f6)";
    els.titleBannerPaddingMm.value = Number(banner.paddingMm ?? 0.5);
    updateTitleBannerControls(tile);
  }
  if (tile.type === "icon") {
    tile.icon.position = "overlay";
    els.iconValueInput.value = tile.icon.value || "";
    els.iconValueColor.value = tile.icon.valueColor || "#ffffff";
    els.iconValueSize.value = tile.icon.valueSize || 60;
    els.iconOutlineColor.value = tile.icon.outlineColor || "#111111";
    els.iconOutlineWidth.value = tile.icon.outlineWidth || 2;
    els.iconOffsetX.value = tile.icon.offsetX || 0;
    els.iconOffsetY.value = tile.icon.offsetY || 0;
  }
  if (isImage || isIcon) {
    tile.imageFill = {
      ...getDefaultImageFill(tile.type),
      ...(tile.imageFill || {}),
    };
    els.imageFillMode.value = tile.imageFill.mode || "image";
    els.imageFillColor.value = tile.imageFill.color || "#f8f8f8";
    els.imageFillGradient.value =
      tile.imageFill.gradient || "linear-gradient(180deg, #ffffff, #e8f2ff)";
    updateImageFillControls(tile);
  } else {
    updateImageFillControls(null);
  }

  positionFloatingPalette(tile.id);
}

function updateTitleBannerControls(tile) {
  const isTitle = tile && tile.type === "title";
  if (!isTitle) {
    els.titleBannerColorRow.hidden = true;
    els.titleBannerGradientRow.hidden = true;
    els.titleBannerImageRow.hidden = true;
    return;
  }
  const banner = tile.titleBanner || {};
  const enabled = !!banner.enabled;
  const mode = banner.mode || "color";
  els.titleBannerMode.disabled = !enabled;
  els.titleBannerPaddingMm.disabled = !enabled;
  els.titleBannerColorRow.hidden = !enabled || mode !== "color";
  els.titleBannerGradientRow.hidden = !enabled || mode !== "gradient";
  els.titleBannerImageRow.hidden = !enabled || mode !== "image";
}

function updateImageFillControls(tile) {
  const isImageTile = tile && (tile.type === "main-image" || tile.type === "card-background" || tile.type === "icon");
  if (!isImageTile) {
    els.imageFillColorRow.hidden = true;
    els.imageFillGradientRow.hidden = true;
    els.uploadImageBtn.disabled = false;
    return;
  }
  const fill = {
    ...getDefaultImageFill(tile.type),
    ...(tile.imageFill || {}),
  };
  const mode = fill.mode || "image";
  els.imageFillColorRow.hidden = mode !== "color";
  els.imageFillGradientRow.hidden = mode !== "gradient";
  els.uploadImageBtn.disabled = false;
}

function moveLayer(tileId, direction) {
  const index = state.design.tiles.findIndex((t) => t.id === tileId);
  if (index < 0) return;
  const target = index + direction;
  if (target < 0 || target >= state.design.tiles.length) return;
  const tmp = state.design.tiles[index];
  state.design.tiles[index] = state.design.tiles[target];
  state.design.tiles[target] = tmp;
}

function renderLayersPanel() {
  if (!els.layersList) return;
  els.layersList.innerHTML = "";
  for (let idx = state.design.tiles.length - 1; idx >= 0; idx -= 1) {
    const tile = state.design.tiles[idx];
    const li = document.createElement("li");
    li.className = "layer-item";
    if (tile.id === state.selectedTileId) li.classList.add("active");

    const nameBtn = document.createElement("button");
    nameBtn.type = "button";
    nameBtn.className = "layer-name";
    nameBtn.textContent = tile.label || tile.type;
    nameBtn.addEventListener("click", () => {
      state.selectedTileId = tile.id;
      render();
    });

    const controls = document.createElement("div");
    controls.className = "layer-controls";

    const upBtn = document.createElement("button");
    upBtn.type = "button";
    upBtn.className = "layer-move";
    upBtn.textContent = "↑";
    upBtn.title = "Move up";
    upBtn.disabled = idx === state.design.tiles.length - 1;
    upBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      moveLayer(tile.id, 1);
      render();
    });

    const downBtn = document.createElement("button");
    downBtn.type = "button";
    downBtn.className = "layer-move";
    downBtn.textContent = "↓";
    downBtn.title = "Move down";
    downBtn.disabled = idx === 0;
    downBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      moveLayer(tile.id, -1);
      render();
    });

    const visLabel = document.createElement("label");
    visLabel.className = "layer-check";
    visLabel.title = "Toggle visibility";
    const visCheck = document.createElement("input");
    visCheck.type = "checkbox";
    visCheck.checked = !tile.hidden;
    visCheck.addEventListener("change", () => {
      tile.hidden = !visCheck.checked;
      render();
    });
    const visText = document.createElement("span");
    visText.textContent = "Show";
    visLabel.append(visCheck, visText);

    const bgLabel = document.createElement("label");
    bgLabel.className = "layer-check";
    bgLabel.title = "Transparent background";
    const bgCheck = document.createElement("input");
    bgCheck.type = "checkbox";
    bgCheck.checked = !!tile.transparentBg;
    bgCheck.addEventListener("change", () => {
      tile.transparentBg = bgCheck.checked;
      render();
    });
    const bgText = document.createElement("span");
    bgText.textContent = "Transparent";
    bgLabel.append(bgCheck, bgText);

    controls.append(upBtn, downBtn, visLabel, bgLabel);
    li.append(nameBtn, controls);
    els.layersList.appendChild(li);
  }
}

function positionFloatingPalette(tileId) {
  const tileEl = els.cardStage.querySelector(`[data-tile-id="${tileId}"]`);
  if (!tileEl || els.floatingPalette.hidden) return;
  const paletteRect = els.floatingPalette.getBoundingClientRect();
  const tileRect = tileEl.getBoundingClientRect();
  let top = tileRect.top - 6;
  const minTop = 72;
  const maxTop = window.innerHeight - paletteRect.height - 10;
  top = clamp(top, minTop, Math.max(minTop, maxTop));
  els.floatingPalette.style.top = `${Math.round(top)}px`;
}

function beginDrag(event, tileId, mode) {
  event.preventDefault();
  event.stopPropagation();
  const tile = state.design.tiles.find((t) => t.id === tileId);
  if (!tile) return;
  state.selectedTileId = tileId;
  state.dragAction = {
    tileId,
    mode,
    startX: event.clientX,
    startY: event.clientY,
    xIn: tile.xIn,
    yIn: tile.yIn,
    wIn: tile.wIn,
    hIn: tile.hIn,
  };
  window.addEventListener("pointermove", onPointerMove);
  window.addEventListener("pointerup", onPointerUp, { once: true });
}

function onPointerMove(event) {
  if (!state.dragAction) return;
  const action = state.dragAction;
  const tile = state.design.tiles.find((t) => t.id === action.tileId);
  if (!tile) return;

  const dxIn = toInches(event.clientX - action.startX);
  const dyIn = toInches(event.clientY - action.startY);
  if (action.mode === "move") {
    tile.xIn = snap(action.xIn + dxIn);
    tile.yIn = snap(action.yIn + dyIn);
  } else {
    tile.wIn = snap(action.wIn + dxIn);
    tile.hIn = snap(action.hIn + dyIn);
  }
  constrainTile(tile);
  render();
}

function onPointerUp() {
  window.removeEventListener("pointermove", onPointerMove);
  state.dragAction = null;
  persistDraft();
}

function setTilePlacement(tile, placement) {
  const size = getCardSize();
  const top = SAFE_BORDER_IN;
  const middle = size.h / 2 - tile.hIn / 2;
  const bottom = size.h - SAFE_BORDER_IN - tile.hIn;
  tile.titlePlacement = placement;
  if (placement === "middle") tile.yIn = middle;
  if (placement === "bottom") tile.yIn = bottom;
  if (placement === "top") tile.yIn = top;
  tile.yIn = snap(tile.yIn);
  constrainTile(tile);
}

function clampTileInputs(tile) {
  tile.xIn = snap(Number(els.tileXInput.value || tile.xIn));
  tile.yIn = snap(Number(els.tileYInput.value || tile.yIn));
  tile.wIn = snap(Number(els.tileWInput.value || tile.wIn));
  tile.hIn = snap(Number(els.tileHInput.value || tile.hIn));
  constrainTile(tile);
}

function ensureTitleContent(tile) {
  if (!tile || tile.type !== "title") return;
}

function bindEvents() {
  window.addEventListener("resize", render);

  const closeToolsMenu = () => {
    if (!els.toolsMenu || !els.toolsMenuBtn) return;
    els.toolsMenu.hidden = true;
    els.toolsMenuBtn.setAttribute("aria-expanded", "false");
  };

  els.toolsMenuBtn?.addEventListener("click", (event) => {
    event.stopPropagation();
    if (!els.toolsMenu || !els.toolsMenuBtn) return;
    const willOpen = els.toolsMenu.hidden;
    els.toolsMenu.hidden = !willOpen;
    els.toolsMenuBtn.setAttribute("aria-expanded", String(willOpen));
  });
  els.toolsMenu?.addEventListener("click", (event) => {
    event.stopPropagation();
  });
  document.addEventListener("click", closeToolsMenu);
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") closeToolsMenu();
  });

  els.themeToggleBtn?.addEventListener("click", () => {
    const current = document.body.getAttribute("data-theme") === "dark" ? "dark" : "light";
    const next = current === "dark" ? "light" : "dark";
    applyTheme(next);
    localStorage.setItem(THEME_KEY, next);
  });

  els.feedbackBtn?.addEventListener("click", () => {
    window.location.href = "mailto:help@pnpfinder.com?subject=Card%20Protyper%20Feedback";
  });

  els.openTutorialBtn?.addEventListener("click", () => {
    openTutorialPrompt(true);
  });
  els.previewCloseBtn?.addEventListener("click", () => {
    closePreviewModal();
  });
  els.previewModal?.addEventListener("click", (event) => {
    if (event.target === els.previewModal) closePreviewModal();
  });
  els.tutorialViewBtn?.addEventListener("click", () => {
    startTutorialGuide();
  });
  els.tutorialSkipBtn?.addEventListener("click", () => {
    closeTutorial(true);
  });
  els.tutorialPrevBtn?.addEventListener("click", () => {
    if (!state.tutorial.active) return;
    state.tutorial.index = Math.max(0, state.tutorial.index - 1);
    renderTutorialGuide();
  });
  els.tutorialNextBtn?.addEventListener("click", () => {
    if (!state.tutorial.active) return;
    const steps = getTutorialSteps();
    if (state.tutorial.index >= steps.length - 1) {
      closeTutorial(true);
      return;
    }
    state.tutorial.index += 1;
    renderTutorialGuide();
  });
  els.tutorialGuideSkipBtn?.addEventListener("click", () => {
    closeTutorial(true);
  });
  els.tutorialModal?.addEventListener("click", (event) => {
    if (event.target === els.tutorialModal) {
      closeTutorial(true);
    }
  });

  els.cardSizeSelect.addEventListener("change", () => {
    state.design.card.preset = els.cardSizeSelect.value;
    if (els.cardSizeSelect.value === "custom") {
      els.customSizeWrap.hidden = false;
      state.design.card.widthIn = Number(els.customWidth.value || 2.5);
      state.design.card.heightIn = Number(els.customHeight.value || 3.5);
    } else {
      els.customSizeWrap.hidden = true;
      const preset = cardSizes[els.cardSizeSelect.value] || cardSizes.poker;
      state.design.card.widthIn = preset.w;
      state.design.card.heightIn = preset.h;
    }
    state.design.tiles.forEach(constrainTile);
    render();
  });

  [els.customWidth, els.customHeight].forEach((input) => {
    input.addEventListener("input", () => {
      if (els.cardSizeSelect.value !== "custom") return;
      state.design.card.widthIn = clamp(Number(els.customWidth.value || 2.5), 1, 10);
      state.design.card.heightIn = clamp(Number(els.customHeight.value || 3.5), 1, 14);
      state.design.tiles.forEach(constrainTile);
      render();
    });
  });

  els.snapToggle.addEventListener("change", () => {
    state.design.snapEnabled = els.snapToggle.checked;
    render();
  });

  els.gridStepInput.addEventListener("input", () => {
    state.design.gridStepIn = clamp(Number(els.gridStepInput.value || 0.005), 0.005, 0.5);
  });

  els.showSafeBorder.addEventListener("change", () => {
    state.design.showSafeBorder = els.showSafeBorder.checked;
    render();
  });

  els.tilePalette.querySelectorAll("button").forEach((btn) => {
    btn.addEventListener("dragstart", (event) => {
      state.dragTileType = btn.dataset.tileType;
      event.dataTransfer.effectAllowed = "copy";
      event.dataTransfer.setData("text/plain", state.dragTileType);
    });
  });

  els.cardStage.addEventListener("dragover", (event) => {
    event.preventDefault();
  });

  els.cardStage.addEventListener("pointerdown", (event) => {
    if (event.target === els.cardStage) {
      state.selectedTileId = null;
      render();
    }
  });

  els.cardStage.addEventListener("drop", (event) => {
    event.preventDefault();
    const type = event.dataTransfer.getData("text/plain") || state.dragTileType;
    if (!type) return;
    const tile = createTile(type);
    if (type === "card-background") {
      state.design.tiles.unshift(tile);
    } else {
      state.design.tiles.push(tile);
    }
    state.selectedTileId = tile.id;
    render();
    setStatus(`${tile.label} added at center.`);
  });

  els.tileNameInput.addEventListener("input", () => {
    const tile = getSelectedTile();
    if (!tile) return;
    tile.label = els.tileNameInput.value;
    render();
  });

  els.tileTransparentBgInput.addEventListener("change", () => {
    const tile = getSelectedTile();
    if (!tile) return;
    tile.transparentBg = els.tileTransparentBgInput.checked;
    render();
  });

  els.tileShowOutlineInput.addEventListener("change", () => {
    const tile = getSelectedTile();
    if (!tile) return;
    tile.showOutline = els.tileShowOutlineInput.checked;
    render();
  });

  els.titleBannerEnabled.addEventListener("change", () => {
    const tile = getSelectedTile();
    if (!tile || tile.type !== "title") return;
    tile.titleBanner = {
      enabled: els.titleBannerEnabled.checked,
      mode: "color",
      color: "#1f2430",
      gradient: "linear-gradient(90deg, #1f2937, #3b82f6)",
      imageDataUrl: "",
      paddingMm: 0.5,
      ...(tile.titleBanner || {}),
    };
    tile.titleBanner.enabled = els.titleBannerEnabled.checked;
    updateTitleBannerControls(tile);
    render();
  });

  els.titleBannerMode.addEventListener("change", () => {
    const tile = getSelectedTile();
    if (!tile || tile.type !== "title") return;
    tile.titleBanner = {
      enabled: false,
      mode: "color",
      color: "#1f2430",
      gradient: "linear-gradient(90deg, #1f2937, #3b82f6)",
      imageDataUrl: "",
      paddingMm: 0.5,
      ...(tile.titleBanner || {}),
    };
    tile.titleBanner.mode = els.titleBannerMode.value;
    updateTitleBannerControls(tile);
    render();
  });

  els.titleBannerColor.addEventListener("input", () => {
    const tile = getSelectedTile();
    if (!tile || tile.type !== "title") return;
    tile.titleBanner = {
      enabled: false,
      mode: "color",
      color: "#1f2430",
      gradient: "linear-gradient(90deg, #1f2937, #3b82f6)",
      imageDataUrl: "",
      paddingMm: 0.5,
      ...(tile.titleBanner || {}),
    };
    tile.titleBanner.color = els.titleBannerColor.value;
    render();
  });

  els.titleBannerGradient.addEventListener("change", () => {
    const tile = getSelectedTile();
    if (!tile || tile.type !== "title") return;
    tile.titleBanner = {
      enabled: false,
      mode: "gradient",
      color: "#1f2430",
      gradient: "linear-gradient(90deg, #1f2937, #3b82f6)",
      imageDataUrl: "",
      paddingMm: 0.5,
      ...(tile.titleBanner || {}),
    };
    tile.titleBanner.gradient = els.titleBannerGradient.value;
    render();
  });

  els.titleBannerPaddingMm.addEventListener("input", () => {
    const tile = getSelectedTile();
    if (!tile || tile.type !== "title") return;
    tile.titleBanner = {
      enabled: false,
      mode: "color",
      color: "#1f2430",
      gradient: "linear-gradient(90deg, #1f2937, #3b82f6)",
      imageDataUrl: "",
      paddingMm: 0.5,
      ...(tile.titleBanner || {}),
    };
    tile.titleBanner.paddingMm = clamp(Number(els.titleBannerPaddingMm.value || 0.5), 0, 10);
    render();
  });

  els.titleBannerUploadBtn.addEventListener("click", () => {
    els.titleBannerFileInput.click();
  });

  els.titleBannerFileInput.addEventListener("change", async () => {
    const tile = getSelectedTile();
    const file = els.titleBannerFileInput.files?.[0];
    els.titleBannerFileInput.value = "";
    if (!tile || tile.type !== "title" || !file) return;
    try {
      const dataUrl = await validateAndReadImage(file);
      tile.titleBanner = {
        enabled: true,
        mode: "image",
        color: "#1f2430",
        gradient: "linear-gradient(90deg, #1f2937, #3b82f6)",
        imageDataUrl: "",
        paddingMm: 0.5,
        ...(tile.titleBanner || {}),
      };
      tile.titleBanner.enabled = true;
      tile.titleBanner.mode = "image";
      tile.titleBanner.imageDataUrl = dataUrl;
      updateTitleBannerControls(tile);
      render();
      setStatus("Title banner image uploaded.");
    } catch (error) {
      setStatus(error.message);
    }
  });

  [els.tileXInput, els.tileYInput, els.tileWInput, els.tileHInput].forEach((input) => {
    input.addEventListener("input", () => {
      const tile = getSelectedTile();
      if (!tile) return;
      clampTileInputs(tile);
      render();
    });
  });

  els.tileRotationSelect.addEventListener("change", () => {
    const tile = getSelectedTile();
    if (!tile) return;
    tile.rotationDeg = normalizeRotation(els.tileRotationSelect.value);
    render();
  });

  els.deleteTileBtn.addEventListener("click", () => {
    if (!state.selectedTileId) return;
    state.design.tiles = state.design.tiles.filter((t) => t.id !== state.selectedTileId);
    state.selectedTileId = null;
    render();
  });

  const styleInputs = [
    els.fontFamilySelect,
    els.fontSizeInput,
    els.lineHeightInput,
    els.textFlowModeSelect,
    els.textVAlignSelect,
    els.letterSpacingInput,
    els.textColorInput,
    els.bgColorInput,
  ];
  styleInputs.forEach((input) => {
    input.addEventListener("input", () => {
      const tile = getSelectedTile();
      if (!tile || !isTextTile(tile)) return;
      if (input === els.bgColorInput) {
        tile.transparentBg = false;
      }
      tile.style = {
        ...tile.style,
        fontFamily: els.fontFamilySelect.value,
        fontSize: Number(els.fontSizeInput.value || 16),
        lineHeight: Number(els.lineHeightInput.value || 1.2),
        textFlowMode: els.textFlowModeSelect.value || "rotated",
        verticalAlign: els.textVAlignSelect.value || "middle",
        letterSpacing: Number(els.letterSpacingInput.value || 0),
        color: els.textColorInput.value,
        background: els.bgColorInput.value,
      };
      render();
    });
  });

  els.toolbarButtons.forEach((btn) => {
    btn.addEventListener("click", () => {
      const cmd = btn.dataset.cmd;
      const tile = getSelectedTile();
      if (!tile || !isTextTile(tile)) return;

      if (cmd === "justifyLeft" || cmd === "justifyCenter" || cmd === "justifyRight" || cmd === "justifyFull") {
        const map = {
          justifyLeft: "left",
          justifyCenter: "center",
          justifyRight: "right",
          justifyFull: "justify",
        };
        tile.style = {
          ...tile.style,
          textAlign: map[cmd],
        };
        render();
        return;
      }

      const editable = state.liveEditable || getEditableForSelectedTile();
      if (!editable) return;
      editable.focus();
      document.execCommand(cmd, false, null);
      tile.textHtml = editable.innerHTML;
      persistDraft();
    });
  });

  els.uploadImageBtn.addEventListener("click", () => {
    els.imageFileInput.click();
  });

  els.imageFillMode.addEventListener("change", () => {
    const tile = getSelectedTile();
    if (!tile || (tile.type !== "main-image" && tile.type !== "card-background" && tile.type !== "icon")) return;
    tile.imageFill = {
      ...getDefaultImageFill(tile.type),
      ...(tile.imageFill || {}),
      mode: els.imageFillMode.value,
    };
    if (tile.imageFill.mode === "color" || tile.imageFill.mode === "gradient") {
      tile.transparentBg = false;
    }
    updateImageFillControls(tile);
    render();
  });

  els.imageFillColor.addEventListener("input", () => {
    const tile = getSelectedTile();
    if (!tile || (tile.type !== "main-image" && tile.type !== "card-background" && tile.type !== "icon")) return;
    tile.imageFill = {
      ...getDefaultImageFill(tile.type),
      ...(tile.imageFill || {}),
      color: els.imageFillColor.value,
    };
    render();
  });

  els.imageFillGradient.addEventListener("change", () => {
    const tile = getSelectedTile();
    if (!tile || (tile.type !== "main-image" && tile.type !== "card-background" && tile.type !== "icon")) return;
    tile.imageFill = {
      ...getDefaultImageFill(tile.type),
      ...(tile.imageFill || {}),
      gradient: els.imageFillGradient.value,
    };
    render();
  });

  els.imageFileInput.addEventListener("change", async () => {
    const tile = getSelectedTile();
    const file = els.imageFileInput.files?.[0];
    els.imageFileInput.value = "";
    if (!tile || !file) return;
    try {
      const dataUrl = await validateAndReadImage(file);
      tile.imageDataUrl = dataUrl;
      tile.imageFill = {
        ...getDefaultImageFill(tile.type),
        ...(tile.imageFill || {}),
        mode: "image",
      };
      updateImageFillControls(tile);
      render();
      setStatus("Image uploaded.");
    } catch (error) {
      setStatus(error.message);
    }
  });

  els.iconValueInput.addEventListener("input", () => {
    const tile = getSelectedTile();
    if (!tile || tile.type !== "icon") return;
    tile.icon.value = els.iconValueInput.value;
    render();
  });

  [
    els.iconValueColor,
    els.iconValueSize,
    els.iconOutlineColor,
    els.iconOutlineWidth,
    els.iconOffsetX,
    els.iconOffsetY,
  ].forEach((input) => {
    input.addEventListener("input", () => {
      const tile = getSelectedTile();
      if (!tile || tile.type !== "icon") return;
      tile.icon.position = "overlay";
      tile.icon.valueColor = els.iconValueColor.value;
      tile.icon.valueSize = Number(els.iconValueSize.value || 60);
      tile.icon.outlineColor = els.iconOutlineColor.value;
      tile.icon.outlineWidth = Number(els.iconOutlineWidth.value || 2);
      tile.icon.offsetX = Number(els.iconOffsetX.value || 0);
      tile.icon.offsetY = Number(els.iconOffsetY.value || 0);
      render();
    });
  });

  els.projectName.addEventListener("input", () => {
    state.design.name = els.projectName.value || "Untitled Design";
    persistDraft();
  });

  els.saveProjectBtn.addEventListener("click", () => {
    downloadProjectJson();
  });

  els.loadProjectBtn.addEventListener("click", () => {
    els.projectFileInput.click();
  });

  els.projectFileInput.addEventListener("change", async () => {
    const file = els.projectFileInput.files?.[0];
    els.projectFileInput.value = "";
    if (!file) return;
    try {
      const raw = await file.text();
      const parsed = JSON.parse(raw);
      loadDesignFromData(parsed, true);
      setStatus("Project JSON imported.");
    } catch (error) {
      setStatus("Invalid JSON file.");
    }
  });

  els.resetProjectBtn.addEventListener("click", () => {
    const ok = window.confirm("Start over with a blank project? Unsaved changes will be lost.");
    if (!ok) return;
    resetProject({ preserveSettings: true });
  });

  els.exportPngBtn.addEventListener("click", async () => {
    try {
      const png = await captureCardPngDataUrl();
      const a = document.createElement("a");
      a.href = png;
      a.download = `${safeFile(state.design.name || "card")}.png`;
      a.click();
      setStatus("PNG exported.");
    } catch (error) {
      console.error(error);
      setStatus(`PNG export failed: ${error?.message || "Unknown error."}`);
    }
  });

  els.previewCardBtn?.addEventListener("click", async () => {
    try {
      const png = await captureCardPngDataUrl();
      if (els.previewImage) {
        els.previewImage.src = png;
      }
      if (els.previewDownloadBtn) {
        els.previewDownloadBtn.onclick = () => {
          const a = document.createElement("a");
          a.href = png;
          a.download = `${safeFile(state.design.name || "card")}.png`;
          a.click();
        };
      }
      if (els.previewModal) {
        els.previewModal.hidden = false;
      }
      setStatus("Preview ready.");
    } catch (error) {
      console.error(error);
      setStatus(`Preview failed: ${error?.message || "Unknown error."}`);
    }
  });

  els.exportPdfBtn.addEventListener("click", async () => {
    try {
      await exportPdf();
      setStatus("PDF exported.");
    } catch (error) {
      console.error(error);
      setStatus("PDF export failed.");
    }
  });
}

async function validateAndReadImage(file) {
  if (!["image/png", "image/jpeg"].includes(file.type)) {
    throw new Error("Only PNG and JPG images are allowed.");
  }
  if (file.size > MAX_IMAGE_BYTES) {
    throw new Error("Image too large. Max file size is 2.5 MB.");
  }
  const dataUrl = await readFileAsDataUrl(file);
  const img = await loadImage(dataUrl);
  if (img.width > MAX_IMAGE_DIMENSION || img.height > MAX_IMAGE_DIMENSION) {
    throw new Error("Image dimensions exceed 3000x3000 limit.");
  }
  return dataUrl;
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

function loadImage(src) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = reject;
    img.src = src;
  });
}

function downloadProjectJson() {
  const blob = new Blob([JSON.stringify(cleanDesign(state.design), null, 2)], {
    type: "application/json",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${safeFile(state.design.name || "design")}.json`;
  a.click();
  URL.revokeObjectURL(url);
  setStatus("Project JSON downloaded.");
}

function persistDraft() {
  state.design.updatedAt = new Date().toISOString();
}

function loadDraft() {
  return false;
}

function cleanDesign(design) {
  return JSON.parse(JSON.stringify(design));
}

function loadDesignFromData(data, resetProjectId) {
  const preset = data.card?.preset || "poker";
  const presetSize = cardSizes[preset] || cardSizes.poker;
  state.design = {
    name: data.name || "Imported Design",
    card: {
      preset,
      widthIn: Number(data.card?.widthIn || presetSize.w),
      heightIn: Number(data.card?.heightIn || presetSize.h),
    },
    snapEnabled: data.snapEnabled !== false,
    gridStepIn: clamp(Number(data.gridStepIn || 0.005), 0.005, 0.5),
    showSafeBorder: data.showSafeBorder !== false,
    tiles: (data.tiles || []).map((tile) => ({
      ...createTile(tile.type || "effect"),
      ...tile,
      rotationDeg: normalizeRotation(tile.rotationDeg ?? 0),
      hidden: tile.hidden === true,
      showOutline: tile.showOutline !== false,
      transparentBg:
        tile.transparentBg != null
          ? tile.transparentBg
          : tile.type === "card-background" || tile.type === "main-image" || tile.type === "icon",
      style: {
        ...getDefaultTextStyle(tile.type || "effect"),
        ...(tile.style || {}),
      },
      imageFill: {
        ...getDefaultImageFill(tile.type || "effect"),
        ...(tile.imageFill || {}),
      },
      titleBanner: {
        enabled: false,
        mode: "color",
        color: "#1f2430",
        gradient: "linear-gradient(90deg, #1f2937, #3b82f6)",
        imageDataUrl: "",
        paddingMm: 0.5,
        ...(tile.titleBanner || {}),
      },
      icon: {
        value: "1",
        position: "overlay",
        valueColor: "#ffffff",
        valueSize: 60,
        outlineColor: "#111111",
        outlineWidth: 2,
        offsetX: 0,
        offsetY: 0,
        ...(tile.icon || {}),
      },
    })),
    updatedAt: data.updatedAt || new Date().toISOString(),
  };
  state.selectedTileId = state.design.tiles[0]?.id || null;
  if (resetProjectId) state.projectId = null;
  syncInputsFromState();
  render();
}

function resetProject(options = {}) {
  const { preserveSettings = false } = options;
  state.projectId = null;
  state.selectedTileId = null;
  state.liveEditable = null;
  const settings = preserveSettings
    ? {
        card: { ...state.design.card },
        snapEnabled: state.design.snapEnabled,
        gridStepIn: state.design.gridStepIn,
        showSafeBorder: state.design.showSafeBorder,
      }
    : {};
  state.design = buildDefaultDesign(settings);
  const t1 = createTile("title");
  const t2 = createTile("main-image");
  const t3 = createTile("effect");
  state.design.tiles.push(t1, t2, t3);
  state.selectedTileId = t1.id;
  syncInputsFromState();
  render();
  setStatus("Started a new project.");
}

function syncInputsFromState() {
  els.projectName.value = state.design.name || "";
  els.snapToggle.checked = state.design.snapEnabled;
  els.gridStepInput.value = state.design.gridStepIn;
  els.showSafeBorder.checked = state.design.showSafeBorder;

  const preset = state.design.card.preset;
  els.cardSizeSelect.value = preset;
  if (preset === "custom") {
    els.customSizeWrap.hidden = false;
    els.customWidth.value = state.design.card.widthIn;
    els.customHeight.value = state.design.card.heightIn;
  } else {
    els.customSizeWrap.hidden = true;
  }
}

function safeFile(name) {
  return String(name).trim().replace(/[^a-z0-9-_]+/gi, "-").replace(/-+/g, "-").replace(/^-|-$/g, "");
}

async function captureCardPngDataUrl() {
  if (typeof html2canvas !== "function") {
    throw new Error("Export dependency failed to load. Refresh the page and try again.");
  }

  const selected = state.selectedTileId;
  const wasShowingSafeBorder = state.design.showSafeBorder;
  const widthPx = Math.max(1, Number(els.cardStage.clientWidth) || 0);
  const cardWidthPx = Math.max(1, getCardSize().w * getPxPerIn());
  const targetScale = Math.max(2, Math.min(4, Math.floor((DPI_EXPORT * getCardSize().w) / widthPx)));
  const fallbackScale = Math.max(2, Math.min(4, Math.round(cardWidthPx / widthPx)));
  state.selectedTileId = null;
  state.design.showSafeBorder = false;
  els.cardStage.classList.add("export-mode");
  render();
  try {
    const baseOptions = {
      backgroundColor: null,
      ignoreElements: (element) => element.classList?.contains("export-hide"),
      useCORS: true,
      logging: false,
      onclone: (doc) => {
        const stage = doc.querySelector("#cardStage");
        if (!stage) return;
        // Avoid CSS functions that can break some html2canvas parsers.
        stage.style.background = "#ffffff";
        stage.style.backgroundImage = "none";
        stage.style.boxShadow = "none";
        stage.style.borderColor = "rgba(36, 55, 88, 0.24)";
        stage.querySelectorAll(".tile").forEach((tileEl) => {
          tileEl.style.backgroundImage = "none";
        });
      },
    };

    try {
      const canvas = await html2canvas(els.cardStage, {
        ...baseOptions,
        scale: targetScale,
      });
      return canvas.toDataURL("image/png");
    } catch (firstError) {
      const canvas = await html2canvas(els.cardStage, {
        ...baseOptions,
        scale: fallbackScale,
        foreignObjectRendering: true,
      });
      return canvas.toDataURL("image/png");
    }
  } finally {
    els.cardStage.classList.remove("export-mode");
    state.selectedTileId = selected;
    state.design.showSafeBorder = wasShowingSafeBorder;
    render();
  }
}

function getLayoutConfig(layout) {
  if (layout === "grid3x3") return { cols: 3, rows: 3, centerGutterIn: 0, cardRotate: 0 };
  if (layout === "grid2x3") return { cols: 3, rows: 2, centerGutterIn: 0, cardRotate: 0 };
  return { cols: 2, rows: 4, centerGutterIn: 0.25, cardRotate: 90 };
}

function getPageSizePoints(pageKey) {
  const p = pageSizes[pageKey] || pageSizes.letter;
  return { w: inchesToPoints(p.w), h: inchesToPoints(p.h) };
}

function getCardSizePointsForLayout(layoutKey) {
  const card = getCardSize();
  const base = { w: inchesToPoints(card.w), h: inchesToPoints(card.h) };
  if (layoutKey === "gutterfold") return { w: base.h, h: base.w };
  return base;
}

function getPositions(layoutConfig, pageW, pageH, cardW, cardH, safeMarginPt = inchesToPoints(0.25)) {
  const totalW =
    layoutConfig.cols * cardW + (layoutConfig.cols - 1) * 0 + (layoutConfig.cols === 2 ? inchesToPoints(layoutConfig.centerGutterIn || 0) : 0);
  const totalH = layoutConfig.rows * cardH;
  const marginX = Math.max((pageW - totalW) / 2, safeMarginPt);
  const marginY = Math.max((pageH - totalH) / 2, safeMarginPt);
  const centerGutter = inchesToPoints(layoutConfig.centerGutterIn || 0);
  const out = [];
  for (let row = 0; row < layoutConfig.rows; row += 1) {
    for (let col = 0; col < layoutConfig.cols; col += 1) {
      let x = marginX + col * cardW;
      if (layoutConfig.cols === 2 && col === 1) x += centerGutter;
      const y = pageH - marginY - cardH - row * cardH;
      out.push({ x, y, width: cardW, height: cardH });
    }
  }
  return out;
}

function drawCutGuides(page, box) {
  const l = 7;
  const s = 0.8;
  const col = PDFLib.rgb(0.25, 0.25, 0.25);
  const x1 = box.x;
  const y1 = box.y;
  const x2 = box.x + box.width;
  const y2 = box.y + box.height;

  page.drawLine({ start: { x: x1 - l, y: y1 }, end: { x: x1, y: y1 }, thickness: s, color: col });
  page.drawLine({ start: { x: x1, y: y1 - l }, end: { x: x1, y: y1 }, thickness: s, color: col });
  page.drawLine({ start: { x: x2, y: y1 - l }, end: { x: x2, y: y1 }, thickness: s, color: col });
  page.drawLine({ start: { x: x2, y: y1 }, end: { x: x2 + l, y: y1 }, thickness: s, color: col });
  page.drawLine({ start: { x: x1 - l, y: y2 }, end: { x: x1, y: y2 }, thickness: s, color: col });
  page.drawLine({ start: { x: x1, y: y2 }, end: { x: x1, y: y2 + l }, thickness: s, color: col });
  page.drawLine({ start: { x: x2, y: y2 }, end: { x: x2 + l, y: y2 }, thickness: s, color: col });
  page.drawLine({ start: { x: x2, y: y2 }, end: { x: x2, y: y2 + l }, thickness: s, color: col });
}

async function exportPdf() {
  const pageKey = els.pdfPageSize.value;
  const layoutKey = els.pdfLayout.value;
  const copies = clamp(Number(els.pdfCopies.value || 1), 1, 300);
  const layout = getLayoutConfig(layoutKey);
  const pageSize = getPageSizePoints(pageKey);
  const cardSizePts = getCardSizePointsForLayout(layoutKey);
  const positions = getPositions(layout, pageSize.w, pageSize.h, cardSizePts.w, cardSizePts.h);
  const cardsPerPage = positions.length;

  if (!positions.length) throw new Error("Invalid PDF layout.");

  const pngData = await captureCardPngDataUrl();
  const imageData =
    layoutKey === "gutterfold" ? await rotateDataUrlClockwise(pngData) : pngData;
  const pngBytes = await fetch(imageData).then((r) => r.arrayBuffer());
  const pdf = await PDFLib.PDFDocument.create();
  const image = await pdf.embedPng(pngBytes);
  const totalPages = Math.ceil(copies / cardsPerPage);

  for (let pageIndex = 0; pageIndex < totalPages; pageIndex += 1) {
    const page = pdf.addPage([pageSize.w, pageSize.h]);
    const pageBoxes = [];
    for (let i = 0; i < cardsPerPage; i += 1) {
      const global = pageIndex * cardsPerPage + i;
      if (global >= copies) break;
      const box = positions[i];
      page.drawImage(image, {
        x: box.x,
        y: box.y,
        width: box.width,
        height: box.height,
      });
      pageBoxes.push(box);
    }
    pageBoxes.forEach((box) => {
      drawCutGuides(page, box);
    });
  }

  const bytes = await pdf.save();
  const blob = new Blob([bytes], { type: "application/pdf" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${safeFile(state.design.name || "card")}-${layoutKey}-${pageKey}.pdf`;
  a.click();
  URL.revokeObjectURL(url);
}

async function rotateDataUrlClockwise(dataUrl) {
  const img = await loadImage(dataUrl);
  const canvas = document.createElement("canvas");
  canvas.width = img.height;
  canvas.height = img.width;
  const ctx = canvas.getContext("2d");
  ctx.translate(canvas.width / 2, canvas.height / 2);
  ctx.rotate(Math.PI / 2);
  ctx.drawImage(img, -img.width / 2, -img.height / 2);
  return canvas.toDataURL("image/png");
}

function bootstrap() {
  sortFontFamilyOptions();
  bindEvents();
  initTheme();
  applyTooltips();
  const loadedDraft = loadDraft();
  if (!loadedDraft) {
    resetProject();
    openTutorialPrompt(false);
    return;
  }
  syncInputsFromState();
  render();
  openTutorialPrompt(false);
}

bootstrap();
