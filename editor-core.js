(function (root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  root.PrototyperCore = api;
})(typeof globalThis !== "undefined" ? globalThis : window, function () {
  "use strict";

  const SAFE_BORDER_IN = 0.125;
  const LEGACY_PX_PER_IN = 200;
  const CARD_SIZES = {
    poker: { widthIn: 2.5, heightIn: 3.5 },
    tarot: { widthIn: 2.75, heightIn: 4.75 },
    mini: { widthIn: 1.75, heightIn: 2.5 },
  };

  const uid = () => `${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
  const clone = (value) =>
    typeof structuredClone === "function"
      ? structuredClone(value)
      : JSON.parse(JSON.stringify(value));
  const clamp = (value, min, max) => Math.max(min, Math.min(max, value));
  const snap = (value, step, enabled = true) =>
    enabled ? Math.round(value / Math.max(0.005, step || 0.05)) * Math.max(0.005, step || 0.05) : value;
  const normalizeRotation = (value) => {
    const number = Number(value);
    if (!Number.isFinite(number)) return 0;
    return ((number % 360) + 360) % 360;
  };
  const legacyPxToPt = (pixels) => Number(pixels || 16) * 72 / LEGACY_PX_PER_IN;

  function defaultStyle(type) {
    if (type === "title") {
      return {
        preset: "clean-band", fontFamily: "'Manrope', sans-serif", fontSizePt: 18,
        fontWeight: 700, lineHeight: 1.08, textAlign: "center", verticalAlign: "middle",
        letterSpacing: 0.4, color: "#f8fbff", background: "#17233d",
      };
    }
    if (type === "flavor") {
      return {
        preset: "transparent", fontFamily: "'Cormorant Garamond', serif", fontSizePt: 9,
        fontWeight: 500, lineHeight: 1.22, textAlign: "left", verticalAlign: "middle",
        letterSpacing: 0, color: "#4b5565", background: "#ffffff",
      };
    }
    return {
      preset: "paper", fontFamily: "'Manrope', sans-serif", fontSizePt: 9.5,
      fontWeight: 500, lineHeight: 1.3, textAlign: "left", verticalAlign: "top",
      letterSpacing: 0, color: "#202839", background: "#f8fafc",
    };
  }

  function baseTile(type, card, overrides = {}) {
    const width = card.widthIn;
    const height = card.heightIn;
    const presets = {
      "card-background": { xIn: 0, yIn: 0, wIn: width, hIn: height },
      title: { xIn: 0.18, yIn: 0.17, wIn: width - 0.36, hIn: 0.42 },
      "main-image": { xIn: 0.18, yIn: 0.68, wIn: width - 0.36, hIn: height * 0.48 },
      effect: { xIn: 0.18, yIn: height * 0.67, wIn: width - 0.36, hIn: height * 0.25 },
      flavor: { xIn: 0.18, yIn: height * 0.82, wIn: width - 0.36, hIn: 0.42 },
      icon: { xIn: width - 0.65, yIn: 0.15, wIn: 0.48, hIn: 0.48 },
    };
    const rect = presets[type] || presets.effect;
    return {
      id: uid(), type, label: ({
        "card-background": "Background", "main-image": "Artwork", title: "Title",
        effect: "Rules", flavor: "Flavor", icon: "Icon",
      })[type] || "Element",
      ...rect, rotationDeg: 0, hidden: false, transparentBg: type !== "title" && type !== "effect",
      showOutline: false, textHtml: type === "title" ? "<div>CARD TITLE</div>" :
        type === "effect" ? "<div>Describe what this card does.</div>" :
        type === "flavor" ? "<div><em>Add a memorable line of flavor.</em></div>" : "",
      style: defaultStyle(type),
      imageDataUrl: "",
      imageFill: {
        mode: type === "card-background" ? "gradient" : "image",
        color: "#eef2f8", gradient: "linear-gradient(160deg, #f8fbff 0%, #dbe7ff 55%, #eee9ff 100%)",
        fit: type === "main-image" ? "cover" : "contain",
      },
      opacity: 1,
      strokeColor: "transparent",
      strokeWidth: 0,
      sourceAsset: null,
      icon: {
        valueEnabled: false, value: "1", tint: "#ffffff", badgeEnabled: true,
        badgeColor: "#1769ff", valueColor: "#ffffff", valueSizePt: 11,
        outlineColor: "#111827", outlineWidth: 0, offsetX: 0, offsetY: 0,
      },
      ...overrides,
    };
  }

  function makeTemplate(templateId = "ability", cardInput = {}) {
    const preset = cardInput.preset || "poker";
    const size = CARD_SIZES[preset] || CARD_SIZES.poker;
    const card = { preset, widthIn: cardInput.widthIn || size.widthIn, heightIn: cardInput.heightIn || size.heightIn };
    const w = card.widthIn;
    const h = card.heightIn;
    let tiles = [];
    if (templateId === "ability") {
      tiles = [
        baseTile("card-background", card),
        baseTile("title", card),
        baseTile("main-image", card),
        baseTile("effect", card),
        baseTile("icon", card, { icon: { ...baseTile("icon", card).icon, valueEnabled: true } }),
      ];
    } else if (templateId === "character") {
      tiles = [
        baseTile("card-background", card, { imageFill: { mode: "gradient", color: "#121827", gradient: "linear-gradient(160deg,#0d1323,#263b68)", fit: "cover" } }),
        baseTile("main-image", card, { xIn: 0.12, yIn: 0.12, wIn: w - 0.24, hIn: h * 0.59 }),
        baseTile("title", card, { yIn: h * 0.57, style: { ...defaultStyle("title"), preset: "accent-band", background: "#101827dd" } }),
        baseTile("effect", card, { yIn: h * 0.72, hIn: h * 0.2 }),
        baseTile("icon", card, { xIn: w - 0.66, yIn: h - 0.65, icon: { ...baseTile("icon", card).icon, valueEnabled: true } }),
      ];
    } else if (templateId === "item") {
      tiles = [
        baseTile("card-background", card, { imageFill: { mode: "gradient", color: "#f8edd5", gradient: "linear-gradient(160deg,#fffaf0,#ead9b9)", fit: "cover" } }),
        baseTile("title", card, { style: { ...defaultStyle("title"), background: "#6f4b27" } }),
        baseTile("main-image", card, { yIn: 0.72, hIn: h * 0.43 }),
        baseTile("effect", card, { yIn: h * 0.64, hIn: h * 0.19 }),
        baseTile("flavor", card, { yIn: h * 0.84, hIn: h * 0.1 }),
        baseTile("icon", card, { icon: { ...baseTile("icon", card).icon, valueEnabled: true, badgeColor: "#b7791f" } }),
      ];
    }
    return {
      name: templateId === "blank" ? "Blank Card" : `${templateId[0].toUpperCase()}${templateId.slice(1)} Card`,
      templateId, card, snapEnabled: true, gridStepIn: 0.05, showSafeBorder: true,
      tiles, updatedAt: new Date().toISOString(),
    };
  }

  function constrainRect(rect, card, isBackground = false) {
    const inset = isBackground ? 0 : SAFE_BORDER_IN;
    const min = 0.15;
    const next = { ...rect };
    next.wIn = clamp(Number(next.wIn) || min, min, card.widthIn - inset * 2);
    next.hIn = clamp(Number(next.hIn) || min, min, card.heightIn - inset * 2);
    next.xIn = clamp(Number(next.xIn) || 0, inset, card.widthIn - inset - next.wIn);
    next.yIn = clamp(Number(next.yIn) || 0, inset, card.heightIn - inset - next.hIn);
    return next;
  }

  function normalizeTile(raw, card, version = 2) {
    const type = raw?.type || "effect";
    const tile = {
      ...baseTile(type, card), ...raw,
      rotationDeg: version >= 3 ? normalizeRotation(raw.rotationDeg) : normalizeRotation(raw.rotationDeg || 0),
      style: { ...defaultStyle(type), ...(raw.style || {}) },
      imageFill: { ...baseTile(type, card).imageFill, ...(raw.imageFill || {}) },
      icon: { ...baseTile(type, card).icon, ...(raw.icon || {}) },
    };
    if (raw.style?.fontSizePt == null && raw.style?.fontSize != null) {
      tile.style.fontSizePt = legacyPxToPt(tile.style.fontSize);
    }
    if (raw.icon && raw.icon.valueEnabled == null) tile.icon.valueEnabled = Boolean(raw.icon.value);
    Object.assign(tile, constrainRect(tile, card, type === "card-background"));
    return tile;
  }

  function normalizeDesign(raw = {}, version = 2, fallbackName = "Imported Card") {
    const preset = raw.card?.preset || "poker";
    const size = CARD_SIZES[preset] || CARD_SIZES.poker;
    const card = {
      preset,
      widthIn: Number(raw.card?.widthIn) || size.widthIn,
      heightIn: Number(raw.card?.heightIn) || size.heightIn,
    };
    return {
      name: raw.name || fallbackName,
      templateId: raw.templateId || "custom",
      card,
      snapEnabled: raw.snapEnabled !== false,
      gridStepIn: clamp(Number(raw.gridStepIn) || 0.05, 0.005, 0.5),
      showSafeBorder: raw.showSafeBorder !== false,
      tiles: (raw.tiles || []).map((tile) => normalizeTile(tile, card, version)),
      updatedAt: raw.updatedAt || new Date().toISOString(),
    };
  }

  function normalizeProject(raw) {
    const version = Number(raw?.version) || 1;
    if (Array.isArray(raw?.cards)) {
      const cards = raw.cards.map((entry, index) => {
        const id = entry.id || uid();
        const design = normalizeDesign(entry.design || entry, version, entry.name || `Card ${index + 1}`);
        return { id, name: entry.name || design.name, design, selectedTileId: entry.selectedTileId || null };
      });
      return { version: 3, format: "martins-card-prototyper-project", activeCardId: cards.some((c) => c.id === raw.activeCardId) ? raw.activeCardId : cards[0]?.id, cards };
    }
    const design = normalizeDesign(raw, version, raw?.name);
    const id = uid();
    return { version: 3, format: "martins-card-prototyper-project", activeCardId: id, cards: [{ id, name: design.name, design, selectedTileId: design.tiles[0]?.id || null }] };
  }

  function createHistory(limit = 100) {
    const undo = [];
    const redo = [];
    return {
      push(command) { undo.push(command); if (undo.length > limit) undo.shift(); redo.length = 0; },
      undo() { const command = undo.pop(); if (!command) return false; command.undo(); redo.push(command); return true; },
      redo() { const command = redo.pop(); if (!command) return false; command.redo(); undo.push(command); return true; },
      clear() { undo.length = 0; redo.length = 0; },
      get canUndo() { return undo.length > 0; },
      get canRedo() { return redo.length > 0; },
    };
  }

  return {
    SAFE_BORDER_IN, CARD_SIZES, uid, clone, clamp, snap, normalizeRotation, legacyPxToPt,
    defaultStyle, baseTile, makeTemplate, constrainRect, normalizeTile, normalizeDesign,
    normalizeProject, createHistory,
  };
});
