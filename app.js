/* global PDFLib, html2canvas, PrototyperCore, PROTOTYPER_ICONS */
(() => {
  "use strict";
  const C = PrototyperCore;
  const icons = Array.isArray(window.PROTOTYPER_ICONS) ? window.PROTOTYPER_ICONS : [];
  const $ = (selector) => document.querySelector(selector);
  const $$ = (selector) => [...document.querySelectorAll(selector)];
  const byId = (id) => document.getElementById(id);
  const els = {};
  [
    "drawer","card-stage","card-wrap","canvas-viewport","selection-box","guide-x","guide-y","status",
    "undo-btn","redo-btn","quick-card-name","card-dimensions","zoom-out","zoom-in","zoom-fit",
    "empty-inspector","selection-inspector","element-kind","element-label","delete-element",
    "text-content-control","element-text","style-presets","text-essentials","text-align","text-color",
    "panel-color","image-essentials","icon-essentials","choose-asset-btn","replace-upload-btn",
    "icon-tint","badge-color","value-enabled","icon-value","icon-value-row","x-input","y-input",
    "w-input","h-input","rotation-input","text-more","font-family","font-size","frame-toggle",
    "transparent-toggle","artwork-appearance","fill-mode","fill-color","artwork-opacity",
    "artwork-opacity-val","stroke-color","stroke-width","stroke-width-val","asset-search","category-chips","asset-grid","asset-empty","icon-count",
    "upload-image-btn","image-file","asset-license-btn","cards-list","layers-list","new-card-btn","card-name","card-size",
    "custom-size","card-width","card-height","snap-toggle","safe-toggle","save-json-btn","load-json-btn",
    "project-file","start-over-btn","export-scope","preview-btn","preview-top-btn","png-btn","pdf-btn",
    "pdf-page","pdf-layout","pdf-copies","template-dialog","guide-after-template","restore-dialog",
    "discard-draft","restore-draft","license-dialog","tutorial-btn","tutorial-layer","tutorial-spotlight",
    "tutorial-callout","tutorial-progress","tutorial-title","tutorial-copy","tutorial-skip","tutorial-back",
    "tutorial-next","preview-dialog","preview-title","preview-image","preview-download","toast-region",
    "related-btn","related-menu","theme-btn","feedback-btn"
  ].forEach((id) => { els[id.replace(/-([a-z])/g, (_, c) => c.toUpperCase())] = byId(id); });

  const state = {
    cards: [], activeCardId: null, design: null, selectedTileId: null, placingType: null,
    zoom: 1, fitScale: 1, drag: null, assetCategory: "All", assetQuery: "",
    history: C.createHistory(80), suspendHistory: false, tutorialIndex: 0, tutorialActive: false,
  };
  const activeCard = () => state.cards.find((card) => card.id === state.activeCardId);
  const selectedTile = () => state.design?.tiles.find((tile) => tile.id === state.selectedTileId);
  const isText = (tile) => tile && ["title","effect","flavor"].includes(tile.type);
  const isImage = (tile) => tile && ["main-image","card-background","icon"].includes(tile.type);
  const pxPerIn = () => els.cardStage.clientWidth / state.design.card.widthIn;
  const toPx = (value) => value * pxPerIn();
  const fromPx = (value) => value / pxPerIn();
  const typeName = (type) => ({
    title:"Title",effect:"Rules",flavor:"Flavor","main-image":"Artwork",icon:"Icon / value","card-background":"Background",
  })[type] || "Element";
  const safeFile = (name) => String(name || "card").trim().replace(/[^a-z0-9-_]+/gi,"-").replace(/-+/g,"-").replace(/^-|-$/g,"") || "card";
  const toast = (message) => {
    els.status.textContent = message;
    const node = document.createElement("div"); node.className = "toast"; node.textContent = message;
    els.toastRegion.append(node); setTimeout(() => node.remove(), 2600);
  };

  function snapshot() {
    syncActive();
    return C.clone({ cards: state.cards, activeCardId: state.activeCardId, selectedTileId: state.selectedTileId });
  }
  function restore(snapshotValue) {
    state.cards = C.clone(snapshotValue.cards); state.activeCardId = snapshotValue.activeCardId;
    const card = activeCard() || state.cards[0]; state.activeCardId = card.id;
    state.design = C.clone(card.design); state.selectedTileId = snapshotValue.selectedTileId;
    renderAll(); scheduleDraft();
  }
  function pushHistory(before, after) {
    if (state.suspendHistory || JSON.stringify(before) === JSON.stringify(after)) return;
    state.history.push({ undo: () => restore(before), redo: () => restore(after) });
    updateHistoryButtons();
  }
  function mutate(fn, options = {}) {
    const before = options.noHistory ? null : snapshot();
    fn(); syncActive(); renderAll();
    if (before) pushHistory(before, snapshot());
    scheduleDraft();
  }
  function syncActive() {
    const card = activeCard();
    if (!card || !state.design) return;
    card.name = state.design.name; card.design = C.clone(state.design); card.selectedTileId = state.selectedTileId;
    state.design.updatedAt = new Date().toISOString();
  }
  function activateCard(id) {
    syncActive();
    const card = state.cards.find((item) => item.id === id); if (!card) return;
    state.activeCardId = id; state.design = C.clone(card.design);
    state.selectedTileId = state.design.tiles.some((tile) => tile.id === card.selectedTileId) ? card.selectedTileId : null;
    renderAll(); scheduleDraft();
  }

  function createProject(templateId = "ability") {
    const design = C.makeTemplate(templateId);
    const id = C.uid();
    state.cards = [{ id, name: design.name, design, selectedTileId: design.tiles[1]?.id || design.tiles[0]?.id || null }];
    state.activeCardId = id; state.design = C.clone(design); state.selectedTileId = state.cards[0].selectedTileId;
    state.history.clear(); renderAll(); scheduleDraft();
  }
  function addCard(templateId) {
    mutate(() => {
      const design = C.makeTemplate(templateId, state.design?.card || {});
      design.name = `${templateId[0].toUpperCase()}${templateId.slice(1)} Card ${state.cards.length + 1}`;
      const card = { id: C.uid(), name: design.name, design, selectedTileId: design.tiles[1]?.id || null };
      state.cards.push(card); syncActive(); state.activeCardId = card.id; state.design = C.clone(design);
      state.selectedTileId = card.selectedTileId;
    });
  }

  function updateGeometry() {
    const { widthIn, heightIn } = state.design.card;
    const availableW = Math.max(280, els.canvasViewport.clientWidth - 70);
    const availableH = Math.max(350, els.canvasViewport.clientHeight - 70);
    const baseWidth = Math.min(640, availableW, availableH * widthIn / heightIn);
    state.fitScale = baseWidth / 500;
    const width = 500 * state.fitScale * state.zoom;
    els.cardWrap.style.width = `${width}px`; els.cardWrap.style.aspectRatio = `${widthIn}/${heightIn}`;
    els.cardDimensions.textContent = `${widthIn.toFixed(2)} × ${heightIn.toFixed(2)} in`;
    els.zoomFit.textContent = state.zoom === 1 ? "Fit" : `${Math.round(state.zoom * 100)}%`;
  }
  function elementStyles(node, tile) {
    node.style.left = `${toPx(tile.xIn)}px`; node.style.top = `${toPx(tile.yIn)}px`;
    node.style.width = `${toPx(tile.wIn)}px`; node.style.height = `${toPx(tile.hIn)}px`;
    node.style.transform = `rotate(${tile.rotationDeg || 0}deg)`;
    node.style.zIndex = String(state.design.tiles.indexOf(tile) + 1);
    node.classList.toggle("design-frame", tile.showOutline === true);
  }
  function renderTile(tile) {
    const node = document.createElement("div");
    node.className = `card-element ${tile.type}`; node.dataset.tileId = tile.id; elementStyles(node, tile);
    if (isText(tile)) {
      node.classList.add("text-element");
      node.innerHTML = tile.textHtml || "";
      const style = tile.style || C.defaultStyle(tile.type);
      node.style.fontFamily = style.fontFamily; node.style.fontSize = `${toPx((style.fontSizePt || 10) / 72)}px`;
      node.style.fontWeight = style.fontWeight || 500; node.style.lineHeight = style.lineHeight || 1.2;
      node.style.letterSpacing = `${style.letterSpacing || 0}px`; node.style.textAlign = style.textAlign || "left";
      node.style.color = style.color || "#111827"; node.style.background = tile.transparentBg ? "transparent" : style.background || "#fff";
      node.style.justifyContent = ({top:"flex-start",middle:"center",bottom:"flex-end"})[style.verticalAlign] || "center";
      node.addEventListener("dblclick", (event) => beginTextEdit(event, node, tile));
    } else if (tile.type === "icon") {
      node.classList.add("icon-element"); node.classList.toggle("badge", tile.icon?.badgeEnabled !== false);
      node.style.setProperty("--badge", tile.icon?.badgeColor || "#1769ff");
      if (tile.imageDataUrl) {
        const glyph = document.createElement("div"); glyph.className = "icon-glyph";
        glyph.style.setProperty("--icon-src", `url("${tile.imageDataUrl}")`);
        glyph.style.setProperty("--icon-tint", tile.icon?.tint || "#fff"); node.append(glyph);
      }
      if (tile.icon?.valueEnabled) {
        const value = document.createElement("span"); value.className = "icon-value"; value.textContent = tile.icon.value || "1";
        value.style.fontSize = `${toPx((tile.icon.valueSizePt || 11) / 72)}px`; node.append(value);
      }
    } else {
      node.classList.add("image-element");
      const fill = tile.imageFill || {};
      const layers = [];
      if (tile.imageDataUrl) layers.push(`url("${tile.imageDataUrl}")`);
      if (fill.mode === "gradient") layers.push(fill.gradient || "linear-gradient(#fff,#e9eef7)");
      node.style.backgroundImage = layers.join(",");
      node.style.backgroundColor = fill.mode === "color" ? fill.color : "transparent";
      node.style.backgroundSize = tile.imageDataUrl ? (fill.fit || (tile.type === "card-background" ? "cover" : "contain")) : "cover";
      node.style.backgroundRepeat = "no-repeat"; node.style.backgroundPosition = "center";
      if (tile.opacity != null && tile.opacity < 1) node.style.opacity = tile.opacity;
      if ((tile.strokeWidth || 0) > 0) {
        node.style.border = `${tile.strokeWidth}px solid ${tile.strokeColor || "#000"}`;
        node.style.boxSizing = "border-box";
      }
    }
    if (tile.hidden) node.hidden = true;
    node.addEventListener("pointerdown", (event) => beginMove(event, tile));
    return node;
  }
  function renderStage() {
    updateGeometry(); els.cardStage.replaceChildren();
    state.design.tiles.forEach((tile) => els.cardStage.append(renderTile(tile)));
    if (state.design.showSafeBorder) {
      const safe = document.createElement("div"); safe.className = "safe-area";
      safe.style.inset = `${toPx(C.SAFE_BORDER_IN)}px`; els.cardStage.append(safe);
    }
    updateSelectionBox();
  }
  function updateSelectionBox() {
    const tile = selectedTile(); const node = tile && els.cardStage.querySelector(`[data-tile-id="${tile.id}"]`);
    if (!tile || !node || tile.hidden) { els.selectionBox.hidden = true; return; }
    els.selectionBox.hidden = false; elementStyles(els.selectionBox, tile);
  }

  function renderCards() {
    els.cardsList.replaceChildren();
    state.cards.forEach((card, index) => {
      const row = document.createElement("div"); row.className = `card-row${card.id === state.activeCardId ? " active" : ""}`;
      const button = document.createElement("button"); button.textContent = card.name || `Card ${index + 1}`;
      button.onclick = () => activateCard(card.id);
      const actions = document.createElement("div"); actions.className = "mini-actions";
      const copy = document.createElement("button"); copy.textContent = "⧉"; copy.title = "Duplicate";
      copy.onclick = () => mutate(() => {
        const design = C.clone(card.design); design.name = `${card.name} Copy`;
        const duplicate = { id:C.uid(), name:design.name, design, selectedTileId:card.selectedTileId };
        state.cards.splice(index + 1, 0, duplicate); state.activeCardId = duplicate.id; state.design = C.clone(design); state.selectedTileId = duplicate.selectedTileId;
      });
      const del = document.createElement("button"); del.textContent = "×"; del.title = "Delete"; del.disabled = state.cards.length === 1;
      del.onclick = () => {
        if (state.cards.length === 1 || !confirm(`Delete "${card.name}"?`)) return;
        mutate(() => { state.cards.splice(index,1); const next = state.cards[Math.max(0,index - 1)]; state.activeCardId = next.id; state.design = C.clone(next.design); state.selectedTileId = next.selectedTileId; });
      };
      actions.append(copy,del); row.append(button,actions); els.cardsList.append(row);
    });
  }
  function renderLayers() {
    els.layersList.replaceChildren();
    [...state.design.tiles].reverse().forEach((tile) => {
      const index = state.design.tiles.indexOf(tile);
      const row = document.createElement("div"); row.className = `layer-row${tile.id === state.selectedTileId ? " active" : ""}`;
      const button = document.createElement("button"); button.textContent = tile.label || typeName(tile.type);
      button.onclick = () => { state.selectedTileId = tile.id; renderAll(); };
      const actions = document.createElement("div"); actions.className = "mini-actions";
      [["↑",1],["↓",-1]].forEach(([label,direction]) => {
        const move = document.createElement("button"); move.textContent = label;
        move.disabled = index + direction < 0 || index + direction >= state.design.tiles.length;
        move.onclick = () => mutate(() => {
          const other = index + direction; [state.design.tiles[index],state.design.tiles[other]] = [state.design.tiles[other],state.design.tiles[index]];
        }); actions.append(move);
      });
      const eye = document.createElement("button"); eye.textContent = tile.hidden ? "○" : "●"; eye.title = "Toggle visibility";
      eye.onclick = () => mutate(() => { tile.hidden = !tile.hidden; }); actions.append(eye);
      row.append(button,actions); els.layersList.append(row);
    });
  }
  const STYLE_PRESETS = {
    title: [
      ["Clean Band","clean-band",{background:"#17233d",color:"#f8fbff",textAlign:"center"}],
      ["Accent Band","accent-band",{background:"#1769ff",color:"#ffffff",textAlign:"left"}],
      ["Minimal","minimal",{background:"#ffffff",color:"#17233d",textAlign:"left",transparent:true}],
    ],
    effect: [
      ["Paper","paper",{background:"#f8fafc",color:"#202839"}],
      ["Soft Panel","soft-panel",{background:"#e8efff",color:"#17233d"}],
      ["Transparent","transparent",{background:"#ffffff",color:"#202839",transparent:true}],
    ],
    flavor: [
      ["Paper","paper",{background:"#fffaf1",color:"#5d4935"}],
      ["Soft Panel","soft-panel",{background:"#f0e9dc",color:"#5d4935"}],
      ["Transparent","transparent",{background:"#ffffff",color:"#4b5565",transparent:true}],
    ],
  };
  function renderInspector() {
    const tile = selectedTile(); els.emptyInspector.hidden = Boolean(tile); els.selectionInspector.hidden = !tile;
    if (!tile) return;
    els.elementKind.textContent = typeName(tile.type); els.elementLabel.value = tile.label || typeName(tile.type);
    els.xInput.value = tile.xIn.toFixed(2); els.yInput.value = tile.yIn.toFixed(2);
    els.wInput.value = tile.wIn.toFixed(2); els.hInput.value = tile.hIn.toFixed(2);
    els.rotationInput.value = Math.round(tile.rotationDeg || 0); els.frameToggle.checked = tile.showOutline === true;
    els.transparentToggle.checked = Boolean(tile.transparentBg);
    const text = isText(tile); els.textContentControl.hidden = !text; els.textEssentials.hidden = !text;
    els.textMore.hidden = !text; els.imageEssentials.hidden = !isImage(tile);
    if (text) {
      els.elementText.value = stripHtml(tile.textHtml); els.textColor.value = tile.style.color || "#202839";
      els.panelColor.value = tile.style.background || "#ffffff"; els.fontFamily.value = tile.style.fontFamily;
      els.fontSize.value = tile.style.fontSizePt || 10;
      $$("#text-align button").forEach((button) => button.classList.toggle("active", button.dataset.value === tile.style.textAlign));
      els.stylePresets.replaceChildren();
      (STYLE_PRESETS[tile.type] || STYLE_PRESETS.effect).forEach(([label,id,values]) => {
        const button = document.createElement("button"); button.textContent = label;
        button.classList.toggle("active", tile.style.preset === id);
        button.onclick = () => mutate(() => {
          tile.style = {...tile.style,...values,preset:id}; if (values.transparent != null) tile.transparentBg = values.transparent;
        }); els.stylePresets.append(button);
      });
    } else {
      els.stylePresets.replaceChildren();
    }
    els.iconEssentials.hidden = tile.type !== "icon";
    if (tile.type === "icon") {
      els.iconTint.value = tile.icon.tint || "#ffffff"; els.badgeColor.value = tile.icon.badgeColor || "#1769ff";
      els.valueEnabled.checked = Boolean(tile.icon.valueEnabled); els.iconValue.value = tile.icon.value || "1";
      els.iconValueRow.hidden = !tile.icon.valueEnabled;
    }
    const isImg = isImage(tile);
    els.artworkAppearance.hidden = !isImg;
    if (isImg) {
      const fill = tile.imageFill || {};
      $$("#fill-mode button").forEach((b) => b.classList.toggle("active", b.dataset.value === (fill.mode || "image")));
      els.fillColor.value = fill.color || "#eef2f8";
      els.artworkOpacity.value = tile.opacity ?? 1;
      els.artworkOpacityVal.textContent = `${Math.round((tile.opacity ?? 1) * 100)}%`;
      els.strokeColor.value = tile.strokeColor || "#000000";
      els.strokeWidth.value = tile.strokeWidth ?? 0;
      els.strokeWidthVal.textContent = `${tile.strokeWidth ?? 0}px`;
    }
  }
  function stripHtml(value) { const div = document.createElement("div"); div.innerHTML = value || ""; return div.textContent || ""; }
  function renderSettings() {
    els.quickCardName.value = state.design.name; els.cardName.value = state.design.name;
    els.cardSize.value = state.design.card.preset; els.cardWidth.value = state.design.card.widthIn; els.cardHeight.value = state.design.card.heightIn;
    els.customSize.hidden = state.design.card.preset !== "custom"; els.snapToggle.checked = state.design.snapEnabled; els.safeToggle.checked = state.design.showSafeBorder;
  }
  function renderAll() { if (!state.design) return; renderStage(); renderCards(); renderLayers(); renderInspector(); renderSettings(); updateHistoryButtons(); }
  function updateHistoryButtons() { els.undoBtn.disabled = !state.history.canUndo; els.redoBtn.disabled = !state.history.canRedo; }

  function setDrawer(name) {
    $$(".drawer-view").forEach((view) => { view.hidden = view.dataset.view !== name; });
    $$(".rail-button").forEach((button) => button.classList.toggle("active", button.dataset.drawer === name));
    if (name === "assets") renderAssetLibrary();
  }
  function enterPlacement(type) {
    state.placingType = type; document.body.classList.add("placing");
    $$(".component-card").forEach((button) => button.classList.toggle("active", button.dataset.addType === type));
    toast(type === "card-background" ? "Click the card to add a full background." : `Click or drag on the card to place ${typeName(type).toLowerCase()}.`);
  }
  function stopPlacement() {
    state.placingType = null; document.body.classList.remove("placing");
    $$(".component-card").forEach((button) => button.classList.remove("active"));
    $(".placement-ghost")?.remove();
  }
  function stagePoint(event) {
    const rect = els.cardStage.getBoundingClientRect();
    return { x: C.clamp(event.clientX - rect.left,0,rect.width), y:C.clamp(event.clientY - rect.top,0,rect.height), rect };
  }
  function beginPlacement(event) {
    if (!state.placingType || event.button !== 0) return;
    event.preventDefault(); const point = stagePoint(event);
    const ghost = document.createElement("div"); ghost.className = "placement-ghost"; els.cardStage.append(ghost);
    state.drag = { kind:"place", type:state.placingType, startX:point.x, startY:point.y, pointerId:event.pointerId, ghost };
    els.cardStage.setPointerCapture(event.pointerId); updatePlacementGhost(point);
  }
  function updatePlacementGhost(point) {
    const drag = state.drag; if (!drag?.ghost) return;
    const left = Math.min(drag.startX,point.x), top = Math.min(drag.startY,point.y);
    drag.ghost.style.left=`${left}px`; drag.ghost.style.top=`${top}px`;
    drag.ghost.style.width=`${Math.max(4,Math.abs(point.x-drag.startX))}px`; drag.ghost.style.height=`${Math.max(4,Math.abs(point.y-drag.startY))}px`;
  }
  function commitPlacement(event) {
    const drag = state.drag; if (drag?.kind !== "place") return;
    const point = stagePoint(event); const dx = Math.abs(point.x-drag.startX), dy = Math.abs(point.y-drag.startY);
    const type = drag.type; const defaults = C.baseTile(type,state.design.card); let rect;
    if (type === "card-background") rect = {xIn:0,yIn:0,wIn:state.design.card.widthIn,hIn:state.design.card.heightIn};
    else if (dx > 12 && dy > 12) rect = {xIn:fromPx(Math.min(drag.startX,point.x)),yIn:fromPx(Math.min(drag.startY,point.y)),wIn:fromPx(dx),hIn:fromPx(dy)};
    else rect = {xIn:fromPx(point.x)-defaults.wIn/2,yIn:fromPx(point.y)-defaults.hIn/2,wIn:defaults.wIn,hIn:defaults.hIn};
    const tile = {...defaults,...C.constrainRect(rect,state.design.card,type==="card-background")};
    state.drag = null; stopPlacement();
    mutate(() => {
      if (type === "card-background") state.design.tiles = state.design.tiles.filter((item) => item.type !== "card-background");
      if (type === "card-background") state.design.tiles.unshift(tile); else state.design.tiles.push(tile);
      state.selectedTileId = tile.id;
    });
    if (isImage(tile) && type !== "card-background") setDrawer("assets");
    else if (isText(tile)) setTimeout(() => els.cardStage.querySelector(`[data-tile-id="${tile.id}"]`)?.dispatchEvent(new MouseEvent("dblclick",{bubbles:true})),0);
  }

  function beginMove(event, tile) {
    if (state.placingType || event.button !== 0 || event.target.closest("[contenteditable=true]")) return;
    event.preventDefault(); event.stopPropagation(); state.selectedTileId = tile.id; renderInspector(); updateSelectionBox();
    const before = snapshot();
    state.drag = {kind:"move",tileId:tile.id,pointerId:event.pointerId,startX:event.clientX,startY:event.clientY,xIn:tile.xIn,yIn:tile.yIn,before};
    event.currentTarget.setPointerCapture(event.pointerId);
  }
  function beginHandle(event) {
    const tile = selectedTile(); if (!tile || event.button !== 0) return;
    event.preventDefault(); event.stopPropagation();
    state.drag = {kind:event.currentTarget.dataset.handle === "rotate" ? "rotate":"resize",handle:event.currentTarget.dataset.handle,tileId:tile.id,pointerId:event.pointerId,startX:event.clientX,startY:event.clientY,xIn:tile.xIn,yIn:tile.yIn,wIn:tile.wIn,hIn:tile.hIn,rotationDeg:tile.rotationDeg||0,before:snapshot()};
    event.currentTarget.setPointerCapture(event.pointerId);
  }
  function smartSnap(tile, x, y) {
    if (!state.design.snapEnabled) return {x,y,gx:null,gy:null};
    const tolerance = fromPx(6), card = state.design.card;
    const xTargets = [C.SAFE_BORDER_IN,card.widthIn/2-tile.wIn/2,card.widthIn-C.SAFE_BORDER_IN-tile.wIn];
    const yTargets = [C.SAFE_BORDER_IN,card.heightIn/2-tile.hIn/2,card.heightIn-C.SAFE_BORDER_IN-tile.hIn];
    state.design.tiles.filter((other)=>other.id!==tile.id&&!other.hidden).forEach((other)=>{
      xTargets.push(other.xIn,other.xIn+other.wIn/2-tile.wIn/2,other.xIn+other.wIn-tile.wIn);
      yTargets.push(other.yIn,other.yIn+other.hIn/2-tile.hIn/2,other.yIn+other.hIn-tile.hIn);
    });
    let gx=null,gy=null;
    const xn=xTargets.find((target)=>Math.abs(target-x)<=tolerance); const yn=yTargets.find((target)=>Math.abs(target-y)<=tolerance);
    if(xn!=null){x=xn;gx=x+tile.wIn/2;}else x=C.snap(x,state.design.gridStepIn,true);
    if(yn!=null){y=yn;gy=y+tile.hIn/2;}else y=C.snap(y,state.design.gridStepIn,true);
    return {x,y,gx,gy};
  }
  function updateGuides(gx,gy) {
    els.guideY.hidden = gx == null; els.guideX.hidden = gy == null;
    if(gx!=null) els.guideY.style.left=`${toPx(gx)}px`; if(gy!=null) els.guideX.style.top=`${toPx(gy)}px`;
  }
  function pointerMove(event) {
    const drag=state.drag; if(!drag||event.pointerId!==drag.pointerId)return;
    if(drag.kind==="place"){updatePlacementGhost(stagePoint(event));return;}
    const tile=selectedTile(); if(!tile)return;
    const dx=fromPx(event.clientX-drag.startX),dy=fromPx(event.clientY-drag.startY);
    if(drag.kind==="move"){
      const snapped=smartSnap(tile,drag.xIn+dx,drag.yIn+dy); tile.xIn=snapped.x;tile.yIn=snapped.y;updateGuides(snapped.gx,snapped.gy);
      Object.assign(tile,C.constrainRect(tile,state.design.card,tile.type==="card-background"));
    } else if(drag.kind==="resize"){
      let x=drag.xIn,y=drag.yIn,w=drag.wIn,h=drag.hIn; const handle=drag.handle;
      if(handle.includes("e"))w=drag.wIn+dx;if(handle.includes("s"))h=drag.hIn+dy;
      if(handle.includes("w")){x=drag.xIn+dx;w=drag.wIn-dx;}if(handle.includes("n")){y=drag.yIn+dy;h=drag.hIn-dy;}
      if(event.shiftKey){const ratio=drag.wIn/drag.hIn;if(Math.abs(dx)>Math.abs(dy))h=w/ratio;else w=h*ratio;}
      Object.assign(tile,C.constrainRect({xIn:x,yIn:y,wIn:w,hIn:h},state.design.card,tile.type==="card-background"));
    } else if(drag.kind==="rotate"){
      const rect=els.cardStage.getBoundingClientRect(),cx=rect.left+toPx(tile.xIn+tile.wIn/2),cy=rect.top+toPx(tile.yIn+tile.hIn/2);
      let angle=Math.atan2(event.clientY-cy,event.clientX-cx)*180/Math.PI+90;
      if(event.shiftKey)angle=Math.round(angle/15)*15; else if(Math.abs(angle-Math.round(angle/90)*90)<4)angle=Math.round(angle/90)*90;
      tile.rotationDeg=C.normalizeRotation(angle);
    }
    const node=els.cardStage.querySelector(`[data-tile-id="${tile.id}"]`); if(node)elementStyles(node,tile);updateSelectionBox();renderInspector();
  }
  function pointerUp(event) {
    const drag=state.drag;if(!drag||event.pointerId!==drag.pointerId)return;
    if(drag.kind==="place"){commitPlacement(event);return;}
    state.drag=null;updateGuides(null,null);syncActive();pushHistory(drag.before,snapshot());renderAll();scheduleDraft();
  }
  function beginTextEdit(event,node,tile){
    event.preventDefault();event.stopPropagation();state.selectedTileId=tile.id;const before=snapshot();
    node.contentEditable="true";node.focus();document.execCommand("selectAll",false,null);
    const finish=()=>{node.contentEditable="false";tile.textHtml=node.innerHTML;syncActive();pushHistory(before,snapshot());renderAll();scheduleDraft();};
    node.addEventListener("blur",finish,{once:true});
    node.addEventListener("keydown",(e)=>{if(e.key==="Escape"){e.preventDefault();node.blur();}},{once:true});
  }

  function renderAssetLibrary() {
    const categories=["All",...new Set(icons.map((icon)=>icon.category))];els.categoryChips.replaceChildren();
    categories.forEach((category)=>{const button=document.createElement("button");button.textContent=category;button.classList.toggle("active",state.assetCategory===category);button.onclick=()=>{state.assetCategory=category;renderAssetLibrary();};els.categoryChips.append(button);});
    const query=state.assetQuery.toLowerCase();const filtered=icons.filter((icon)=>(state.assetCategory==="All"||icon.category===state.assetCategory)&&(!query||`${icon.name} ${icon.category}`.toLowerCase().includes(query)));
    els.iconCount.textContent=String(filtered.length);els.assetGrid.replaceChildren();els.assetEmpty.hidden=filtered.length>0;
    filtered.forEach((icon)=>{const button=document.createElement("button");button.className="asset-card";button.title=`${icon.name} · ${icon.category}`;const image=document.createElement("img");image.src=encodeURI(icon.thumbnail);image.alt=icon.name;button.append(image);button.onclick=()=>chooseAsset(icon);els.assetGrid.append(button);});
  }
  async function chooseAsset(icon) {
    let tile=selectedTile();
    if(!tile||!isImage(tile)){enterPlacement("icon");toast("Place an icon on the card, then choose the symbol again.");return;}
    try{
      const response=await fetch(encodeURI(icon.path));if(!response.ok)throw new Error("Icon could not be loaded.");
      const svg=await response.text();const dataUrl=`data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
      mutate(()=>{tile.imageDataUrl=dataUrl;tile.sourceAsset={id:icon.id,name:icon.name,category:icon.category,attribution:"fractal-symbols"};tile.imageFill.mode="image";});
      toast(`${icon.name} added. Remember to include Fractal Symbols attribution.`);
    }catch(error){toast(error.message);}
  }
  async function uploadImage(file) {
    if(!file)return;if(!["image/png","image/jpeg"].includes(file.type))return toast("Choose a PNG or JPG image.");
    if(file.size>2.5*1024*1024)return toast("Image is larger than 2.5 MB.");
    const dataUrl=await new Promise((resolve,reject)=>{const reader=new FileReader();reader.onload=()=>resolve(reader.result);reader.onerror=reject;reader.readAsDataURL(file);});
    const image=await new Promise((resolve,reject)=>{const node=new Image();node.onload=()=>resolve(node);node.onerror=reject;node.src=dataUrl;});
    if(image.width>3000||image.height>3000)return toast("Image dimensions exceed 3000 × 3000.");
    let tile=selectedTile();if(!tile||!isImage(tile)){enterPlacement("main-image");toast("Place an artwork element, then upload again.");return;}
    mutate(()=>{tile.imageDataUrl=dataUrl;tile.sourceAsset={kind:"upload",name:file.name};tile.imageFill.mode="image";});toast("Image added.");
  }

  const DB_NAME="martins-card-prototyper",STORE="drafts";let draftTimer;
  function openDb(){return new Promise((resolve,reject)=>{const request=indexedDB.open(DB_NAME,1);request.onupgradeneeded=()=>request.result.createObjectStore(STORE);request.onsuccess=()=>resolve(request.result);request.onerror=()=>reject(request.error);});}
  async function saveDraft(){try{syncActive();const db=await openDb();const tx=db.transaction(STORE,"readwrite");tx.objectStore(STORE).put({payload:buildPayload(),savedAt:new Date().toISOString()},"current");}catch{}}
  function scheduleDraft(){clearTimeout(draftTimer);draftTimer=setTimeout(saveDraft,500);}
  async function readDraft(){try{const db=await openDb();return await new Promise((resolve)=>{const request=db.transaction(STORE).objectStore(STORE).get("current");request.onsuccess=()=>resolve(request.result);request.onerror=()=>resolve(null);});}catch{return null;}}
  async function clearDraft(){try{const db=await openDb();db.transaction(STORE,"readwrite").objectStore(STORE).delete("current");}catch{}}
  function buildPayload(){syncActive();return{version:3,format:"martins-card-prototyper-project",savedAt:new Date().toISOString(),activeCardId:state.activeCardId,cards:C.clone(state.cards)};}
  function loadPayload(payload){const project=C.normalizeProject(payload);state.cards=project.cards;state.activeCardId=project.activeCardId;const card=activeCard()||state.cards[0];state.design=C.clone(card.design);state.selectedTileId=card.selectedTileId;state.history.clear();renderAll();scheduleDraft();}

  async function capturePng() {
    if(typeof html2canvas!=="function")throw new Error("Export library did not load.");
    const selected=state.selectedTileId,showSafe=state.design.showSafeBorder;state.selectedTileId=null;state.design.showSafeBorder=false;
    document.body.classList.add("exporting");renderStage();
    try{const canvas=await html2canvas(els.cardStage,{backgroundColor:"#ffffff",scale:Math.max(2,Math.min(4,Math.ceil(750/els.cardStage.clientWidth))),useCORS:true,logging:false});return canvas.toDataURL("image/png");}
    finally{document.body.classList.remove("exporting");state.selectedTileId=selected;state.design.showSafeBorder=showSafe;renderStage();}
  }
  function download(url,name){const link=document.createElement("a");link.href=url;link.download=name;link.click();}
  async function captureCards(scope){
    syncActive();const original=state.activeCardId,ids=scope==="all"?state.cards.map((c)=>c.id):[state.activeCardId],out=[];
    for(const id of ids){if(id!==state.activeCardId)activateCard(id);out.push({name:state.design.name,dataUrl:await capturePng(),attribution:state.design.tiles.some((t)=>t.sourceAsset?.attribution==="fractal-symbols")});}
    if(original!==state.activeCardId)activateCard(original);return out;
  }
  async function exportPng(){const scope=els.exportScope.value;const cards=await captureCards(scope);cards.forEach((card,index)=>setTimeout(()=>download(card.dataUrl,`${scope==="all"?String(index+1).padStart(2,"0")+"-":""}${safeFile(card.name)}.png`),index*120));toast(cards.length===1?"PNG exported.":`${cards.length} PNG files queued.`);}
  async function exportPdf(){
    const scope=els.exportScope.value,cards=await captureCards(scope),pageDef=els.pdfPage.value==="a4"?[595.28,841.89]:[612,792],layout=els.pdfLayout.value;
    const config=layout==="grid3x3"?{cols:3,rows:3,rotate:false}:layout==="grid2x3"?{cols:3,rows:2,rotate:false}:{cols:2,rows:4,rotate:true};
    const pdf=await PDFLib.PDFDocument.create(),copies=scope==="all"?cards.length:C.clamp(Number(els.pdfCopies.value)||1,1,300);
    const sources=scope==="all"?cards:Array.from({length:copies},()=>cards[0]);const embedded=[];
    for(const card of sources){embedded.push({image:await pdf.embedPng(await fetch(card.dataUrl).then((r)=>r.arrayBuffer())),attribution:card.attribution});}
    const cardW=(config.rotate?state.design.card.heightIn:state.design.card.widthIn)*72,cardH=(config.rotate?state.design.card.widthIn:state.design.card.heightIn)*72,gutter=layout==="gutterfold"?18:0;
    const per=config.cols*config.rows;
    for(let offset=0;offset<embedded.length;offset+=per){const page=pdf.addPage(pageDef);const totalW=config.cols*cardW+(config.cols-1)*gutter,totalH=config.rows*cardH,marginX=(pageDef[0]-totalW)/2,marginY=(pageDef[1]-totalH)/2;let attributed=false;
      embedded.slice(offset,offset+per).forEach((entry,i)=>{const col=i%config.cols,row=Math.floor(i/config.cols),x=marginX+col*(cardW+gutter),y=pageDef[1]-marginY-(row+1)*cardH;page.drawImage(entry.image,{x,y,width:cardW,height:cardH,rotate:config.rotate?PDFLib.degrees(90):undefined});drawGuides(page,x,y,cardW,cardH);attributed ||= entry.attribution;});
      if(attributed)page.drawText("Fractal Symbols by Felix Thalin · fractalsymbols.com · CC BY 4.0",{x:36,y:14,size:6,color:PDFLib.rgb(.35,.38,.44)});
    }
    const bytes=await pdf.save(),url=URL.createObjectURL(new Blob([bytes],{type:"application/pdf"}));download(url,`${scope==="all"?"all-cards":safeFile(state.design.name)}-${layout}.pdf`);setTimeout(()=>URL.revokeObjectURL(url),1000);toast("PDF exported.");
  }
  function drawGuides(page,x,y,w,h){const color=PDFLib.rgb(.2,.2,.22),length=7;[[x-length,y,x,y],[x,y-length,x,y],[x+w,y-length,x+w,y],[x+w,y,x+w+length,y],[x-length,y+h,x,y+h],[x,y+h,x,y+h+length],[x+w,y+h,x+w+length,y+h],[x+w,y+h,x+w,y+h+length]].forEach(([x1,y1,x2,y2])=>page.drawLine({start:{x:x1,y:y1},end:{x:x2,y:y2},thickness:.65,color}));}

  const tutorialSteps=[
    {target:".tool-rail",title:"A focused toolkit",copy:"Use Add, Assets, Cards, Layers, Project, and Export without crowding the canvas."},
    {target:".component-list",drawer:"add",title:"Choose a component",copy:"Select a component, then click or drag directly on the card to place it."},
    {target:"#card-stage",title:"A tactile canvas",copy:"Single-click to select. Drag to move, use eight handles to resize, and the top handle to rotate."},
    {target:"#selection-box",title:"Precise transforms",copy:"Elements snap to the safe area, card center, and nearby elements. Arrow keys provide exact nudging."},
    {target:"#inspector",title:"Essentials first",copy:"Common styling stays visible. Exact geometry and typography live under More controls."},
    {target:".rail-button[data-drawer='assets']",title:"Free icon artwork",copy:"Search 327 Fractal Symbols for icons or artwork. PDF sheets add the required attribution automatically."},
    {target:"#preview-top-btn",title:"Preview before export",copy:"Check a clean render, then export PNG or a print-ready PDF from the Export drawer."},
  ];
  function showTutorialStep(){
    const step=tutorialSteps[state.tutorialIndex];if(step.drawer)setDrawer(step.drawer);const target=$(step.target);if(!target){endTutorial();return;}target.scrollIntoView({block:"center",inline:"nearest"});const rect=target.getBoundingClientRect(),pad=7;Object.assign(els.tutorialSpotlight.style,{left:`${rect.left-pad}px`,top:`${rect.top-pad}px`,width:`${rect.width+pad*2}px`,height:`${rect.height+pad*2}px`});els.tutorialProgress.textContent=`Step ${state.tutorialIndex+1} of ${tutorialSteps.length}`;els.tutorialTitle.textContent=step.title;els.tutorialCopy.textContent=step.copy;els.tutorialBack.disabled=state.tutorialIndex===0;els.tutorialNext.textContent=state.tutorialIndex===tutorialSteps.length-1?"Finish":"Next";
    requestAnimationFrame(()=>{const box=els.tutorialCallout.getBoundingClientRect();let left=rect.right+14,top=rect.top;if(left+box.width>innerWidth-12)left=Math.max(12,rect.left-box.width-14);if(top+box.height>innerHeight-12)top=innerHeight-box.height-12;Object.assign(els.tutorialCallout.style,{left:`${Math.max(12,left)}px`,top:`${Math.max(12,top)}px`});});
  }
  function startTutorial(){state.tutorialActive=true;state.tutorialIndex=0;els.tutorialLayer.hidden=false;try{localStorage.setItem("martins_card_prototyper.tutorial_seen_v2","1");}catch{}showTutorialStep();}
  function endTutorial(){state.tutorialActive=false;els.tutorialLayer.hidden=true;}

  function bind() {
    $$(".rail-button").forEach((button)=>button.onclick=()=>setDrawer(button.dataset.drawer));
    $$(".component-card").forEach((button)=>button.onclick=()=>enterPlacement(button.dataset.addType));
    els.cardStage.addEventListener("pointerdown",(event)=>{if(state.placingType)beginPlacement(event);else if(event.target===els.cardStage){state.selectedTileId=null;renderAll();}});
    window.addEventListener("pointermove",pointerMove);window.addEventListener("pointerup",pointerUp);
    $$("#selection-box [data-handle]").forEach((handle)=>handle.addEventListener("pointerdown",beginHandle));
    window.addEventListener("resize",()=>renderStage());
    window.addEventListener("keydown",(event)=>{
      if(event.key==="Escape"){if(state.placingType)stopPlacement();return;}
      if(["INPUT","TEXTAREA","SELECT"].includes(document.activeElement?.tagName)||document.activeElement?.isContentEditable)return;
      if((event.metaKey||event.ctrlKey)&&event.key.toLowerCase()==="z"){event.preventDefault();event.shiftKey?state.history.redo():state.history.undo();updateHistoryButtons();return;}
      if((event.metaKey||event.ctrlKey)&&event.key.toLowerCase()==="y"){event.preventDefault();state.history.redo();updateHistoryButtons();return;}
      const tile=selectedTile();if(!tile)return;
      if(["ArrowLeft","ArrowRight","ArrowUp","ArrowDown"].includes(event.key)){event.preventDefault();const step=event.shiftKey?.1:.01;const before=snapshot();if(event.key==="ArrowLeft")tile.xIn-=step;if(event.key==="ArrowRight")tile.xIn+=step;if(event.key==="ArrowUp")tile.yIn-=step;if(event.key==="ArrowDown")tile.yIn+=step;Object.assign(tile,C.constrainRect(tile,state.design.card,tile.type==="card-background"));syncActive();pushHistory(before,snapshot());renderAll();scheduleDraft();}
      if(event.key==="Delete"||event.key==="Backspace"){event.preventDefault();deleteSelected();}
      if((event.metaKey||event.ctrlKey)&&event.key.toLowerCase()==="d"){event.preventDefault();mutate(()=>{const copy=C.clone(tile);copy.id=C.uid();copy.xIn+=.05;copy.yIn+=.05;state.design.tiles.push(copy);state.selectedTileId=copy.id;});}
    });
    els.undoBtn.onclick=()=>{state.history.undo();updateHistoryButtons();};els.redoBtn.onclick=()=>{state.history.redo();updateHistoryButtons();};
    els.zoomOut.onclick=()=>{state.zoom=C.clamp(state.zoom-.1,.5,2);renderStage();};els.zoomIn.onclick=()=>{state.zoom=C.clamp(state.zoom+.1,.5,2);renderStage();};els.zoomFit.onclick=()=>{state.zoom=1;renderStage();};
    [els.quickCardName,els.cardName].forEach((input)=>input.addEventListener("change",()=>mutate(()=>{state.design.name=input.value.trim()||"Untitled Card";})));
    els.elementLabel.onchange=()=>mutate(()=>{selectedTile().label=els.elementLabel.value;});
    els.elementText.onchange=()=>mutate(()=>{selectedTile().textHtml=`<div>${escapeHtml(els.elementText.value).replace(/\n/g,"<br>")}</div>`;});
    els.deleteElement.onclick=deleteSelected;
    $$("#text-align button").forEach((button)=>button.onclick=()=>mutate(()=>{selectedTile().style.textAlign=button.dataset.value;}));
    [["text-color",(t,v)=>t.style.color=v],["panel-color",(t,v)=>{t.style.background=v;t.transparentBg=false;}],["icon-tint",(t,v)=>t.icon.tint=v],["badge-color",(t,v)=>t.icon.badgeColor=v]].forEach(([id,fn])=>byId(id).addEventListener("input",(event)=>mutate(()=>fn(selectedTile(),event.target.value))));
    els.valueEnabled.onchange=()=>mutate(()=>{selectedTile().icon.valueEnabled=els.valueEnabled.checked;});els.iconValue.onchange=()=>mutate(()=>{selectedTile().icon.value=els.iconValue.value;});
    [["x-input","xIn"],["y-input","yIn"],["w-input","wIn"],["h-input","hIn"]].forEach(([id,key])=>byId(id).onchange=()=>mutate(()=>{selectedTile()[key]=Number(byId(id).value);Object.assign(selectedTile(),C.constrainRect(selectedTile(),state.design.card,selectedTile().type==="card-background"));}));
    els.rotationInput.onchange=()=>mutate(()=>{selectedTile().rotationDeg=C.normalizeRotation(els.rotationInput.value);});els.fontFamily.onchange=()=>mutate(()=>{selectedTile().style.fontFamily=els.fontFamily.value;});els.fontSize.onchange=()=>mutate(()=>{selectedTile().style.fontSizePt=C.clamp(Number(els.fontSize.value),4,72);});
    els.frameToggle.onchange=()=>mutate(()=>{selectedTile().showOutline=els.frameToggle.checked;});els.transparentToggle.onchange=()=>mutate(()=>{selectedTile().transparentBg=els.transparentToggle.checked;});
    $$("#fill-mode button").forEach((button)=>button.onclick=()=>mutate(()=>{selectedTile().imageFill.mode=button.dataset.value;}));
    els.fillColor.oninput=()=>mutate(()=>{selectedTile().imageFill.color=els.fillColor.value;selectedTile().imageFill.mode="color";});
    els.artworkOpacity.oninput=()=>{els.artworkOpacityVal.textContent=`${Math.round(els.artworkOpacity.value*100)}%`;mutate(()=>{selectedTile().opacity=Number(els.artworkOpacity.value);});};
    els.strokeColor.oninput=()=>mutate(()=>{selectedTile().strokeColor=els.strokeColor.value;});
    els.strokeWidth.oninput=()=>{els.strokeWidthVal.textContent=`${els.strokeWidth.value}px`;mutate(()=>{selectedTile().strokeWidth=Number(els.strokeWidth.value);});};
    els.chooseAssetBtn.onclick=()=>setDrawer("assets");els.replaceUploadBtn.onclick=els.uploadImageBtn.onclick=()=>els.imageFile.click();els.imageFile.onchange=()=>{uploadImage(els.imageFile.files?.[0]);els.imageFile.value="";};
    els.assetSearch.oninput=()=>{state.assetQuery=els.assetSearch.value;renderAssetLibrary();};els.assetLicenseBtn.onclick=()=>els.licenseDialog.showModal();
    els.newCardBtn.onclick=()=>{els.templateDialog.dataset.mode="add";els.templateDialog.showModal();};
    $$("[data-template]").forEach((button)=>button.onclick=()=>{const id=button.dataset.template,mode=els.templateDialog.dataset.mode;els.templateDialog.close();mode==="add"?addCard(id):createProject(id);if(els.guideAfterTemplate.checked)setTimeout(startTutorial,300);});
    $$("[data-close-dialog]").forEach((button)=>button.onclick=()=>byId(button.dataset.closeDialog).close());
    els.cardSize.onchange=()=>mutate(()=>{const preset=els.cardSize.value;state.design.card.preset=preset;if(preset!=="custom")Object.assign(state.design.card,C.CARD_SIZES[preset]);state.design.tiles=state.design.tiles.map((tile)=>{const next=C.normalizeTile(tile,state.design.card,3);if(next.type==="card-background")Object.assign(next,{xIn:0,yIn:0,wIn:state.design.card.widthIn,hIn:state.design.card.heightIn});return next;});});
    [els.cardWidth,els.cardHeight].forEach((input)=>input.onchange=()=>mutate(()=>{state.design.card.preset="custom";state.design.card.widthIn=C.clamp(Number(els.cardWidth.value),1,10);state.design.card.heightIn=C.clamp(Number(els.cardHeight.value),1,14);}));
    els.snapToggle.onchange=()=>mutate(()=>{state.design.snapEnabled=els.snapToggle.checked;});els.safeToggle.onchange=()=>mutate(()=>{state.design.showSafeBorder=els.safeToggle.checked;});
    els.saveJsonBtn.onclick=()=>{const blob=new Blob([JSON.stringify(buildPayload(),null,2)],{type:"application/json"}),url=URL.createObjectURL(blob);download(url,`${safeFile(state.design.name)}.json`);setTimeout(()=>URL.revokeObjectURL(url),1000);toast("Project JSON saved.");};
    els.loadJsonBtn.onclick=()=>els.projectFile.click();els.projectFile.onchange=async()=>{try{loadPayload(JSON.parse(await els.projectFile.files[0].text()));toast("Project loaded.");}catch{toast("That project file could not be loaded.");}els.projectFile.value="";};
    els.startOverBtn.onclick=()=>{if(confirm("Start over? Unsaved changes will be replaced.")){els.templateDialog.dataset.mode="new";els.templateDialog.showModal();}};
    [els.previewBtn,els.previewTopBtn].forEach((button)=>button.onclick=async()=>{try{const png=await capturePng();els.previewImage.src=png;els.previewTitle.textContent=state.design.name;els.previewDownload.onclick=()=>download(png,`${safeFile(state.design.name)}.png`);els.previewDialog.showModal();}catch(error){toast(error.message);}});
    els.pngBtn.onclick=()=>exportPng().catch((error)=>toast(error.message));els.pdfBtn.onclick=()=>exportPdf().catch((error)=>{console.error(error);toast("PDF export failed.");});
    els.tutorialBtn.onclick=startTutorial;els.tutorialSkip.onclick=endTutorial;els.tutorialBack.onclick=()=>{state.tutorialIndex=Math.max(0,state.tutorialIndex-1);showTutorialStep();};els.tutorialNext.onclick=()=>{if(state.tutorialIndex===tutorialSteps.length-1)endTutorial();else{state.tutorialIndex++;showTutorialStep();}};
    els.relatedBtn.onclick=(event)=>{event.stopPropagation();const open=els.relatedMenu.hidden;els.relatedMenu.hidden=!open;els.relatedBtn.setAttribute("aria-expanded",open);};document.addEventListener("click",()=>{els.relatedMenu.hidden=true;els.relatedBtn.setAttribute("aria-expanded","false");});
    els.themeBtn.onclick=()=>{const theme=document.documentElement.dataset.theme==="dark"?"light":"dark";document.documentElement.dataset.theme=theme;document.body.dataset.theme=theme;localStorage.setItem("martins_card_prototyper.theme",theme);byId("theme-color").content=theme==="dark"?"#090b10":"#f4f7fb";};
    els.feedbackBtn.onclick=()=>{location.href="mailto:help@pnpfinder.com?subject=Card%20Prototyper%20Feedback";};
  }
  function deleteSelected(){const tile=selectedTile();if(!tile)return;mutate(()=>{state.design.tiles=state.design.tiles.filter((item)=>item.id!==tile.id);state.selectedTileId=null;});}
  function escapeHtml(value){return String(value).replace(/[&<>"']/g,(char)=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[char]));}

  async function bootstrap(){
    bind();renderAssetLibrary();const draft=await readDraft();
    if(draft?.payload){els.restoreDialog.showModal();els.restoreDraft.onclick=()=>{els.restoreDialog.close();loadPayload(draft.payload);toast("Draft restored.");};els.discardDraft.onclick=async()=>{els.restoreDialog.close();await clearDraft();els.templateDialog.dataset.mode="new";els.templateDialog.showModal();};}
    else{createProject("ability");els.templateDialog.dataset.mode="new";els.templateDialog.showModal();}
  }
  bootstrap();
})();
