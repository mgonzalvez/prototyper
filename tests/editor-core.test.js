const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const core = require("../editor-core.js");

test("templates provide polished structures and a true blank option", () => {
  assert.deepEqual(["ability", "character", "item"].map((id) => core.makeTemplate(id).tiles.length), [5, 5, 6]);
  assert.equal(core.makeTemplate("blank").tiles.length, 0);
  assert.equal(core.makeTemplate("ability").gridStepIn, 0.05);
});

test("legacy designs migrate to physical typography and version 3", () => {
  const project = core.normalizeProject({
    version: 2, name: "Legacy", card: { preset: "poker", widthIn: 2.5, heightIn: 3.5 },
    tiles: [{ type: "title", style: { fontSize: 82 }, rotationDeg: 90 }],
  });
  assert.equal(project.version, 3);
  assert.equal(project.cards[0].design.tiles[0].style.fontSizePt, 29.52);
  assert.equal(project.cards[0].design.tiles[0].rotationDeg, 90);
});

test("geometry stays inside printable bounds", () => {
  const rect = core.constrainRect({ xIn: -2, yIn: 9, wIn: 9, hIn: 9 }, { widthIn: 2.5, heightIn: 3.5 });
  assert.deepEqual(rect, { xIn: 0.125, yIn: 0.125, wIn: 2.25, hIn: 3.25 });
});

test("history supports undo and redo", () => {
  let value = 1;
  const history = core.createHistory();
  history.push({ undo: () => { value = 1; }, redo: () => { value = 2; } });
  value = 2;
  assert.equal(history.undo(), true); assert.equal(value, 1);
  assert.equal(history.redo(), true); assert.equal(value, 2);
});

test("icon manifest is deduplicated and all assets exist", () => {
  const root = path.join(__dirname, "..");
  const source = fs.readFileSync(path.join(root, "icon-manifest.js"), "utf8");
  const sandbox = { window: {} };
  vm.runInNewContext(source, sandbox);
  const icons = sandbox.window.PROTOTYPER_ICONS;
  assert.equal(icons.length, 327);
  assert.equal(new Set(icons.map((icon) => icon.id)).size, icons.length);
  for (const icon of icons) {
    assert.ok(fs.existsSync(path.join(root, icon.path)), icon.path);
    assert.ok(fs.existsSync(path.join(root, icon.thumbnail)), icon.thumbnail);
  }
});

test("markup ids are unique and local resources exist", () => {
  const root = path.join(__dirname, "..");
  const html = fs.readFileSync(path.join(root, "index.html"), "utf8");
  const app = fs.readFileSync(path.join(root, "app.js"), "utf8");
  const ids = [...html.matchAll(/\bid="([^"]+)"/g)].map((match) => match[1]);
  assert.equal(ids.length, new Set(ids).size);
  const idSet = new Set(ids);
  const registry = app.match(/\[\n([\s\S]*?)\n  \]\.forEach/)[1];
  const registeredIds = [...registry.matchAll(/"([^"]+)"/g)].map((match) => match[1]);
  assert.deepEqual(registeredIds.filter((id) => !idSet.has(id)), []);
  const refs = [...html.matchAll(/(?:src|href)="([^"]+)"/g)]
    .map((match) => match[1].split("?")[0])
    .filter((ref) => !/^(?:https?:|mailto:|#)/.test(ref));
  for (const ref of refs) assert.ok(fs.existsSync(path.join(root, ref)), ref);
});
