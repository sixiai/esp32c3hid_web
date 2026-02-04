const ROWS = 5;
const COLS = 5;
const MAX_STEPS = 16;
const FN_ROW = 4;
const FN_COL = 2;

const ACTION_NONE = 0;
const ACTION_NORMAL = 1;
const ACTION_CHORD = 2;
const ACTION_MACRO = 3;

const MAP_WIRE_VER = 3;
const MAP_LAYERS = 2;

const CMD_GET_MAP = 0x10;
const CMD_SET_MAP = 0x11;
const CMD_SAVE = 0x12;
const CMD_LOAD = 0x13;
const CMD_RESET_DEFAULT = 0x14;
const CMD_GET_INFO = 0x15;
const CMD_PING = 0x01;
const CMD_PONG = 0x02;
const CMD_ERROR = 0x7F;

const CMD_NAME = {
  [CMD_GET_MAP]: '读取配置',
  [CMD_SET_MAP]: '写入 RAM',
  [CMD_SAVE]: '保存到 Flash',
  [CMD_LOAD]: '从 Flash 读取',
  [CMD_RESET_DEFAULT]: '恢复默认并写入',
  [CMD_GET_INFO]: '获取信息'
};

let port;
let reader;
let writer;
let seq = 1;
let cfg = null;
let defaultCfg = null;
let heartbeatTimer = null;
let saveAfterWrite = false;
let loadAfterLoad = false;
let resetAfterReset = false;
const pending = new Map();
let selected = { layer: 'base', r: 0, c: 0 };

const statusEl = document.getElementById('status');
const gridBase = document.getElementById('gridBase');
const gridFn = document.getElementById('gridFn');
const btnConnect = document.getElementById('btnConnect');
const btnRead = document.getElementById('btnRead');
const btnWrite = document.getElementById('btnWrite');
const btnSave = document.getElementById('btnSave');
const btnLoad = document.getElementById('btnLoad');
const btnReset = document.getElementById('btnReset');
const resultBox = document.getElementById('resultBox');

const actionTypeEl = document.getElementById('actionType');
const keysInput = document.getElementById('keysInput');
const modChecks = Array.from(document.querySelectorAll('.mod'));
const macroTable = document.querySelector('#macroTable tbody');
const btnAddStep = document.getElementById('btnAddStep');
const btnApply = document.getElementById('btnApply');
const editorTitle = document.getElementById('editorTitle');
const btnCapture = document.getElementById('btnCapture');
const captureHint = document.getElementById('captureHint');
let captureMode = null;

function setStatus(msg) {
  if (statusEl) statusEl.textContent = msg;
}

function setResult(msg, ok) {
  if (!resultBox) return;
  resultBox.textContent = msg;
  resultBox.classList.remove('ok', 'err');
  if (ok === true) resultBox.classList.add('ok');
  if (ok === false) resultBox.classList.add('err');
}

function enableControls(on) {
  btnRead.disabled = !on;
  btnWrite.disabled = !on;
  btnSave.disabled = !on;
  btnLoad.disabled = !on;
  btnReset.disabled = !on;
}

function createEmptyConfig() {
  return {
    base: [...Array(ROWS)].map(() => [...Array(COLS)].map(() => emptyAction())),
    fn:   [...Array(ROWS)].map(() => [...Array(COLS)].map(() => emptyAction()))
  };
}

function cloneAction(a) {
  return {
    type: a.type,
    mod: a.mod,
    keys: (a.keys || [0,0,0,0,0,0]).slice(0, 6),
    macroLen: a.macroLen || 0,
    macro: (a.macro || []).map(s => ({
      op: s.op || 0,
      mod: s.mod || 0,
      keys: (s.keys || [0,0,0,0,0,0]).slice(0, 6),
      delay: s.delay || 0
    }))
  };
}

function cloneConfig(src) {
  return {
    base: src.base.map(row => row.map(cloneAction)),
    fn: src.fn.map(row => row.map(cloneAction))
  };
}

function initDefaultConfig() {
  defaultCfg = createEmptyConfig();
  const setKey = (config, layer, r, c, keycode) => {
    const a = config[layer][r][c];
    a.type = ACTION_NORMAL;
    a.mod = 0;
    a.keys = [0,0,0,0,0,0];
    if (keycode >= 0xE0 && keycode <= 0xE7) {
      a.mod = 1 << (keycode - 0xE0);
    } else {
      a.keys[0] = keycode;
    }
  };
  // Base layer defaults
  setKey(defaultCfg, 'base', 0, 0, 0x29); // Esc
  setKey(defaultCfg, 'base', 0, 1, 0x1E); // 1
  setKey(defaultCfg, 'base', 0, 2, 0x1F); // 2
  setKey(defaultCfg, 'base', 0, 3, 0x20); // 3
  setKey(defaultCfg, 'base', 0, 4, 0x21); // 4

  setKey(defaultCfg, 'base', 1, 0, 0x2B); // Tab
  setKey(defaultCfg, 'base', 1, 1, 0x14); // Q
  setKey(defaultCfg, 'base', 1, 2, 0x1A); // W
  setKey(defaultCfg, 'base', 1, 3, 0x08); // E
  setKey(defaultCfg, 'base', 1, 4, 0x15); // R

  setKey(defaultCfg, 'base', 2, 0, 0x39); // Caps
  setKey(defaultCfg, 'base', 2, 1, 0x04); // A
  setKey(defaultCfg, 'base', 2, 2, 0x16); // S
  setKey(defaultCfg, 'base', 2, 3, 0x07); // D
  setKey(defaultCfg, 'base', 2, 4, 0x09); // F

  setKey(defaultCfg, 'base', 3, 0, 0xE1); // LShift
  setKey(defaultCfg, 'base', 3, 1, 0x1D); // Z
  setKey(defaultCfg, 'base', 3, 2, 0x1B); // X
  setKey(defaultCfg, 'base', 3, 3, 0x06); // C
  setKey(defaultCfg, 'base', 3, 4, 0x19); // V

  setKey(defaultCfg, 'base', 4, 0, 0xE0); // LCtrl
  setKey(defaultCfg, 'base', 4, 1, 0xE3); // LWin
  // base[4][2] = Fn (leave none)
  setKey(defaultCfg, 'base', 4, 3, 0xE2); // LAlt
  setKey(defaultCfg, 'base', 4, 4, 0x2C); // Space

  // Fn layer defaults: Fn+1..4 -> F1..F4
  setKey(defaultCfg, 'fn', 0, 1, 0x3A); // F1
  setKey(defaultCfg, 'fn', 0, 2, 0x3B); // F2
  setKey(defaultCfg, 'fn', 0, 3, 0x3C); // F3
  setKey(defaultCfg, 'fn', 0, 4, 0x3D); // F4

  cfg = cloneConfig(defaultCfg);
}

function emptyAction() {
  return { type: ACTION_NONE, mod: 0, keys: [0,0,0,0,0,0], macroLen: 0, macro: [] };
}

const KEY_LABELS = new Map();
(() => {
  // A-Z
  for (let i = 0; i < 26; i++) KEY_LABELS.set(0x04 + i, String.fromCharCode(65 + i));
  // 1-0
  const nums = ['1','2','3','4','5','6','7','8','9','0'];
  for (let i = 0; i < nums.length; i++) KEY_LABELS.set(0x1E + i, nums[i]);
  KEY_LABELS.set(0x28, 'Enter');
  KEY_LABELS.set(0x29, 'Esc');
  KEY_LABELS.set(0x2A, 'Backspace');
  KEY_LABELS.set(0x2B, 'Tab');
  KEY_LABELS.set(0x2C, 'Space');
  KEY_LABELS.set(0x2D, '-');
  KEY_LABELS.set(0x2E, '=');
  KEY_LABELS.set(0x2F, '[');
  KEY_LABELS.set(0x30, ']');
  KEY_LABELS.set(0x31, '\\');
  KEY_LABELS.set(0x33, ';');
  KEY_LABELS.set(0x34, '\'');
  KEY_LABELS.set(0x35, '`');
  KEY_LABELS.set(0x36, ',');
  KEY_LABELS.set(0x37, '.');
  KEY_LABELS.set(0x38, '/');
  KEY_LABELS.set(0x39, 'Caps');
  for (let i = 0; i < 12; i++) KEY_LABELS.set(0x3A + i, `F${i + 1}`);
  KEY_LABELS.set(0x4A, 'Home');
  KEY_LABELS.set(0x4B, 'PgUp');
  KEY_LABELS.set(0x4D, 'End');
  KEY_LABELS.set(0x4E, 'PgDn');
  KEY_LABELS.set(0x4F, '→');
  KEY_LABELS.set(0x50, '←');
  KEY_LABELS.set(0x51, '↓');
  KEY_LABELS.set(0x52, '↑');
  KEY_LABELS.set(0x4C, 'Del');
  KEY_LABELS.set(0x49, 'Ins');
})();

function keycodeToLabel(code) {
  if (KEY_LABELS.has(code)) return KEY_LABELS.get(code);
  return `0x${code.toString(16).padStart(2,'0')}`;
}

function tokenToKeycode(token) {
  if (!token) return null;
  const t = token.trim();
  if (!t) return null;
  if (t.startsWith('0x') || t.startsWith('0X')) {
    const v = parseInt(t, 16);
    return Number.isFinite(v) ? (v & 0xFF) : null;
  }
  const upper = t.toUpperCase();
  if (upper.length === 1) {
    const ch = upper.charCodeAt(0);
    if (ch >= 65 && ch <= 90) return 0x04 + (ch - 65);
    if (ch >= 48 && ch <= 57) {
      const map = [0x27, 0x1E,0x1F,0x20,0x21,0x22,0x23,0x24,0x25,0x26];
      return map[ch - 48];
    }
  }
  const alias = {
    ESC: 0x29, ESCAPE: 0x29, ENTER: 0x28, RETURN: 0x28,
    TAB: 0x2B, SPACE: 0x2C, BACKSPACE: 0x2A, DEL: 0x4C,
    DELETE: 0x4C, CAPS: 0x39, CAPSLOCK: 0x39,
    UP: 0x52, DOWN: 0x51, LEFT: 0x50, RIGHT: 0x4F,
    HOME: 0x4A, END: 0x4D, PGUP: 0x4B, PGDN: 0x4E,
    INS: 0x49, INSERT: 0x49
  };
  if (alias[upper] !== undefined) return alias[upper];
  if (upper.startsWith('F')) {
    const n = parseInt(upper.slice(1), 10);
    if (n >= 1 && n <= 12) return 0x3A + (n - 1);
  }
  return null;
}

function hidFromKeyEvent(e) {
  const code = e.code || '';
  if (code.startsWith('Key') && code.length === 4) {
    const ch = code.charCodeAt(3);
    return 0x04 + (ch - 65);
  }
  if (code.startsWith('Digit')) {
    const d = parseInt(code.slice(5), 10);
    if (!Number.isFinite(d)) return null;
    const map = [0x27, 0x1E,0x1F,0x20,0x21,0x22,0x23,0x24,0x25,0x26];
    return map[d];
  }
  if (code.startsWith('F')) {
    const n = parseInt(code.slice(1), 10);
    if (n >= 1 && n <= 12) return 0x3A + (n - 1);
  }
  const table = {
    Escape: 0x29,
    Enter: 0x28,
    Tab: 0x2B,
    Space: 0x2C,
    Backspace: 0x2A,
    Minus: 0x2D,
    Equal: 0x2E,
    BracketLeft: 0x2F,
    BracketRight: 0x30,
    Backslash: 0x31,
    Semicolon: 0x33,
    Quote: 0x34,
    Backquote: 0x35,
    Comma: 0x36,
    Period: 0x37,
    Slash: 0x38,
    CapsLock: 0x39,
    ArrowRight: 0x4F,
    ArrowLeft: 0x50,
    ArrowDown: 0x51,
    ArrowUp: 0x52,
    Home: 0x4A,
    End: 0x4D,
    PageUp: 0x4B,
    PageDown: 0x4E,
    Insert: 0x49,
    Delete: 0x4C
  };
  return table[code] ?? null;
}

function actionLabel(a) {
  if (a.type === ACTION_NONE) return '空';
  if (a.type === ACTION_MACRO) return '宏';
  const mods = [];
  if (a.mod & 0x01) mods.push('Ctrl');
  if (a.mod & 0x02) mods.push('Shift');
  if (a.mod & 0x04) mods.push('Alt');
  if (a.mod & 0x08) mods.push('Win');
  const keys = a.keys.filter(k => k !== 0).map(k => keycodeToLabel(k));
  return [...mods, ...keys].join('+') || '-';
}

function isFnKeyPos(r, c) {
  return r === FN_ROW && c === FN_COL;
}

function actionEqual(a, b) {
  if (a.type !== b.type) return false;
  if (a.mod !== b.mod) return false;
  for (let i = 0; i < 6; i++) {
    if ((a.keys[i] || 0) !== (b.keys[i] || 0)) return false;
  }
  if (a.type !== ACTION_MACRO) return true;
  const mlen = Math.min(a.macroLen || 0, MAX_STEPS);
  const blen = Math.min(b.macroLen || 0, MAX_STEPS);
  if (mlen !== blen) return false;
  for (let i = 0; i < mlen; i++) {
    const as = a.macro[i] || {};
    const bs = b.macro[i] || {};
    if ((as.op || 0) !== (bs.op || 0)) return false;
    if ((as.mod || 0) !== (bs.mod || 0)) return false;
    for (let k = 0; k < 6; k++) {
      if ((as.keys?.[k] || 0) !== (bs.keys?.[k] || 0)) return false;
    }
    if ((as.delay || 0) !== (bs.delay || 0)) return false;
  }
  return true;
}

function renderGrid(layer, target) {
  target.innerHTML = '';
  for (let r = 0; r < ROWS; r++) {
    for (let c = 0; c < COLS; c++) {
      const div = document.createElement('div');
      div.className = 'key';
      if (isFnKeyPos(r, c)) {
        div.textContent = 'Fn';
        div.classList.add('locked');
      } else {
        const label = actionLabel(cfg[layer][r][c]);
        div.textContent = label;
        if (label === '空') div.classList.add('empty');
        if (defaultCfg && !actionEqual(cfg[layer][r][c], defaultCfg[layer][r][c])) {
          div.classList.add('changed');
        }
        if (selected.layer === layer && selected.r === r && selected.c === c) {
          div.classList.add('active');
        }
        div.onclick = () => {
          applyCurrentEditor();
          selected = { layer, r, c };
          renderAll();
          loadEditor();
        };
      }
      target.appendChild(div);
    }
  }
}

function renderAll() {
  renderGrid('base', gridBase);
  renderGrid('fn', gridFn);
}

function getAction() {
  return cfg[selected.layer][selected.r][selected.c];
}

function loadEditor() {
  const a = getAction();
  editorTitle.textContent = `${selected.layer.toUpperCase()} [${selected.r},${selected.c}]`;
  actionTypeEl.value = a.type;
  modChecks.forEach(ch => ch.checked = !!(a.mod & (1 << parseInt(ch.dataset.bit, 10))));
  keysInput.value = a.keys.filter(k => k !== 0).map(k => keycodeToLabel(k)).join(',');
  renderMacroTable(a);
}

function renderMacroTable(a) {
  macroTable.innerHTML = '';
  (a.macro || []).forEach((s) => addMacroRow(s));
}

function addMacroRow(step = null) {
  if (macroTable.children.length >= MAX_STEPS) return;
  const tr = document.createElement('tr');

  const opTd = document.createElement('td');
  const opSel = document.createElement('select');
  ['按下','释放','延迟'].forEach((t,i) => {
    const opt = document.createElement('option');
    opt.value = i;
    opt.textContent = t;
    opSel.appendChild(opt);
  });
  opSel.value = step ? step.op : 0;
  opTd.appendChild(opSel);

  const modTd = document.createElement('td');
  const modWrap = document.createElement('div');
  modWrap.className = 'macro-mods';
  const modDefs = [
    { bit: 0, label: 'Ctrl' },
    { bit: 1, label: 'Shift' },
    { bit: 2, label: 'Alt' },
    { bit: 3, label: 'Win' }
  ];
  modDefs.forEach(def => {
    const lab = document.createElement('label');
    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.className = 'macro-mod';
    cb.dataset.bit = String(def.bit);
    if (step && (step.mod & (1 << def.bit))) cb.checked = true;
    lab.appendChild(cb);
    lab.appendChild(document.createTextNode(def.label));
    modWrap.appendChild(lab);
  });
  modTd.appendChild(modWrap);

  const keysTd = document.createElement('td');
  const keysIn = document.createElement('input');
  keysIn.placeholder = 'A,B';
  keysIn.value = step ? step.keys.filter(k=>k!==0).map(k=>keycodeToLabel(k)).join(',') : '';
  keysTd.appendChild(keysIn);
  const capRow = document.createElement('div');
  capRow.className = 'macro-capture';
  const capBtn = document.createElement('button');
  capBtn.type = 'button';
  capBtn.textContent = '录入';
  const capHint = document.createElement('span');
  capHint.textContent = '点击录入';
  capBtn.onclick = () => {
    clearCapture();
    captureMode = {
      type: 'macro',
      input: keysIn,
      modBoxes: modWrap.querySelectorAll('input.macro-mod'),
      btn: capBtn,
      hint: capHint
    };
    capBtn.classList.add('active');
    capHint.textContent = '请按下按键';
  };
  capRow.appendChild(capBtn);
  capRow.appendChild(capHint);
  keysTd.appendChild(capRow);

  const delayTd = document.createElement('td');
  const delayIn = document.createElement('input');
  delayIn.placeholder = '0';
  delayIn.value = step ? step.delay : 0;
  delayTd.appendChild(delayIn);

  const delTd = document.createElement('td');
  const delBtn = document.createElement('button');
  delBtn.textContent = 'X';
  delBtn.onclick = () => tr.remove();
  delTd.appendChild(delBtn);

  tr.append(opTd, modTd, keysTd, delayTd, delTd);
  macroTable.appendChild(tr);
}

btnAddStep.onclick = () => addMacroRow();

function clearCapture() {
  if (!captureMode) return;
  if (captureMode.type === 'main') {
    if (btnCapture) btnCapture.classList.remove('active');
    if (captureHint) captureHint.textContent = '点击后按下键盘';
  } else if (captureMode.type === 'macro') {
    if (captureMode.btn) captureMode.btn.classList.remove('active');
    if (captureMode.hint) captureMode.hint.textContent = '点击录入';
  }
  captureMode = null;
}

function applyModsFromEvent(modBoxes, e) {
  if (!modBoxes) return;
  modBoxes.forEach(cb => {
    const bit = parseInt(cb.dataset.bit, 10);
    let on = false;
    if (bit === 0) on = e.ctrlKey;
    else if (bit === 1) on = e.shiftKey;
    else if (bit === 2) on = e.altKey;
    else if (bit === 3) on = e.metaKey;
    cb.checked = on;
  });
}

if (btnCapture) {
  btnCapture.onclick = () => {
    clearCapture();
    captureMode = { type: 'main' };
    btnCapture.classList.add('active');
    if (captureHint) captureHint.textContent = '请按下要设置的按键';
    keysInput.focus();
  };
}

window.addEventListener('keydown', (e) => {
  if (!captureMode) return;
  e.preventDefault();
  const code = hidFromKeyEvent(e);
  if (captureMode.type === 'main') {
    applyModsFromEvent(modChecks, e);
    if (code !== null) keysInput.value = keycodeToLabel(code);
    else keysInput.value = '';
  } else if (captureMode.type === 'macro') {
    applyModsFromEvent(captureMode.modBoxes, e);
    if (code !== null) captureMode.input.value = keycodeToLabel(code);
    else captureMode.input.value = '';
  }
  clearCapture();
});

function applyCurrentEditor() {
  if (isFnKeyPos(selected.r, selected.c)) {
    setResult('Fn 键不可修改', false);
    return;
  }
  const a = getAction();
  a.type = parseInt(actionTypeEl.value, 10);
  a.mod = modChecks.reduce((m, ch) => ch.checked ? (m | (1 << parseInt(ch.dataset.bit, 10))) : m, 0);
  a.keys = parseKeys(keysInput.value);
  a.macro = readMacroTable();
  a.macroLen = a.macro.length;
  renderAll();
}

btnApply.onclick = () => {
  applyCurrentEditor();
};

function parseKeys(str) {
  const keys = [0,0,0,0,0,0];
  if (!str.trim()) return keys;
  const parts = str.split(',').map(s => s.trim()).filter(Boolean);
  for (let i = 0; i < Math.min(6, parts.length); i++) {
    const v = tokenToKeycode(parts[i]);
    keys[i] = (v === null) ? 0 : (v & 0xFF);
  }
  return keys;
}

function readMacroTable() {
  const steps = [];
  for (const tr of macroTable.children) {
    const tds = tr.querySelectorAll('td');
    const op = parseInt(tds[0].querySelector('select').value, 10);
    const modChecks = tds[1].querySelectorAll('input.macro-mod');
    let mod = 0;
    modChecks.forEach(cb => {
      if (cb.checked) mod |= (1 << parseInt(cb.dataset.bit, 10));
    });
    const keys = parseKeys(tds[2].querySelector('input').value);
    const delayVal = parseInt(tds[3].querySelector('input').value, 10);
    const delay = Number.isFinite(delayVal) ? (delayVal & 0xFFFF) : 0;
    steps.push({ op, mod, keys, delay });
  }
  return steps.slice(0, MAX_STEPS);
}

function buildBinaryConfig() {
  if (!defaultCfg) initDefaultConfig();
  const bytes = [];
  bytes.push(MAP_WIRE_VER, ROWS, COLS, MAP_LAYERS, 0);
  bytes.push(0, 0); // count placeholder
  let count = 0;
  for (const layer of ['base','fn']) {
    const layerIdx = layer === 'base' ? 0 : 1;
    for (let r = 0; r < ROWS; r++) {
      for (let c = 0; c < COLS; c++) {
        const a = cfg[layer][r][c];
        const d = defaultCfg[layer][r][c];
        if (actionEqual(a, d)) continue;
        count++;
        const isMacro = a.type === ACTION_MACRO;
        const mlen = isMacro ? Math.min(a.macroLen || 0, MAX_STEPS) : 0;
        bytes.push(layerIdx & 0xFF);
        bytes.push(r & 0xFF);
        bytes.push(c & 0xFF);
        bytes.push(a.type & 0xFF);
        bytes.push(a.mod & 0xFF);
        for (let i = 0; i < 6; i++) bytes.push((a.keys[i] || 0) & 0xFF);
        bytes.push(mlen & 0xFF);
        if (isMacro) {
          for (let i = 0; i < mlen; i++) {
            const s = a.macro[i] || { op: 0, mod: 0, keys: [0,0,0,0,0,0], delay: 0 };
            bytes.push(s.op & 0xFF);
            bytes.push(s.mod & 0xFF);
            for (let k = 0; k < 6; k++) bytes.push((s.keys[k] || 0) & 0xFF);
            bytes.push(s.delay & 0xFF);
            bytes.push((s.delay >> 8) & 0xFF);
          }
        }
      }
    }
  }
  bytes[5] = count & 0xFF;
  bytes[6] = (count >> 8) & 0xFF;
  const data = Uint8Array.from(bytes);
  const crc = crc16(data);
  const payload = new Uint8Array(data.length + 2);
  payload.set(data, 0);
  payload[data.length] = crc & 0xFF;
  payload[data.length + 1] = (crc >> 8) & 0xFF;
  return payload;
}

function parseBinaryConfig(payload) {
  if (!defaultCfg) initDefaultConfig();
  if (payload.length < 9) throw new Error('Payload too short');
  const crcRx = payload[payload.length - 2] | (payload[payload.length - 1] << 8);
  const data = payload.slice(0, payload.length - 2);
  if (crc16(data) !== crcRx) throw new Error('CRC mismatch');

  let off = 0;
  const ver = data[off++], rows = data[off++], cols = data[off++], layers = data[off++], _mode = data[off++];
  const count = data[off++] | (data[off++] << 8);
  if (ver !== MAP_WIRE_VER || rows !== ROWS || cols !== COLS || layers !== MAP_LAYERS) {
    throw new Error('Format mismatch');
  }

  cfg = cloneConfig(defaultCfg);

  for (let i = 0; i < count; i++) {
    if (off + 12 > data.length) throw new Error('Truncated entry');
    const layer = data[off++];
    const row = data[off++];
    const col = data[off++];
    if (layer >= MAP_LAYERS || row >= ROWS || col >= COLS) throw new Error('Bad index');
    const a = emptyAction();
    a.type = data[off++];
    a.mod = data[off++];
    a.keys = [];
    for (let k = 0; k < 6; k++) a.keys.push(data[off++]);
    let mlen = data[off++];
    if (a.type !== ACTION_MACRO && mlen !== 0) throw new Error('Bad macro len');
    if (a.type === ACTION_MACRO) {
      if (mlen > MAX_STEPS) throw new Error('Macro too long');
      a.macroLen = mlen;
      a.macro = [];
      for (let s = 0; s < mlen; s++) {
        if (off + 10 > data.length) throw new Error('Truncated macro');
        const op = data[off++];
        const mod = data[off++];
        const keys = [];
        for (let k = 0; k < 6; k++) keys.push(data[off++]);
        const delay = data[off++] | (data[off++] << 8);
        a.macro.push({ op, mod, keys, delay });
      }
    }
    if (layer === 0) cfg.base[row][col] = a;
    else cfg.fn[row][col] = a;
  }
  if (off !== data.length) throw new Error('Extra data');
}

function crc16(buf) {
  let crc = 0xFFFF;
  for (let b of buf) {
    crc ^= (b << 8);
    for (let i = 0; i < 8; i++) {
      if (crc & 0x8000) crc = ((crc << 1) ^ 0x1021) & 0xFFFF;
      else crc = (crc << 1) & 0xFFFF;
    }
  }
  return crc;
}

function cobsEncode(data) {
  const out = [];
  let codeIndex = 0;
  let code = 1;
  out.push(0);
  for (const b of data) {
    if (b === 0) {
      out[codeIndex] = code;
      code = 1;
      codeIndex = out.length;
      out.push(0);
    } else {
      out.push(b);
      code++;
      if (code === 0xFF) {
        out[codeIndex] = code;
        code = 1;
        codeIndex = out.length;
        out.push(0);
      }
    }
  }
  out[codeIndex] = code;
  return new Uint8Array(out);
}

function cobsDecode(data) {
  const out = [];
  for (let i = 0; i < data.length;) {
    const code = data[i++];
    if (code === 0) throw new Error('Bad COBS');
    for (let j = 1; j < code; j++) out.push(data[i++]);
    if (code !== 0xFF && i < data.length) out.push(0);
  }
  return new Uint8Array(out);
}

function cmdName(cmd) {
  return CMD_NAME[cmd] || `0x${cmd.toString(16)}`;
}

function trackPending(cmd, seqId) {
  const name = cmdName(cmd);
  setResult(`正在${name}...`);
  const timeout = setTimeout(() => {
    if (pending.has(seqId)) {
      pending.delete(seqId);
      setResult(`超时：${name}`, false);
    }
  }, 5000);
  pending.set(seqId, { cmd, timeout });
}

function clearPending(seqId, ok, msg) {
  const p = pending.get(seqId);
  if (p) {
    clearTimeout(p.timeout);
    pending.delete(seqId);
  }
  if (msg) setResult(msg, ok);
}

async function sendFrame(type, payload = new Uint8Array(), opts = {}) {
  if (!writer) return;
  const hdr = new Uint8Array(1 + 1 + 1 + 2);
  hdr[0] = 1; // ver
  hdr[1] = type;
  hdr[2] = seq++ & 0xFF;
  if (!opts.silent) {
    trackPending(type, hdr[2]);
  }
  hdr[3] = payload.length & 0xFF;
  hdr[4] = (payload.length >> 8) & 0xFF;

  const raw = new Uint8Array(1 + hdr.length + payload.length + 2);
  raw[0] = 0xA5;
  raw.set(hdr, 1);
  raw.set(payload, 1 + hdr.length);
  const crc = crc16(raw.slice(1, 1 + hdr.length + payload.length));
  raw[raw.length - 2] = crc & 0xFF;
  raw[raw.length - 1] = (crc >> 8) & 0xFF;

  const enc = cobsEncode(raw);
  const framed = new Uint8Array(enc.length + 1);
  framed.set(enc, 0);
  framed[framed.length - 1] = 0x00;
  await writer.write(framed);
}

async function readLoop() {
  let buffer = [];
  try {
    while (true) {
      const { value, done } = await reader.read();
      if (done) break;
      for (const b of value) {
        if (b === 0x00) {
          try {
            const dec = cobsDecode(new Uint8Array(buffer));
            handleFrame(dec);
          } catch (e) {
            console.error(e);
          }
          buffer = [];
        } else {
          buffer.push(b);
        }
      }
    }
  } catch (e) {
    console.error(e);
  }
  setStatus('已断开');
  enableControls(false);
  if (heartbeatTimer) {
    clearInterval(heartbeatTimer);
    heartbeatTimer = null;
  }
}

function handleFrame(frame) {
  if (frame.length < 1 + 5 + 2) return;
  if (frame[0] !== 0xA5) return;
  const type = frame[2];
  const seqId = frame[3];
  const len = frame[4] | (frame[5] << 8);
  if (frame.length < 1 + 5 + len + 2) return;
  const crcRx = frame[1 + 5 + len] | (frame[1 + 5 + len + 1] << 8);
  const crc = crc16(frame.slice(1, 1 + 5 + len));
  if (crc !== crcRx) {
    clearPending(seqId, false, '校验失败');
    return;
  }
  const payload = frame.slice(6, 6 + len);

  if (type === CMD_ERROR) {
    const msg = new TextDecoder().decode(payload) || '错误';
    saveAfterWrite = false;
    loadAfterLoad = false;
    resetAfterReset = false;
    clearPending(seqId, false, `失败：${msg}`);
    return;
  }

  if (type === CMD_GET_INFO) {
    clearPending(seqId, true, '连接正常');
    return;
  }
  if (type === CMD_PONG) {
    return;
  }

  if (type === CMD_SET_MAP) {
    clearPending(seqId, true, '写入 RAM 成功');
    if (saveAfterWrite) {
      saveAfterWrite = false;
      sendFrame(CMD_SAVE);
    }
    return;
  }
  if (type === CMD_SAVE) { clearPending(seqId, true, '保存到 Flash 成功'); return; }
  if (type === CMD_LOAD) {
    clearPending(seqId, true, '从 Flash 读取成功');
    if (loadAfterLoad) {
      loadAfterLoad = false;
      sendFrame(CMD_GET_MAP);
    }
    return;
  }
  if (type === CMD_RESET_DEFAULT) {
    clearPending(seqId, true, '恢复默认并写入成功');
    if (resetAfterReset) {
      resetAfterReset = false;
      sendFrame(CMD_GET_MAP);
    }
    return;
  }

  if (type === CMD_GET_MAP) {
    try {
      parseBinaryConfig(payload);
      renderAll();
      loadEditor();
      clearPending(seqId, true, '读取配置成功');
    } catch (e) {
      clearPending(seqId, false, '解析失败');
    }
  }
}

btnConnect.onclick = async () => {
  if (!('serial' in navigator)) {
    setResult('浏览器不支持 Web Serial', false);
    return;
  }
  port = await navigator.serial.requestPort();
  await port.open({ baudRate: 115200 });
  reader = port.readable.getReader();
  writer = port.writable.getWriter();
  setStatus('已连接');
  setResult('等待操作…');
  enableControls(true);
  initDefaultConfig();
  renderAll();
  loadEditor();
  readLoop();
  if (heartbeatTimer) clearInterval(heartbeatTimer);
  heartbeatTimer = setInterval(() => {
    sendFrame(CMD_PING, new Uint8Array(), { silent: true });
  }, 10000);
  await sendFrame(CMD_GET_INFO);
};

// 初始化未连接时的默认显示
initDefaultConfig();
renderAll();
loadEditor();

btnRead.onclick = async () => { await sendFrame(CMD_GET_MAP); };
btnWrite.onclick = async () => {
  applyCurrentEditor();
  const payload = buildBinaryConfig();
  await sendFrame(CMD_SET_MAP, payload);
};
btnSave.onclick = async () => {
  applyCurrentEditor();
  const payload = buildBinaryConfig();
  saveAfterWrite = true;
  await sendFrame(CMD_SET_MAP, payload);
};
btnLoad.onclick = async () => {
  loadAfterLoad = true;
  await sendFrame(CMD_LOAD);
};
btnReset.onclick = async () => {
  resetAfterReset = true;
  await sendFrame(CMD_RESET_DEFAULT);
};
