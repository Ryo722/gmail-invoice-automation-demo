'use strict';

/**
 * Minimal, dependency-free Google Apps Script mock.
 *
 * Loads the real ../../Code.gs into an isolated node:vm context and provides
 * in-memory stand-ins for the services it uses (GmailApp, SpreadsheetApp,
 * DriveApp, PropertiesService, LockService, Utilities, Logger).
 *
 * Only the behaviour that Code.gs relies on is modelled. Faults can be injected
 * per operation through env.faults (see fault_()).
 */

const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const crypto = require('node:crypto');

// GAS_CODE_PATH lets mutation checks point the harness at a modified copy.
const CODE_PATH = process.env.GAS_CODE_PATH || path.join(__dirname, '..', '..', 'Code.gs');

const DEFAULT_MAX_ROWS = 1000;
const DEFAULT_MAX_COLUMNS = 26;

function toSignedBytes(buffer) {
  return Array.from(buffer, b => (b > 127 ? b - 256 : b));
}

function toBuffer(bytes) {
  return Buffer.from(bytes.map(b => b & 0xff));
}

function createEnv(options = {}) {
  const env = {
    faults: {},
    calls: [],
    properties: Object.assign(
      {SHEET_ID: 'sheet-1', FOLDER_ID: 'folder-1'},
      options.properties || {}
    ),
    lock: {holder: null, unavailable: false},
    labels: new Map(),
    threads: [],
    messages: new Map(),
    spreadsheets: new Map(),
    folders: new Map(),
    filesById: new Map(),
    nextFileId: 1,
    toasts: [],
  };

  function record(op, detail) {
    env.calls.push(Object.assign({op}, detail || {}));
  }

  // Fault hook: env.faults[op] may be a function (args) => void that throws,
  // or a number N meaning "throw on the next N calls".
  function fault_(op, args) {
    const f = env.faults[op];
    if (!f) return;
    if (typeof f === 'number') {
      if (f > 0) {
        env.faults[op] = f - 1;
        throw new Error(`Injected fault: ${op}`);
      }
      return;
    }
    f(args);
  }

  // ---------- Utilities ----------
  class Blob {
    constructor(bytes, contentType, name) {
      this._bytes = Buffer.from(bytes);
      this._contentType = contentType;
      this._name = name;
    }
    getBytes() { return toSignedBytes(this._bytes); }
    getContentType() { return this._contentType; }
    getName() { return this._name; }
    setName(name) { this._name = name; return this; }
    copyBlob() { return new Blob(this._bytes, this._contentType, this._name); }
    getDataAsString() { return this._bytes.toString('utf8'); }
  }

  const Utilities = {
    DigestAlgorithm: {SHA_256: 'SHA_256'},
    Charset: {UTF_8: 'UTF_8'},
    computeDigest(algorithm, value) {
      if (algorithm !== 'SHA_256') throw new Error('unsupported digest');
      const buf = typeof value === 'string'
        ? Buffer.from(value, 'utf8')
        : toBuffer(value);
      return toSignedBytes(crypto.createHash('sha256').update(buf).digest());
    },
    newBlob(data, contentType, name) {
      return new Blob(typeof data === 'string' ? Buffer.from(data) : toBuffer(data), contentType, name);
    },
  };

  // ---------- Gmail ----------
  class GmailLabel {
    constructor(name) { this._name = name; }
    getName() { return this._name; }
  }

  class GmailAttachment extends Blob {
    getSize() { return this._bytes.length; }
  }

  class GmailThread {
    constructor(id) { this._id = id; this._messages = []; this._labels = new Set(); }
    getId() { return this._id; }
    getMessages() { return this._messages.slice(); }
    addLabel(label) {
      fault_('thread.addLabel', {threadId: this._id});
      record('thread.addLabel', {threadId: this._id, label: label.getName()});
      this._labels.add(label.getName());
      return this;
    }
    getLabels() { return Array.from(this._labels).map(n => env.labels.get(n)); }
  }

  class GmailMessage {
    constructor(spec, thread) {
      this._spec = spec;
      this._thread = thread;
    }
    getId() { return this._spec.id; }
    getThread() { return this._thread; }
    getFrom() { return this._spec.from; }
    getDate() { return new Date(this._spec.date); }
    getSubject() { return this._spec.subject; }
    getPlainBody() { return this._spec.body; }
    isInTrash() { return Boolean(this._spec.inTrash); }
    getAttachments(opts) {
      fault_('message.getAttachments', {id: this._spec.id});
      return (this._spec.attachments || []).map(a => new GmailAttachment(
        Buffer.from(a.content, 'utf8'), a.contentType, a.name));
    }
  }

  function addMessage(spec) {
    const threadId = spec.threadId || `t-${spec.id}`;
    let thread = env.threads.find(t => t.getId() === threadId);
    if (!thread) {
      thread = new GmailThread(threadId);
      env.threads.push(thread);
    }
    const message = new GmailMessage(Object.assign({attachments: []}, spec), thread);
    thread._messages.push(message);
    env.messages.set(spec.id, message);
    return message;
  }

  const GmailApp = {
    search(query, start, max) {
      record('GmailApp.search', {query, start, max});
      return env.threads.slice(start, start + max);
    },
    getMessageById(id) {
      fault_('GmailApp.getMessageById', {id});
      const m = env.messages.get(id);
      if (!m) throw new Error(`Invalid argument: id (${id})`);
      return m;
    },
    getUserLabelByName(name) { return env.labels.get(name) || null; },
    createLabel(name) {
      record('GmailApp.createLabel', {name});
      const label = new GmailLabel(name);
      env.labels.set(name, label);
      return label;
    },
  };

  // ---------- Spreadsheet ----------
  function coerce(value, format) {
    // Sheets turns numeric-looking strings (including e-notation such as the
    // hex Gmail ID "1800000000e00001") into numbers unless the cell is
    // formatted as plain text ('@').
    if (typeof value === 'string' && format !== '@' && /^-?\d+(\.\d+)?(e[+-]?\d+)?$/i.test(value)) {
      return Number(value);
    }
    return value === undefined || value === null ? '' : value;
  }

  class Sheet {
    constructor(name) {
      this._name = name;
      this._cells = []; // [row][col] zero-based
      this._formats = [];
      this._maxRows = DEFAULT_MAX_ROWS;
      this._maxCols = DEFAULT_MAX_COLUMNS;
      this._validations = [];
      this._frozen = 0;
    }
    getName() { return this._name; }
    getMaxRows() { return this._maxRows; }
    getMaxColumns() { return this._maxCols; }
    getLastRow() {
      for (let r = this._cells.length - 1; r >= 0; r--) {
        const row = this._cells[r] || [];
        if (row.some(v => v !== '' && v !== undefined && v !== null)) return r + 1;
      }
      return 0;
    }
    getLastColumn() {
      let max = 0;
      this._cells.forEach(row => (row || []).forEach((v, c) => {
        if (v !== '' && v !== undefined && v !== null) max = Math.max(max, c + 1);
      }));
      return max;
    }
    setFrozenRows(n) { this._frozen = n; }
    insertRowsAfter(after, howMany) { this._maxRows += howMany; return this; }
    insertColumnsAfter(after, howMany) { this._maxCols += howMany; return this; }
    _get(r, c) { return ((this._cells[r - 1] || [])[c - 1]); }
    _set(r, c, v) {
      if (!this._cells[r - 1]) this._cells[r - 1] = [];
      const fmt = (this._formats[r - 1] || [])[c - 1];
      this._cells[r - 1][c - 1] = coerce(v, fmt);
    }
    _fmt(r, c, f) {
      if (!this._formats[r - 1]) this._formats[r - 1] = [];
      this._formats[r - 1][c - 1] = f;
    }
    getRange(row, col, numRows = 1, numCols = 1) {
      if (typeof row !== 'number') throw new Error('mock supports numeric getRange only');
      if (row < 1 || col < 1 || numRows < 1 || numCols < 1 ||
          row + numRows - 1 > this._maxRows || col + numCols - 1 > this._maxCols) {
        throw new Error(`The coordinates of the range are outside the dimensions of the sheet (${this._name})`);
      }
      return new Range(this, row, col, numRows, numCols);
    }
    appendRow(values) {
      fault_('sheet.appendRow', {sheet: this._name, values});
      const r = this.getLastRow() + 1;
      if (r > this._maxRows) this._maxRows = r;
      if (values.length > this._maxCols) this._maxCols = values.length;
      // New rows inherit the formatting of the row above (as in Sheets).
      if (r > 1 && this._formats[r - 2]) this._formats[r - 1] = this._formats[r - 2].slice();
      values.forEach((v, i) => this._set(r, i + 1, v));
      record('sheet.appendRow', {sheet: this._name, row: r});
      return this;
    }
    // Test helpers
    rows() {
      const last = this.getLastRow();
      const out = [];
      for (let r = 1; r <= last; r++) {
        out.push(((this._cells[r - 1]) || []).map(v => (v === undefined ? '' : v)));
      }
      return out;
    }
  }

  class Range {
    constructor(sheet, row, col, numRows, numCols) {
      Object.assign(this, {sheet, row, col, numRows, numCols});
    }
    getValues() {
      const out = [];
      for (let r = 0; r < this.numRows; r++) {
        const line = [];
        for (let c = 0; c < this.numCols; c++) {
          const v = this.sheet._get(this.row + r, this.col + c);
          line.push(v === undefined || v === null ? '' : v);
        }
        out.push(line);
      }
      return out;
    }
    getValue() { return this.getValues()[0][0]; }
    setValues(values) {
      fault_('range.setValues', {sheet: this.sheet._name, row: this.row, col: this.col, values});
      if (values.length !== this.numRows || values.some(v => v.length !== this.numCols)) {
        throw new Error('The number of rows or columns in the data does not match the range');
      }
      values.forEach((line, r) => line.forEach((v, c) => this.sheet._set(this.row + r, this.col + c, v)));
      record('range.setValues', {sheet: this.sheet._name, row: this.row, col: this.col, rows: this.numRows});
      return this;
    }
    setValue(v) {
      fault_('range.setValue', {sheet: this.sheet._name, row: this.row, col: this.col, value: v});
      for (let r = 0; r < this.numRows; r++) {
        for (let c = 0; c < this.numCols; c++) this.sheet._set(this.row + r, this.col + c, v);
      }
      return this;
    }
    clearContent() {
      for (let r = 0; r < this.numRows; r++) {
        for (let c = 0; c < this.numCols; c++) {
          const row = this.sheet._cells[this.row + r - 1];
          if (row) row[this.col + c - 1] = '';
        }
      }
      return this;
    }
    setNumberFormat(f) {
      for (let r = 0; r < this.numRows; r++) {
        for (let c = 0; c < this.numCols; c++) this.sheet._fmt(this.row + r, this.col + c, f);
      }
      return this;
    }
    setDataValidation(rule) {
      this.sheet._validations.push({row: this.row, col: this.col, numRows: this.numRows, rule});
      return this;
    }
  }

  class Spreadsheet {
    constructor(id) { this._id = id; this._sheets = new Map(); }
    getId() { return this._id; }
    getSheetByName(name) { return this._sheets.get(name) || null; }
    insertSheet(name) {
      const s = new Sheet(name);
      this._sheets.set(name, s);
      return s;
    }
    toast(msg) { env.toasts.push(msg); }
  }

  function addSpreadsheet(id) {
    const ss = new Spreadsheet(id);
    env.spreadsheets.set(id, ss);
    return ss;
  }

  const ui = {
    menus: [],
    createMenu(name) {
      const menu = {name, items: [], subMenus: [],
        addItem(label, fn) { this.items.push({label, fn}); return this; },
        addSeparator() { this.items.push({separator: true}); return this; },
        addSubMenu(sub) { this.subMenus.push(sub); return this; },
        addToUi() { ui.menus.push(this); return this; }};
      return menu;
    },
  };

  const SpreadsheetApp = {
    openById(id) {
      const ss = env.spreadsheets.get(id);
      if (!ss) throw new Error(`Unexpected error while getting the method or property openById on object SpreadsheetApp.`);
      return ss;
    },
    flush() { record('SpreadsheetApp.flush'); },
    getUi() { return ui; },
    newDataValidation() {
      const rule = {};
      const builder = {
        requireValueInList(list, dropdown) { rule.list = list; rule.dropdown = dropdown; return builder; },
        setAllowInvalid(v) { rule.allowInvalid = v; return builder; },
        setHelpText(t) { rule.helpText = t; return builder; },
        build() { return rule; },
      };
      return builder;
    },
  };

  // ---------- Drive ----------
  class DriveFile {
    constructor(id, blob, folder) {
      this._id = id;
      this._blob = blob.copyBlob();
      this._name = blob.getName();
      this._folder = folder;
      this._trashed = false;
    }
    getId() { return this._id; }
    getName() { return this._name; }
    setName(n) {
      fault_('file.setName', {id: this._id, name: n});
      this._name = n;
      return this;
    }
    getUrl() { return `https://drive.example.invalid/file/${this._id}`; }
    getSize() { return this._blob._bytes.length; }
    getBlob() { return this._blob.copyBlob(); }
    isTrashed() { return this._trashed; }
    setTrashed(v) { this._trashed = Boolean(v); return this; }
  }

  function iterator(items) {
    let i = 0;
    return {hasNext: () => i < items.length, next: () => items[i++]};
  }

  class Folder {
    constructor(id) { this._id = id; this._files = []; }
    getId() { return this._id; }
    createFile(blob) {
      fault_('folder.createFile.before', {name: blob.getName()});
      const file = new DriveFile(`file-${env.nextFileId++}`, blob, this);
      this._files.push(file);
      env.filesById.set(file.getId(), file);
      record('folder.createFile', {name: file.getName(), id: file.getId()});
      // Fault *after* the file exists models an ambiguous failure.
      fault_('folder.createFile.after', {file});
      return file;
    }
    getFilesByName(name) {
      return iterator(this._files.filter(f => f.getName() === name));
    }
    getFiles() { return iterator(this._files.slice()); }
    // Test helpers
    liveFiles() { return this._files.filter(f => !f._trashed); }
  }

  function addFolder(id) {
    const f = new Folder(id);
    env.folders.set(id, f);
    return f;
  }

  const DriveApp = {
    getFolderById(id) {
      const f = env.folders.get(id);
      if (!f) throw new Error(`No item with the given ID could be found, or you do not have permission to access it. (${id})`);
      return f;
    },
    getFileById(id) {
      const f = env.filesById.get(id);
      if (!f) throw new Error(`No item with the given ID could be found (${id})`);
      return f;
    },
  };

  // ---------- Properties / Lock / Logger ----------
  const PropertiesService = {
    getScriptProperties() {
      return {
        getProperty(name) {
          return Object.prototype.hasOwnProperty.call(env.properties, name) ? env.properties[name] : null;
        },
      };
    },
  };

  let lockSeq = 0;
  const LockService = {
    getScriptLock() {
      const lock = {
        _id: ++lockSeq,
        tryLock(ms) {
          record('lock.tryLock', {ms});
          if (env.lock.unavailable) return false;
          if (env.lock.holder && env.lock.holder !== lock._id) return false;
          env.lock.holder = lock._id;
          return true;
        },
        waitLock(ms) {
          if (!lock.tryLock(ms)) throw new Error('Lock timeout');
        },
        hasLock() { return env.lock.holder === lock._id; },
        releaseLock() {
          record('lock.releaseLock');
          if (env.lock.holder === lock._id) env.lock.holder = null;
        },
      };
      return lock;
    },
  };

  const logs = [];
  const Logger = {log: (...args) => { logs.push(args.join(' ')); }};

  const context = vm.createContext({
    GmailApp, SpreadsheetApp, DriveApp, PropertiesService, LockService,
    Utilities, Logger, console,
  });
  vm.runInContext(fs.readFileSync(CODE_PATH, 'utf8'), context, {filename: 'Code.gs'});

  Object.assign(env, {
    context,
    logs,
    ui,
    addMessage,
    addSpreadsheet,
    addFolder,
    Utilities,
    /** Evaluate an expression inside the Code.gs context (e.g. 'SAFE_INTAKE'). */
    evalInScript(code) { return vm.runInContext(code, context); },
    /** Call a top-level Code.gs function by name. */
    call(name, ...args) {
      const fn = context[name];
      if (typeof fn !== 'function') throw new Error(`Code.gs does not define ${name}()`);
      return fn(...args);
    },
    sheet(name, ssId = 'sheet-1') {
      const ss = env.spreadsheets.get(ssId);
      return ss ? ss.getSheetByName(name) : null;
    },
    folder(id = 'folder-1') { return env.folders.get(id); },
    countCalls(op) { return env.calls.filter(c => c.op === op).length; },
    sideEffectCounts() {
      return {
        createFile: env.countCalls('folder.createFile'),
        appendRecords: env.calls.filter(c => c.op === 'sheet.appendRow' && c.sheet === 'Records').length,
        addLabel: env.countCalls('thread.addLabel'),
        createLabel: env.countCalls('GmailApp.createLabel'),
      };
    },
  });

  if (options.withDefaults !== false) {
    addSpreadsheet('sheet-1');
    addFolder('folder-1');
  }

  return env;
}

module.exports = {createEnv};
