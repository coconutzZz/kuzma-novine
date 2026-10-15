const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const projectRoot = path.resolve(__dirname, '..');
const source = fs.readFileSync(path.join(projectRoot, 'Code.gs'), 'utf8');
const manifest = JSON.parse(fs.readFileSync(path.join(projectRoot, 'appsscript.json'), 'utf8'));
assert(manifest.oauthScopes.includes('https://www.googleapis.com/auth/gmail.readonly'));
assert(!manifest.oauthScopes.some(scope => /gmail\.modify|mail\.google/.test(scope)));

function harness({ inline = true, pageSize = 100 } = {}) {
  let now = Date.parse('2026-10-15T08:00:00Z');
  const state = { props: {}, folders: [], files: [], books: [], triggers: [], messages: [], gets: [], logs: [], sent: [], failSend: false, quota: 100, failName: '', failed: false, failMoveName: '', moveFailed: false, listCalls: 0 };
  class ClockDate extends Date { static now() { return now; } }
  function iterator(items) { let i = 0; return { hasNext: () => i < items.length, next: () => items[i++] }; }
  class Folder {
    constructor(name, parent) { this.id = 'folder-' + state.folders.length; this.name = name; this.parent = parent; this.description = ''; this.files = []; state.folders.push(this); }
    getId() { return this.id; }
    getUrl() { return 'https://drive.google.com/drive/folders/' + this.id; }
    getDescription() { return this.description; }
    setDescription(value) { this.description = value; return this; }
    getFoldersByName(name) { return iterator(state.folders.filter(f => f.parent === this && f.name === name && !f.trashed)); }
    createFolder(name) { return new Folder(name, this); }
    getFilesByName(name) { return iterator(this.files.filter(f => f.name === name)); }
    getFiles() { return iterator(this.files.filter(f => !f.trashed)); }
    getFolders() { return iterator(state.folders.filter(f => f.parent === this && !f.trashed)); }
    isTrashed() { return Boolean(this.trashed); }
    setTrashed(value) { this.trashed = value; return this; }
    createFile(blob) {
      if (blob.name === state.failName && !state.failed) { state.failed = true; throw new Error('Poskusna napaka pri shranjevanju v Drive.'); }
      const file = {
        ...blob, id: 'file-' + state.files.length, parent: this,
        getId() { return this.id; }, getName() { return this.name; }, getMimeType() { return this.mimeType; },
        getBlob() { return { name: this.name, bytes: this.bytes, setName(name) { this.name = name; return this; } }; },
        isTrashed() { return Boolean(this.trashed); }, getParents() { return iterator([this.parent]); },
        moveTo(folder) {
          if (this.name.includes(state.failMoveName) && state.failMoveName && !state.moveFailed) { state.moveFailed = true; throw new Error('Napaka pri poskusnem premiku.'); }
          this.parent.files.splice(this.parent.files.indexOf(this), 1); this.parent = folder; folder.files.push(this); return this;
        }
      };
      this.files.push(file); state.files.push(file); return file;
    }
  }
  const root = new Folder('My Drive', null);
  class Sheet {
    constructor() { this.name = 'Sheet1'; this.rows = []; this.maxRows = 1000; }
    setName(name) { this.name = name; return this; }
    getName() { return this.name; }
    getLastRow() { let last = this.rows.length; while (last > 0 && !(this.rows[last - 1] || []).some(v => v !== '' && v !== undefined)) last--; return last; }
    getMaxRows() { return this.maxRows; }
    getMaxColumns() { return 26; }
    insertColumnsAfter() {}
    deleteRow(row) { this.rows.splice(row - 1, 1); }
    insertRowsAfter(index, count) { this.maxRows += count; return this; }
    getRange(row, col, height = 1, width = 1) {
      const sheet = this;
      return {
        setValues(values) {
          assert.equal(values.length, height);
          values.forEach((data, i) => { assert.equal(data.length, width); const target = sheet.rows[row - 1 + i] ||= []; data.forEach((value, j) => target[col - 1 + j] = typeof value === 'string' && value.startsWith("'") ? value.slice(1) : value); });
          return this;
        },
        getDisplayValues() { return Array.from({ length: height }, (_, i) => Array.from({ length: width }, (_, j) => String(sheet.rows[row - 1 + i]?.[col - 1 + j] ?? ''))); },
        getValues() { return Array.from({ length: height }, (_, i) => Array.from({ length: width }, (_, j) => sheet.rows[row - 1 + i]?.[col - 1 + j] ?? '')); },
        clearContent() { for (let i = 0; i < height; i++) { const target = sheet.rows[row - 1 + i] ||= []; for (let j = 0; j < width; j++) target[col - 1 + j] = ''; } return this; },
        getSheet() { return sheet; }, getRow() { return row; }, getLastRow() { return row + height - 1; }, getColumn() { return col; }, getLastColumn() { return col + width - 1; },
        setNumberFormat() { return this; }, setFontWeight() { return this; }, setDataValidation() { return this; }, setNote() { return this; }
      };
    }
    setFrozenRows() {} setColumnWidth() {} hideColumns() {} hideSheet() {}
  }
  class Book {
    constructor(name) { this.id = 'book-' + state.books.length; this.name = name; this.sheets = [new Sheet()]; state.books.push(this); }
    getId() { return this.id; } getUrl() { return 'https://docs.google.com/spreadsheets/d/' + this.id; }
    rename(name) { this.name = name; return this; }
    getSheets() { return this.sheets; }
    getSheetByName(name) { return this.sheets.find(s => s.name === name); }
    insertSheet(name) { const sheet = new Sheet(); sheet.setName(name); this.sheets.push(sheet); return sheet; }
    setSpreadsheetTimeZone(zone) { this.zone = zone; }
  }
  const context = {
    Date: ClockDate,
    console: { log: (...v) => state.logs.push(v.join(' ')), error: (...v) => state.logs.push(v.join(' ')) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, waitLock() {}, releaseLock() {} }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: key => state.props[key] ?? null, setProperty: (key, value) => state.props[key] = String(value), setProperties: values => Object.assign(state.props, values) }) },
    DriveApp: { getRootFolder: () => root, createFolder: name => root.createFolder(name), getFolderById: id => { const folder = state.folders.find(f => f.id === id); assert(folder, id); return folder; }, getFileById: id => { const file = state.files.find(f => f.id === id); assert(file, id); return file; } },
    SpreadsheetApp: {
      create: name => new Book(name), openById: id => { const book = state.books.find(b => b.id === id); assert(book); return book; }, flush() {},
      newDataValidation: () => ({ requireValueInRange() { return this; }, requireValueInList() { return this; }, setAllowInvalid() { return this; }, setHelpText() { return this; }, requireCheckbox() { return this; }, build() { return this; } })
    },
    MailApp: { getRemainingDailyQuota: () => state.quota, sendEmail: mail => { state.sent.push(mail); if (state.failSend) throw new Error('Neznan izid pošiljanja.'); } },
    Session: { getScriptTimeZone: () => 'Europe/Vienna' },
    ScriptApp: {
      getProjectTriggers: () => state.triggers,
      newTrigger: handler => ({ timeBased() { return this; }, forSpreadsheet() { return this; }, onEdit() { return this; }, everyMinutes(minutes) { assert.equal(minutes, 5); return this; }, create() { const trigger = { getHandlerFunction: () => handler }; state.triggers.push(trigger); return trigger; } }),
      deleteTrigger: trigger => state.triggers.splice(state.triggers.indexOf(trigger), 1)
    },
    Utilities: {
      base64DecodeWebSafe: data => [...Buffer.from(data, 'base64url')],
      base64Decode: data => [...Buffer.from(data, 'base64')],
      newBlob: (bytes, mimeType, name) => ({ bytes, mimeType, name, getDataAsString: charset => { assert.match(charset, /^utf-?8$/i); return Buffer.from(bytes).toString('utf8'); } })
    },
    Gmail: { Users: {
      getProfile: () => ({ emailAddress: 'editor@example.com' }),
      Messages: {
        list: (user, options) => {
          state.listCalls++; assert.equal(user, 'me'); assert.equal(options.includeSpamTrash, false);
          const after = Number(options.q.match(/after:(\d+)/)[1]) * 1000;
          const before = Number(options.q.match(/before:(\d+)/)[1]) * 1000;
          const excluded = options.q.includes('-in:sent') ? ['SENT', 'DRAFT', 'SPAM', 'TRASH'] : ['DRAFT', 'SPAM', 'TRASH'];
          const all = state.messages.filter(m => Number(m.internalDate) > after && Number(m.internalDate) < before && !m.labelIds.some(label => excluded.includes(label)));
          const offset = Number(options.pageToken || 0);
          return { messages: all.slice(offset, offset + pageSize).map(m => ({ id: m.id, threadId: m.threadId })), nextPageToken: offset + pageSize < all.length ? String(offset + pageSize) : undefined };
        },
        get: (user, id, options) => { state.gets.push(id); const message = state.messages.find(m => m.id === id); assert(message); if (options.format === 'raw') return { raw: Buffer.from('From: ' + message.payload.headers[1].value + '\r\nSubject: ' + id + '\r\n\r\nIzvirno besedilo').toString('base64url') }; assert.equal(options.format, 'full'); return message; },
        Attachments: { get: (user, messageId, attachmentId) => {
          assert.equal(user, 'me'); assert.equal(attachmentId, 'external-file'); return { data: Buffer.from('external attachment').toString('base64url') };
        } }
      }
    } }
  };
  vm.createContext(context);
  vm.runInContext(inline ? source : source.replace('INCLUDE_INLINE_IMAGES: true', 'INCLUDE_INLINE_IMAGES: false'), context);
  function attachment(filename, mimeType, { disposition = 'attachment', external = false } = {}) {
    return { filename, mimeType, headers: [{ name: 'Content-Disposition', value: disposition }], body: external ? { attachmentId: 'external-file' } : { data: Buffer.from(filename).toString('base64url') } };
  }
  function message(id, subject, parts, opts = {}) {
    const textPart = { mimeType: 'text/plain', body: { data: Buffer.from(opts.body || 'Besedilo prispevka.').toString('base64url') } };
    const result = { id, threadId: opts.threadId || 'thread-' + id, labelIds: opts.labels || ['INBOX', 'UNREAD'], internalDate: String(opts.date || now + 120000), payload: { mimeType: 'multipart/mixed', headers: [{ name: 'Subject', value: subject }, { name: 'From', value: opts.sender || 'Avtor <author@example.com>' }, { name: 'Message-ID', value: '<' + id + '@example.com>' }, { name: 'To', value: opts.to || 'editor@example.com' }], parts: [textPart, ...parts] } };
    state.messages.push(result); return result;
  }
  const run = name => vm.runInContext(name + '()', context);
  const sheet = name => state.books[0].getSheetByName(name);
  function confirm(id, selection) {
    const pending = sheet('Neujemajoči maili');
    const index = pending.rows.findIndex(row => row[11] === id);
    assert(index >= 1, 'Manjka čakajoče sporočilo: ' + id);
    pending.getRange(index + 1, 9, 1, 2).setValues([[selection, true]]);
    context.testEvent = { source: state.books[0], range: pending.getRange(index + 1, 10) };
    vm.runInContext('onPendingEdit(testEvent)', context);
  }
  return { state, root, run, attachment, message, advance: ms => now += ms, context, confirm, sheet,
    articleFolders: () => state.folders.filter(f => f.parent?.name === 'Naše novine 2026' && !f.name.startsWith('_') && !f.trashed),
    rows: () => sheet('Prejeti članki').rows };
}

function label(article) { return article[9] + ' [' + article[6] + ']'; }
function log(h, id) { return h.sheet('Dnevnik pošte').rows.find(row => row[0] === id); }
function pendingRows(h) { return h.sheet('Neujemajoči maili').rows.slice(1); }
function setStatus(h, row, status) {
  const range = h.sheet('Prejeti članki').getRange(row, 5);
  range.setValues([[status]]);
  h.context.testEvent = { source: h.state.books[0], range };
  vm.runInContext('onPendingEdit(testEvent)', h.context);
}

{
  const h = harness(); h.run('setup'); h.run('setup');
  assert.equal(h.state.triggers.length, 2); assert.equal(h.state.books.length, 1);
  assert.equal(h.state.books[0].name, 'Naše novine 2026 - evidenca');
  const labels = ['INBOX', 'UNREAD'];
  h.message('day1', 'Obletnica gasilskega društva', [h.attachment('clanek.docx', 'application/msword')], { labels });
  h.advance(6 * 60000); h.run('processSubmissions');
  assert.equal(h.rows().length, 2); assert.equal(h.rows()[1][4], 'Prejeto'); assert.equal(h.rows()[1][11], 1);
  assert.equal(h.articleFolders()[0].name, 'Obletnica gasilskega društva');
  assert.equal(h.articleFolders()[0].files.filter(f => f.mimeType === 'message/rfc822').length, 1);
  h.sheet('Sodelavci').getRange(2, 1, 1, 2).setValues([['author@example.com', 'Gasilsko društvo']]);
  h.sheet('Prejeti članki').getRange(2, 8).setValues([['Gasilsko društvo']]);
  h.advance(3 * 86400000);
  h.message('day3', 'Fotografije', [h.attachment('foto.jpg', 'image/jpeg'), h.attachment('foto.jpg', 'image/jpeg')], { body: 'Pošiljam slike za članek o obletnici gasilskega društva.' });
  h.advance(6 * 60000); h.run('processSubmissions');
  assert.equal(h.rows().length, 2); assert.equal(h.rows()[1][10], 0); assert.equal(pendingRows(h).length, 1);
  assert(pendingRows(h)[0][5].includes('Obletnica gasilskega društva')); assert(pendingRows(h)[0][6].includes('isti e-naslov'));
  assert.equal(log(h, 'day3')[12], 'Čaka na izbiro'); assert.equal(pendingRows(h)[0][8], '');
  h.confirm('day3', label(h.rows()[1]));
  assert.equal(pendingRows(h).length, 0); assert.equal(h.rows().length, 2); assert.equal(h.rows()[1][10], 2);
  assert.match(h.rows()[1][13], /Prejete nove slike \(2\)/); assert.equal(log(h, 'day3')[12], 'Povezano');
  assert.equal(h.articleFolders()[0].files.filter(f => f.name.includes('foto.jpg')).length, 2);
  const stagedFolder = h.state.folders.find(f => f.id === log(h, 'day3')[10]); assert(stagedFolder.trashed);
  h.advance(7 * 86400000);
  h.message('day10', 'Popravek članka', [h.attachment('clanek.docx', 'application/msword')], { body: 'Tukaj je popravljena različica članka Obletnica gasilskega društva.' });
  h.advance(6 * 60000); h.run('processSubmissions'); assert.equal(pendingRows(h).length, 1);
  h.confirm('day10', label(h.rows()[1]));
  assert.equal(h.rows()[1][11], 2); assert.match(h.rows()[1][13], /Prejeta popravljena različica/);
  assert(h.articleFolders()[0].files.some(f => f.name === 'day1_001_clanek.docx'));
  assert(h.articleFolders()[0].files.some(f => f.name === 'day10_001_clanek.docx'));
  const files = h.state.files.length; const rows = h.rows().length;
  h.advance(6 * 60000); h.run('processSubmissions'); assert.equal(h.state.files.length, files); assert.equal(h.rows().length, rows);
  assert.deepEqual(labels, ['INBOX', 'UNREAD']);
  console.log('USPEH: celoten potek članek → ločene fotografije → poznejši popravek, potrditev, dnevnik, začasne mape, brez podvojitev ali prepisovanja.');
}
{
  const h = harness(); h.run('setup');
  h.message('base', 'Fire brigade anniversary', [h.attachment('article.docx', 'application/msword')], { sender: 'Ana <ana@brigade.example>' });
  h.advance(6 * 60000); h.run('processSubmissions');
  h.sheet('Prejeti članki').getRange(2, 8).setValues([['Gasilsko društvo Lipa']]);
  h.sheet('Sodelavci').getRange(2, 1, 1, 2).setValues([['photos@brigade.example', 'Gasilsko društvo Lipa']]);
  h.message('different', 'Material', [h.attachment('p.jpg', 'image/jpeg')], { sender: 'Bojan <photos@brigade.example>', body: 'Here are the photos for our fire brigade anniversary article.' });
  h.advance(6 * 60000); h.run('processSubmissions');
  assert(pendingRows(h)[0][6].includes('ista organizacija')); assert(pendingRows(h)[0][5].includes('Fire brigade anniversary'));
  h.confirm('different', label(h.rows()[1])); assert(h.rows()[1][8].includes('photos@brigade.example'));
  h.message('other-new', 'Povsem nov izlet v gore', [h.attachment('new.docx', 'application/msword')]);
  h.advance(6 * 60000); h.run('processSubmissions');
  // Znani avtor ima lahko nov članek: urednik ima izrecno možnost ustvarjanja.
  // Drugi avtor brez podobnosti je samodejno ustvarjen kot nov članek.
  assert.equal(h.rows().length, 3);
  h.message('known-new', 'Nov prispevek o mladini', [h.attachment('new.docx', 'application/msword')], { sender: 'Ana <ana@brigade.example>' });
  h.advance(6 * 60000); h.run('processSubmissions'); assert.equal(pendingRows(h).length, 1);
  h.confirm('known-new', '[NOVO] Ustvari nov članek'); assert.equal(h.rows().length, 4); assert.equal(pendingRows(h).length, 0);
  console.log('USPEH: drugačen pošiljatelj, organizacija, naravno besedilo, učenje naslova sodelavca in možnost novega članka.');
}
{
  const h = harness(); h.run('setup');
  h.message('original', 'Tradicija kraja', [h.attachment('a.docx', 'application/msword')]); h.advance(6 * 60000); h.run('processSubmissions');
  h.message('retry', 'Slike', [h.attachment('one.jpg', 'image/jpeg'), h.attachment('two.jpg', 'image/jpeg')]);
  h.advance(6 * 60000); h.run('processSubmissions');
  h.state.failMoveName = 'two.jpg'; h.confirm('retry', label(h.rows()[1]));
  assert.equal(log(h, 'retry')[12], 'Povezovanje'); assert.equal(pendingRows(h).length, 1); assert.equal(pendingRows(h)[0][9], false);
  assert.match(pendingRows(h)[0][10], /Napaka/);
  h.advance(6 * 60000); h.run('processSubmissions');
  assert.equal(log(h, 'retry')[12], 'Povezano'); assert.equal(pendingRows(h).length, 0); assert.equal(h.rows()[1][10], 2);
  assert.equal(h.articleFolders()[0].files.filter(f => f.mimeType === 'image/jpeg').length, 2);
  console.log('USPEH: prekinjen premik se varno nadaljuje brez dvojnih slik; vrstica izgine šele po uspehu.');
}
{
  const h = harness(); h.run('setup');
  h.message('partial', 'Članek', [h.attachment('a.docx', 'application/msword')]); h.state.failName = 'partial_001_a.docx';
  h.advance(6 * 60000); assert.throws(() => h.run('processSubmissions'), /opravil/);
  assert.equal(log(h, 'partial')[12], 'Priprava'); assert.equal(h.rows().length, 1);
  const staged = h.state.files.length;
  h.advance(6 * 60000); h.run('processSubmissions'); assert.equal(h.rows().length, 2); assert.equal(h.state.files.length, staged + 1);
  h.message('unknown', 'Fotografije brez naslova', [h.attachment('x.jpg', 'image/jpeg')], { sender: 'Nekdo <nobody@example.org>' });
  h.advance(6 * 60000); h.run('processSubmissions'); assert.equal(pendingRows(h).length, 1);
  h.confirm('unknown', ''); assert.equal(pendingRows(h).length, 1); assert.match(pendingRows(h)[0][10], /Najprej izberite/);
  h.confirm('unknown', '[NOVO] Ustvari nov članek'); assert.equal(pendingRows(h).length, 1); assert.equal(h.rows().length, 2);
  assert.match(pendingRows(h)[0][10], /samo iz priponke \.doc ali \.docx/);
  h.confirm('unknown', label(h.rows()[1])); assert.equal(pendingRows(h).length, 0); assert.equal(h.rows().length, 2);
  console.log('USPEH: prekinjeno shranjevanje, neznane fotografije, manjkajoča izbira in zavrnitev novega članka brez Wordove datoteke.');
}
{
  const h = harness({ pageSize: 1 }); h.run('setup');
  // API lahko najprej vrne fotografije; zaporedje obdelave sledi dejanskemu prejemu.
  h.message('later', 'Fotografije', [h.attachment('p.jpg', 'image/jpeg')], { date: Date.parse('2026-10-15T08:03:00Z') });
  h.message('earlier', 'Naše gasilsko društvo', [h.attachment('a.docx', 'application/msword')], { date: Date.parse('2026-10-15T08:02:00Z') });
  h.advance(6 * 60000); h.run('processSubmissions');
  assert.equal(h.rows().length, 2); assert(pendingRows(h)[0][5].includes('Naše gasilsko društvo')); assert.equal(h.state.listCalls, 2);
  const folder = h.articleFolders()[0];
  h.sheet('Prejeti članki').getRange(2, 5).setValues([['V prelomu']]);
  h.confirm('later', label(h.rows()[1])); assert.equal(h.rows()[1][4], 'V prelomu'); assert.equal(folder.files.length, 4);
  console.log('USPEH: strani rezultatov, kronološka obdelava in ohranitev uredniškega stanja.');
}
{
  const h = harness(); h.run('setup');
  const root = h.state.folders.find(f => f.name === 'Naše novine 2026');
  const folder = root.createFolder('Stari članek'); folder.createFile({ name: '001_foto.jpg', mimeType: 'image/jpeg', bytes: [1] });
  const articleSheet = h.sheet('Prejeti članki'); articleSheet.setName('Submissions');
  articleSheet.rows = [ ['Email subject', 'Sender', 'Received date', 'Drive folder link', 'Status', 'Gmail message ID'],
    ['Stari članek', 'Avtor <legacy@example.com>', new Date('2026-10-01T12:00:00Z'), folder.getUrl(), 'Received', 'legacy-id'] ];
  h.run('setup'); assert.equal(h.rows().length, 2); assert.equal(h.rows()[1][6], 'NN-legacy-id'); assert.equal(h.rows()[1][4], 'Prejeto'); assert.equal(h.rows()[1][10], 1);
  assert.equal(log(h, 'legacy-id')[12], 'Povezano pred nadgradnjo'); assert.equal(h.state.triggers.length, 2);
  h.run('setup'); assert.equal(h.rows().length, 2);
  console.log('USPEH: nadgradnja prvotne angleške evidence v slovenščino brez podvojitve članka ali sprožilcev.');
}
{
  const h = harness(); h.run('setup');
  h.message('review-base', 'Sto let našega društva', [h.attachment('clanek.docx', 'application/msword')]);
  h.advance(6 * 60000); h.run('processSubmissions');
  setStatus(h, 2, 'pripravljeno za lekturo');
  assert.equal(h.state.sent.length, 1); assert.equal(h.state.sent[0].to, 'uros.jures@gmail.com');
  assert.equal(h.state.sent[0].subject, '[NN-review-base][L1] Sto let našega društva');
  assert.equal(h.state.sent[0].attachments[0].name, 'NN-review-base_L1.docx');
  assert.match(h.state.sent[0].body, /NN-review-base_L1_lekturirano\.docx/);
  assert.equal(h.rows()[1][4], 'V lekturi');
  setStatus(h, 2, 'pripravljeno za lekturo'); assert.equal(h.state.sent.length, 1);
  h.message('review-return', 'Re: ' + h.state.sent[0].subject, [h.attachment('NN-review-base_L1_lekturirano.docx', 'application/msword')], { sender: 'Uroš <uros.jures@gmail.com>' });
  h.advance(6 * 60000); h.run('processSubmissions');
  assert.equal(h.rows()[1][4], 'Lekturirano'); assert.equal(pendingRows(h).length, 0); assert.equal(h.rows().length, 2);
  assert.equal(h.rows()[1][11], 2); assert.equal(h.sheet('Lekture').rows[1][9], 'Lekturirano');
  assert.equal(log(h, 'review-return')[18], 'Lekturirana različica');
  assert(h.articleFolders()[0].files.some(f => f.name === 'review-base_001_clanek.docx'));
  assert(h.articleFolders()[0].files.some(f => f.name.includes('_lekturirano.docx')));
  console.log('USPEH: izbira stanja pošlje točen dokument privzetemu lektorju, zadeva vsebuje ID in krog; pravilen odgovor se poveže in označi Lekturirano.');
}
{
  const h = harness(); h.run('setup');
  h.message('revision-base', 'Razvoj vasi', [h.attachment('clanek.docx', 'application/msword')]);
  h.advance(6 * 60000); h.run('processSubmissions'); setStatus(h, 2, 'pripravljeno za lekturo');
  const subject1 = h.state.sent[0].subject;
  h.message('correction', 'Popravek', [h.attachment('nov.docx', 'application/msword')]); h.advance(6 * 60000); h.run('processSubmissions');
  h.confirm('correction', label(h.rows()[1]));
  assert.equal(h.rows()[1][4], 'Potrebna nova lektura'); const newDoc = h.rows()[1][16];
  h.message('old-review', 'Re: ' + subject1, [h.attachment('vrnjeno_lekturirano.docx', 'application/msword')], { sender: 'Uroš <uros.jures@gmail.com>' });
  h.advance(6 * 60000); h.run('processSubmissions');
  assert.equal(h.rows()[1][4], 'Potrebna nova lektura'); assert.equal(h.rows()[1][16], newDoc);
  assert.equal(log(h, 'old-review')[18], 'Lektura prejšnje različice');
  setStatus(h, 2, 'pripravljeno za lekturo'); assert.equal(h.state.sent.length, 2); assert.match(h.state.sent[1].subject, /\[L2\]/);
  h.message('new-review', 'Re: ' + h.state.sent[1].subject, [h.attachment('vrnjeno_lekturirano.docx', 'application/msword')], { sender: 'Uroš <uros.jures@gmail.com>' });
  h.advance(6 * 60000); h.run('processSubmissions'); assert.equal(h.rows()[1][4], 'Lekturirano');
  console.log('USPEH: starejša lektura ne potrdi novega popravka; druga različica se pošlje v drugi krog.');
}
{
  const h = harness(); h.run('setup');
  h.message('guard-base', 'Kulturni večer', [h.attachment('a.docx', 'application/msword')]); h.advance(6 * 60000); h.run('processSubmissions');
  setStatus(h, 2, 'pripravljeno za lekturo');
  h.message('wrong-sender', h.state.sent[0].subject, [h.attachment('a_lekturirano.docx', 'application/msword')], { sender: 'Nekdo <other@example.net>' });
  h.message('wrong-name', h.state.sent[0].subject, [h.attachment('a_final.docx', 'application/msword')], { sender: 'Uroš <uros.jures@gmail.com>' });
  h.advance(6 * 60000); h.run('processSubmissions'); assert.equal(pendingRows(h).length, 2); assert.equal(h.rows()[1][4], 'V lekturi');
  assert(pendingRows(h).every(r => r[6].includes('ID članka')));
  console.log('USPEH: napačen pošiljatelj ali ime datoteke zahtevata urednikovo potrditev.');
}
{
  const h = harness(); h.run('setup');
  h.message('send-guard', 'Zgodovina šole', [h.attachment('a.docx', 'application/msword')]); h.advance(6 * 60000); h.run('processSubmissions');
  h.state.failSend = true; setStatus(h, 2, 'pripravljeno za lekturo');
  assert.equal(h.state.sent.length, 1); assert.equal(h.rows()[1][4], 'Preverite pošiljanje');
  setStatus(h, 2, 'pripravljeno za lekturo'); assert.equal(h.state.sent.length, 1);
  h.state.failSend = false; h.sheet('Lekture').getRange(2, 10).setValues([['Poslano']]);
  setStatus(h, 2, 'pripravljeno za lekturo'); assert.equal(h.state.sent.length, 1); assert.equal(h.rows()[1][4], 'V lekturi');
  console.log('USPEH: nejasen izid pošiljanja ne sproži podvojenega sporočila.');
}
{
  const h = harness(); h.run('setup');
  h.message('self-base', 'Lokalna zgodovina', [h.attachment('a.docx', 'application/msword')]);
  h.advance(6 * 60000); h.run('processSubmissions');
  h.sheet('Prejeti članki').getRange(2, 16).setValues([['editor@example.com']]);
  setStatus(h, 2, 'pripravljeno za lekturo');
  h.message('self-review', h.state.sent[0].subject, [h.attachment('a_lekturirano.docx', 'application/msword')], { sender: 'editor@example.com', labels: ['SENT', 'INBOX'] });
  h.message('ordinary-sent', 'Poslan dokument', [h.attachment('x.pdf', 'application/pdf')], { sender: 'editor@example.com', to: 'someone@example.com', labels: ['SENT'] });
  h.advance(6 * 60000); h.run('processSubmissions');
  assert.equal(h.rows()[1][4], 'Lekturirano'); assert.equal(h.rows().length, 2); assert.equal(log(h, 'ordinary-sent'), undefined);
  console.log('USPEH: lektor lahko uporablja isti račun kot uredništvo; običajna poslana pošta se ne uvozi.');
}
{
  const h = harness(); h.run('setup');
  h.message('timing-base', 'Življenje v občini', [h.attachment('a.docx', 'application/msword')]); h.advance(6 * 60000); h.run('processSubmissions');
  h.message('timing-correction', 'Popravek', [h.attachment('nov.docx', 'application/msword')], { date: Date.parse('2026-10-15T08:07:00Z') });
  h.message('timing-images', 'Fotografije', [h.attachment('p.jpg', 'image/jpeg')], { date: Date.parse('2026-10-15T08:08:00Z') });
  h.advance(6 * 60000); h.run('processSubmissions');
  h.confirm('timing-images', label(h.rows()[1])); h.confirm('timing-correction', label(h.rows()[1]));
  assert(h.rows()[1][16].includes('timing-correction_001_nov.docx'));
  console.log('USPEH: urednik lahko popravek in fotografije potrdi v poljubnem vrstnem redu; izbrani dokument ostane pravilen.');
}
{
  const h = harness(); h.run('setup');
  h.message('ignore-pdf', 'Samostojna izdaja', [h.attachment('spis.pdf', 'application/pdf')], { sender: 'pdf@example.net' });
  h.message('ignore-docm', 'Druga izdaja', [h.attachment('spis.docm', 'application/vnd.ms-word.document.macroenabled.12')], { sender: 'docm@example.net' });
  h.message('unmatched-jpg', 'Letno poročilo', [h.attachment('spis.jpg', 'image/jpeg')], { sender: 'jpg@example.net' });
  h.message('ignore-double-ext', 'Podatkovni spis', [h.attachment('spis.docx.exe', 'application/msword')], { sender: 'exe@example.net' });
  h.message('accept-doc', 'Dogajanje v kraju', [h.attachment('clanek.DOC', 'application/octet-stream')], { sender: 'doc@example.net' });
  h.message('accept-docx', 'Obnova šolske stavbe', [h.attachment('clanek.DOCX', 'application/octet-stream'), h.attachment('fotografija.jpg', 'image/jpeg'), h.attachment('priloga.pdf', 'application/pdf')], { sender: 'docx@example.net' });
  h.advance(6 * 60000); h.run('processSubmissions');
  assert.equal(h.rows().length, 3); assert.equal(pendingRows(h).length, 1);
  assert.equal(pendingRows(h)[0][11], 'unmatched-jpg');
  assert.deepEqual(h.rows().slice(1).map(r => r[5]), ['accept-doc', 'accept-docx']);
  for (const id of ['ignore-pdf', 'ignore-docm', 'ignore-double-ext']) {
    assert.equal(log(h, id), undefined); assert(!h.state.files.some(f => f.name.startsWith(id + '_')));
  }
  assert.equal(h.rows()[2][10], 1); assert.equal(h.rows()[2][11], 2);
  h.confirm('unmatched-jpg', label(h.rows()[2]));
  assert.equal(pendingRows(h).length, 0); assert.equal(h.rows()[2][10], 2);
  h.message('followup-pdf', '[' + h.rows()[2][6] + '] Dodatno gradivo', [h.attachment('dokazilo.pdf', 'application/pdf')], { sender: 'additional@example.net' });
  h.advance(6 * 60000); h.run('processSubmissions');
  assert.equal(pendingRows(h).length, 1);
  h.confirm('followup-pdf', '[NOVO] Ustvari nov članek'); assert.equal(h.rows().length, 3); assert.equal(pendingRows(h).length, 1);
  h.confirm('followup-pdf', label(h.rows()[2])); assert.equal(pendingRows(h).length, 0); assert.equal(h.rows()[2][11], 3);
  console.log('USPEH: nova članka samo iz DOC/DOCX, samostojni PDF/DOCM se preskočijo, neprepoznane slike pa čakajo na povezavo; pripadajoče slike/PDF in poznejše dopolnitve se ohranijo.');
}
{
  const h = harness(); h.run('setup');
  h.message('unrecognized-images', 'Pozdrav', [h.attachment('001.jpg', 'image/jpeg'), h.attachment('002.PNG', 'application/octet-stream')], {
    sender: 'Neznani avtor <unknown@different.example>', body: 'Dober dan.'
  });
  h.advance(6 * 60000); h.run('processSubmissions');
  assert.equal(h.rows().length, 1); assert.equal(h.articleFolders().length, 0);
  assert.equal(pendingRows(h).length, 1); assert.equal(pendingRows(h)[0][5], '');
  assert.equal(pendingRows(h)[0][8], ''); assert.match(pendingRows(h)[0][6], /Slik ni bilo mogoče povezati/);
  assert.equal(log(h, 'unrecognized-images')[12], 'Čaka na izbiro');
  const temp = h.state.folders.find(folder => folder.id === log(h, 'unrecognized-images')[10]);
  assert.equal(temp.files.length, 3); assert(!temp.trashed);
  const fileCount = h.state.files.length;
  h.advance(6 * 60000); h.run('processSubmissions');
  assert.equal(pendingRows(h).length, 1); assert.equal(h.state.files.length, fileCount);
  h.message('later-article', 'Srečanje krajanov', [h.attachment('clanek.docx', 'application/msword')]);
  h.advance(6 * 60000); h.run('processSubmissions');
  h.confirm('unrecognized-images', label(h.rows()[1]));
  assert.equal(pendingRows(h).length, 0); assert.equal(h.rows().length, 2);
  assert.equal(h.rows()[1][10], 2); assert.equal(log(h, 'unrecognized-images')[12], 'Povezano');
  assert(temp.trashed);
  console.log('USPEH: slike brez znanega pošiljatelja, naslova, namiga ali obstoječega članka ostanejo v čakalnici brez podvojitev in se po urednikovi izbiri povežejo.');
}

