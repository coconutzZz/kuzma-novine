/**
 * Naše novine 2026: Gmail -> article folders in Drive -> Google Sheets.
 * Run setup() once. Gmail is accessed with read-only permission.
 * Article name = email subject. One email = one submission.
 */
const CONFIG = Object.freeze({
  MAIN_FOLDER_NAME: 'Naše novine 2026',
  REGISTER_NAME: 'Naše novine 2026 - Evidencija',
  SHEET_NAME: 'Submissions',

  // Optional: IDs of an existing My Drive folder / spreadsheet.
  // Leave blank to create them automatically. Set before running setup().
  MAIN_FOLDER_ID: '',
  SPREADSHEET_ID: '',

  // Optional restriction, e.g. 'to:redakcija@example.com' or 'label:Articles'.
  // Empty = all incoming messages with attachments, excluding spam/trash/sent.
  EXTRA_GMAIL_QUERY: '',

  // Empty = new mail from installation onward. To import older mail, set an
  // ISO date with timezone BEFORE setup(), e.g. '2026-01-01T00:00:00+01:00'.
  START_FROM: '',
  INCLUDE_INLINE_IMAGES: true, // Also saves embedded image parts, e.g. logos.
  CHECK_EVERY_MINUTES: 5,
  MAX_RUN_MS: 240000,
  OVERLAP_SECONDS: 86400 // Recheck a day for delayed Gmail search indexing.
});

const HEADERS = [
  'Email subject', 'Sender', 'Received date', 'Drive folder link',
  'Status', 'Gmail message ID'
];

/** Create/reuse the main folder and register, and install the timer. */
function setup() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const props = PropertiesService.getScriptProperties();
    const profile = Gmail.Users.getProfile('me'); // Check API/authorization.
    const mailbox = props.getProperty('NN_MAILBOX');
    if (mailbox && mailbox !== profile.emailAddress) {
      throw new Error('Use the same Gmail account that originally ran setup().');
    }

    let start = props.getProperty('NN_START');
    if (!start) {
      const millis = CONFIG.START_FROM ? Date.parse(CONFIG.START_FROM) : Date.now();
      if (!Number.isFinite(millis) || millis > Date.now()) {
        throw new Error('START_FROM must be a valid date in the past, with timezone.');
      }
      start = String(Math.floor(millis / 1000));
    }

    let folderId = CONFIG.MAIN_FOLDER_ID || props.getProperty('NN_FOLDER_ID');
    if (!folderId) {
      // Reuse one existing root-level folder with this exact name, if present.
      const matches = DriveApp.getRootFolder().getFoldersByName(CONFIG.MAIN_FOLDER_NAME);
      let folder = matches.hasNext() ? matches.next() : null;
      if (matches.hasNext()) {
        throw new Error('Multiple main folders found. Set MAIN_FOLDER_ID explicitly.');
      }
      folder = folder || DriveApp.createFolder(CONFIG.MAIN_FOLDER_NAME);
      folderId = folder.getId();
      props.setProperty('NN_FOLDER_ID', folderId);
    }
    const mainFolder = DriveApp.getFolderById(folderId);

    let bookId = CONFIG.SPREADSHEET_ID || props.getProperty('NN_BOOK_ID');
    let book;
    if (bookId) {
      book = SpreadsheetApp.openById(bookId);
    } else {
      book = SpreadsheetApp.create(CONFIG.REGISTER_NAME);
      bookId = book.getId();
      props.setProperty('NN_BOOK_ID', bookId);
      book.getSheets()[0].setName(CONFIG.SHEET_NAME);
    }
    const sheet = book.getSheetByName(CONFIG.SHEET_NAME) || book.insertSheet(CONFIG.SHEET_NAME);
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    }
    checkHeaders_(sheet);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setColumnWidth(1, 320);
    sheet.setColumnWidth(2, 270);
    sheet.setColumnWidth(3, 165);
    sheet.setColumnWidth(4, 340);
    sheet.setColumnWidth(5, 110);
    sheet.hideColumns(6); // Keep IDs intact: they prevent duplicate rows.
    book.setSpreadsheetTimeZone(Session.getScriptTimeZone());

    props.setProperties({
      NN_FOLDER_ID: folderId,
      NN_BOOK_ID: bookId,
      NN_MAILBOX: profile.emailAddress,
      NN_START: start
    });
    if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'processSubmissions')) {
      ScriptApp.newTrigger('processSubmissions').timeBased()
        .everyMinutes(CONFIG.CHECK_EVERY_MINUTES).create();
    }
    console.log('Main folder: ' + mainFolder.getUrl());
    console.log('Submission register: ' + book.getUrl());
    console.log('Ready. Incoming submissions will be checked every five minutes.');
  } finally {
    lock.releaseLock();
  }
}

/** Timer entry point. Also safe to run manually after setup(). */
function processSubmissions() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    const deadline = Date.now() + CONFIG.MAX_RUN_MS;
    const props = PropertiesService.getScriptProperties();
    const start = Number(props.getProperty('NN_START'));
    if (!start || !props.getProperty('NN_BOOK_ID') || !props.getProperty('NN_FOLDER_ID')) {
      throw new Error('Run setup() first.');
    }
    const mainFolder = DriveApp.getFolderById(props.getProperty('NN_FOLDER_ID'));
    const sheet = SpreadsheetApp.openById(props.getProperty('NN_BOOK_ID'))
      .getSheetByName(CONFIG.SHEET_NAME);
    checkHeaders_(sheet);
    const count = sheet.getLastRow() - 1;
    const registered = new Set(count > 0
      ? sheet.getRange(2, 6, count, 1).getDisplayValues().map(row => row[0]) : []);

    const checkpoint = Number(props.getProperty('NN_CHECKPOINT') || start);
    const after = Math.max(start - 1, checkpoint - CONFIG.OVERLAP_SECONDS);
    const before = Math.floor(Date.now() / 1000) - 60;
    if (before <= start) return;
    const query = [
      'has:attachment -in:sent -in:drafts -in:spam -in:trash',
      'after:' + after, 'before:' + before, CONFIG.EXTRA_GMAIL_QUERY
    ].filter(Boolean).join(' ');

    let pageToken;
    let failures = 0;
    let saved = 0;
    do {
      requireTime_(deadline);
      const options = { q: query, maxResults: 100, includeSpamTrash: false };
      if (pageToken) options.pageToken = pageToken;
      const page = Gmail.Users.Messages.list('me', options);
      for (const item of page.messages || []) {
        requireTime_(deadline);
        if (registered.has(item.id)) continue;
        try {
          const message = Gmail.Users.Messages.get('me', item.id, { format: 'full' });
          if ((message.labelIds || []).some(label =>
            ['SENT', 'DRAFT', 'SPAM', 'TRASH'].includes(label))) continue;
          const received = new Date(Number(message.internalDate));
          if (received.getTime() < start * 1000 || received.getTime() >= before * 1000) continue;
          const attachments = collectAttachments_(message.payload);
          if (attachments.length === 0) continue;

          const subject = decodeHeader_(header_(message.payload, 'Subject')) || '(no subject)';
          const sender = decodeHeader_(header_(message.payload, 'From'));
          const folder = articleFolder_(mainFolder, subject, message.id);
          for (let i = 0; i < attachments.length; i++) {
            requireTime_(deadline);
            const part = attachments[i];
            // Numbering preserves two attachments with identical filenames,
            // and lets retries skip files already copied successfully.
            const name = String(i + 1).padStart(3, '0') + '_' + cleanName_(part.filename);
            if (folder.getFilesByName(name).hasNext()) continue;
            const body = part.body || {};
            const data = body.attachmentId
              ? Gmail.Users.Messages.Attachments.get('me', message.id, body.attachmentId).data
              : body.data;
            if (data === undefined || data === null) throw new Error('Attachment data missing: ' + name);
            const bytes = Utilities.base64DecodeWebSafe(data);
            folder.createFile(Utilities.newBlob(bytes, part.mimeType || 'application/octet-stream', name));
          }

          const row = sheet.getLastRow() + 1;
          if (row > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 100);
          sheet.getRange(row, 1, 1, 2).setNumberFormat('@');
          sheet.getRange(row, 3).setNumberFormat('yyyy-mm-dd hh:mm:ss');
          sheet.getRange(row, 6).setNumberFormat('@');
          sheet.getRange(row, 1, 1, HEADERS.length).setValues([[
            safeCell_(subject), safeCell_(sender), received, folder.getUrl(),
            'Received', message.id
          ]]);
          SpreadsheetApp.flush();
          registered.add(message.id);
          saved++;
        } catch (error) {
          failures++;
          console.error('Message ' + item.id + ': ' + String(error));
        }
      }
      pageToken = page.nextPageToken;
    } while (pageToken);

    // Advance only after every page and file succeeded. Errors are retried on
    // the next run; a row is written only after ALL supported files are saved.
    if (failures) throw new Error(failures + ' submission(s) failed; next run will retry. See execution log.');
    props.setProperty('NN_CHECKPOINT', String(before));
    console.log('Registered ' + saved + ' new submission(s).');
  } finally {
    lock.releaseLock();
  }
}

/** Exact article title normally; append the message ID on a title collision. */
function articleFolder_(mainFolder, subject, messageId) {
  const articleName = cleanName_(subject).slice(0, 180) || 'Untitled article';
  const marker = 'Gmail submission ID: ' + messageId;
  const baseMatches = mainFolder.getFoldersByName(articleName);
  let collision = false;
  while (baseMatches.hasNext()) {
    const folder = baseMatches.next();
    if (folder.getDescription() === marker) return folder;
    collision = true;
  }
  const name = collision ? articleName + ' [' + messageId + ']' : articleName;
  const matches = mainFolder.getFoldersByName(name);
  while (matches.hasNext()) {
    const folder = matches.next();
    if (folder.getDescription() === marker) return folder;
  }
  return mainFolder.createFolder(name).setDescription(marker);
}

/** Traverse MIME parts; copy images, Word files, and PDFs without conversion. */
function collectAttachments_(payload) {
  const result = [];
  function visit(part) {
    if (!part) return;
    const type = (part.mimeType || '').toLowerCase().split(';')[0].trim();
    const mimeExtensions = {
      'application/pdf': 'pdf', 'application/msword': 'doc',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'docx',
      'application/vnd.ms-word.document.macroenabled.12': 'docm'
    };
    const body = part.body || {};
    const knownMime = type.startsWith('image/') || Boolean(mimeExtensions[type]);
    if (part.filename || (knownMime && (body.attachmentId || body.data !== undefined))) {
      const filename = part.filename || 'attachment.' + (mimeExtensions[type] || type.split('/')[1]);
      const disposition = header_(part, 'Content-Disposition');
      const inline = /^inline\b/i.test(disposition) ||
        (!/^attachment\b/i.test(disposition) && Boolean(header_(part, 'Content-ID')));
      const image = type.startsWith('image/') || /\.(jpe?g|png|gif|webp|heic|heif|tiff?|bmp|avif|svg)$/i.test(filename);
      const document = Boolean(mimeExtensions[type]) || /\.(pdf|doc|docx|docm)$/i.test(filename);
      if ((image || document) && (!inline || CONFIG.INCLUDE_INLINE_IMAGES)) {
        result.push(Object.assign({}, part, { filename: filename }));
      }
      return; // Do not unpack ZIP files or attached .eml messages.
    }
    (part.parts || []).forEach(visit);
  }
  visit(payload);
  return result;
}

function header_(part, name) {
  const header = ((part && part.headers) || []).find(h => h.name.toLowerCase() === name.toLowerCase());
  return header ? header.value : '';
}

/** Decode RFC 2047 subjects/display names, including Croatian diacritics. */
function decodeHeader_(value) {
  return String(value).replace(/(\?=)[ \t\r\n]+(?==\?)/g, '$1')
    .replace(/=\?([^?]+)\?([bq])\?([^?]*)\?=/gi, (original, charset, encoding, encoded) => {
      try {
        let bytes;
        if (encoding.toLowerCase() === 'b') {
          bytes = Utilities.base64Decode(encoded);
        } else {
          bytes = [];
          const text = encoded.replace(/_/g, ' ');
          for (let i = 0; i < text.length; i++) {
            if (text[i] === '=' && /^[0-9a-f]{2}$/i.test(text.slice(i + 1, i + 3))) {
              bytes.push(parseInt(text.slice(i + 1, i + 3), 16));
              i += 2;
            } else {
              bytes.push(text.charCodeAt(i));
            }
          }
          bytes = bytes.map(b => b > 127 ? b - 256 : b);
        }
        return Utilities.newBlob(bytes).getDataAsString(charset);
      } catch (error) {
        return original;
      }
    });
}

function cleanName_(value) {
  return String(value).replace(/[\\/\u0000-\u001f\u007f]/g, ' ').trim() || 'unnamed';
}

function safeCell_(value) {
  // Treat untrusted email headers as text, never as spreadsheet formulas.
  const text = String(value);
  return /^[\s]*[=+@-]/.test(text) ? "'" + text : text;
}

function checkHeaders_(sheet) {
  if (!sheet || sheet.getRange(1, 1, 1, HEADERS.length).getDisplayValues()[0]
    .some((value, i) => value !== HEADERS[i])) {
    throw new Error('Register headers are missing or changed. Restore columns A:F, including the hidden message ID.');
  }
}

function requireTime_(deadline) {
  if (Date.now() >= deadline) {
    throw new Error('Time budget reached. Progress is saved; the next run will retry remaining work.');
  }
}

/** Optional: stop the timer. Drive files, emails, and register rows remain. */
function stopAutomation() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'processSubmissions')
    .forEach(t => ScriptApp.deleteTrigger(t));
  console.log('Automatic checking stopped. Run setup() to start again.');
}
