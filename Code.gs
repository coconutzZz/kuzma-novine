/**
 * Naše novine 2026: prejem člankov in ročna povezava poznejših dopolnitev.
 * Enkrat zaženite setup(). Obstoječo pošto beremo; MailApp pošilja gradivo lektorju.
 */
const CONFIG = Object.freeze({
  MAIN_FOLDER_NAME: 'Naše novine 2026',
  REGISTER_NAME: 'Naše novine 2026 - evidenca',
  SHEET_NAME: 'Prejeti članki',
  PENDING_SHEET_NAME: 'Neujemajoči maili',
  LOG_SHEET_NAME: 'Dnevnik pošte',
  CONTACTS_SHEET_NAME: 'Sodelavci',
  CHOICES_SHEET_NAME: '_Izbira člankov',
  STAGING_FOLDER_NAME: '_Začasna pošta',
  REVIEW_SHEET_NAME: 'Lekture',
  DEFAULT_REVIEWER: 'uros.jures@gmail.com',
  // Prazno pomeni samodejno ustvarjanje; pri nadgradnji ostanejo shranjeni ID-ji.
  MAIN_FOLDER_ID: '',
  SPREADSHEET_ID: '',
  // Primer: 'to:urednistvo@example.com'. Prazno: vsa prejeta pošta s priponkami.
  EXTRA_GMAIL_QUERY: '',
  // Nastavite pred prvo namestitvijo, če želite uvoziti tudi starejšo pošto.
  START_FROM: '',
  INCLUDE_INLINE_IMAGES: true,
  CHECK_EVERY_MINUTES: 5,
  MAX_RUN_MS: 240000,
  OVERLAP_SECONDS: 86400,
  MIN_MATCH_SCORE: 25
});

const A_KEYS = ['subject', 'sender', 'received', 'folderUrl', 'status', 'sourceId',
  'id', 'org', 'emails', 'title', 'images', 'documents', 'last', 'update', 'folderId',
  'reviewer', 'reviewDoc', 'reviewRound', 'reviewSent', 'reviewSubject', 'reviewStatus', 'currentReview', 'materialId'];
const A_HEADERS = ['Zadeva prvega e-poštnega sporočila', 'Pošiljatelj', 'Datum prejema',
  'Mapa članka', 'Stanje', 'ID prvega sporočila', 'ID članka', 'Organizacija',
  'E-naslovi sodelavcev', 'Naslov članka', 'Število slik', 'Število dokumentov',
  'Zadnje prejeto gradivo', 'Zadnja dopolnitev', 'ID mape', 'Lektor', 'Dokument za lekturo',
  'Krog lekture', 'Poslano v lekturo', 'Zadeva za lekturo', 'Obvestilo o lekturi', 'ID lekture', 'ID zadnjega gradiva'];
const L_KEYS = ['id', 'thread', 'internetId', 'sender', 'email', 'org', 'subject',
  'hint', 'received', 'preview', 'staging', 'files', 'state', 'target', 'reason',
  'linked', 'error', 'body', 'kind', 'reviewId'];
const L_HEADERS = ['ID sporočila', 'ID pogovora', 'Internetni ID sporočila',
  'Pošiljatelj', 'E-naslov pošiljatelja', 'Organizacija', 'Zadeva',
  'Razpoznani naslov članka', 'Datum prejema', 'Predogled sporočila',
  'ID začasne mape', 'Seznam datotek', 'Stanje povezave', 'ID članka',
  'Razlog povezave ali predloga', 'Datum povezave', 'Zadnja napaka', 'Besedilo za primerjavo', 'Vrsta gradiva', 'ID lekture'];
const P_KEYS = ['received', 'sender', 'org', 'subject', 'preview', 'suggestions',
  'reason', 'stagingUrl', 'selection', 'confirm', 'error', 'id', 'kind'];
const P_HEADERS = ['Datum prejema', 'Pošiljatelj', 'Organizacija', 'Zadeva',
  'Predogled sporočila', 'Predlagani članki', 'Razlog za predlog', 'Začasna mapa',
  'Izbrani članek', 'Potrdi povezavo', 'Napaka / navodilo', 'ID sporočila', 'Vrsta gradiva'];
const R_KEYS = ['id', 'article', 'round', 'recipient', 'subject', 'docId', 'docName',
  'created', 'sent', 'state', 'responseId', 'returned', 'error'];
const R_HEADERS = ['ID lekture', 'ID članka', 'Krog lekture', 'E-naslov lektorja', 'Zadeva',
  'ID poslanega dokumenta', 'Ime poslanega dokumenta', 'Pripravljeno', 'Poslano',
  'Stanje pošiljanja', 'ID odgovora', 'Prejeto od lektorja', 'Napaka / navodilo'];
const MATERIAL = Object.freeze({ NORMAL: 'Dopolnitev', REVIEWED: 'Lekturirana različica', OLD_REVIEW: 'Lektura prejšnje različice' });
const ARTICLE_STATUSES = ['Prejeto', 'Čaka na slike', 'pripravljeno za lekturo', 'V lekturi',
  'Lekturirano', 'Potrebna nova lektura', 'V prelomu', 'Objavljeno', 'Prenos v teku', 'Napaka pri lekturi', 'Preverite pošiljanje'];
const STATE = Object.freeze({ STAGING: 'Priprava', READY: 'Pripravljeno',
  PENDING: 'Čaka na izbiro', LINKING: 'Povezovanje', LINKED: 'Povezano', LEGACY: 'Povezano pred nadgradnjo' });
const NEW_ARTICLE = '[NOVO] Ustvari nov članek';

/** Namesti ali nadgradi evidenco, ne da bi ponovno uvažal obstoječe članke. */
function setup() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const props = PropertiesService.getScriptProperties();
    const profile = Gmail.Users.getProfile('me');
    if (props.getProperty('NN_MAILBOX') && props.getProperty('NN_MAILBOX') !== profile.emailAddress) {
      throw new Error('Namestitev zaženite z istim Gmail računom kot prvotno.');
    }
    let start = props.getProperty('NN_START');
    if (!start) {
      const date = CONFIG.START_FROM ? Date.parse(CONFIG.START_FROM) : Date.now();
      if (!Number.isFinite(date) || date > Date.now()) throw new Error('START_FROM mora biti veljaven pretekli datum s časovnim pasom.');
      start = String(Math.floor(date / 1000));
    }
    let rootId = CONFIG.MAIN_FOLDER_ID || props.getProperty('NN_FOLDER_ID');
    if (!rootId) {
      const matches = DriveApp.getRootFolder().getFoldersByName(CONFIG.MAIN_FOLDER_NAME);
      const root = matches.hasNext() ? matches.next() : DriveApp.createFolder(CONFIG.MAIN_FOLDER_NAME);
      if (matches.hasNext()) throw new Error('Najdenih je več glavnih map. Vnesite MAIN_FOLDER_ID.');
      rootId = root.getId();
      props.setProperty('NN_FOLDER_ID', rootId);
    }
    const root = DriveApp.getFolderById(rootId);
    let bookId = CONFIG.SPREADSHEET_ID || props.getProperty('NN_BOOK_ID');
    let book;
    if (bookId) book = SpreadsheetApp.openById(bookId);
    else {
      book = SpreadsheetApp.create(CONFIG.REGISTER_NAME);
      bookId = book.getId();
      props.setProperty('NN_BOOK_ID', bookId);
      book.getSheets()[0].setName(CONFIG.SHEET_NAME);
    }
    book.rename(CONFIG.REGISTER_NAME);
    book.setSpreadsheetTimeZone(Session.getScriptTimeZone());
    const old = book.getSheetByName('Submissions');
    if (old && !book.getSheetByName(CONFIG.SHEET_NAME)) old.setName(CONFIG.SHEET_NAME);
    else if (old) throw new Error('Prisotna sta stari in novi zavihek člankov. Najprej združite podatke.');
    const ctx = sheets_(book, true);
    const stage = oneFolder_(root, CONFIG.STAGING_FOLDER_NAME);
    props.setProperties({ NN_FOLDER_ID: rootId, NN_BOOK_ID: bookId,
      NN_MAILBOX: profile.emailAddress, NN_START: start, NN_STAGING_ID: stage.getId() });
    migrateArticles_(ctx);
    refreshChoices_(ctx);
    articleControls_(ctx);
    const triggers = ScriptApp.getProjectTriggers();
    if (!triggers.some(t => t.getHandlerFunction() === 'processSubmissions')) {
      ScriptApp.newTrigger('processSubmissions').timeBased().everyMinutes(CONFIG.CHECK_EVERY_MINUTES).create();
    }
    if (!triggers.some(t => t.getHandlerFunction() === 'onPendingEdit')) {
      ScriptApp.newTrigger('onPendingEdit').forSpreadsheet(book).onEdit().create();
    }
    console.log('Glavna mapa: ' + root.getUrl());
    console.log('Evidenca: ' + book.getUrl());
    console.log('Namestitev je končana. Preverjanje pošte poteka vsakih pet minut.');
  } finally { lock.releaseLock(); }
}

/** Časovni sprožilec: prejem, začasno shranjevanje in predlogi povezav. */
function processSubmissions() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    const deadline = Date.now() + CONFIG.MAX_RUN_MS;
    const ctx = context_();
    let failures = recover_(ctx, deadline);
    const start = Number(ctx.props.getProperty('NN_START'));
    const checkpoint = Number(ctx.props.getProperty('NN_CHECKPOINT') || start);
    const before = Math.floor(Date.now() / 1000) - 60;
    if (before <= start) return;
    const after = Math.max(start - 1, checkpoint - CONFIG.OVERLAP_SECONDS);
    const q = ['has:attachment -in:drafts -in:spam -in:trash',
      'after:' + after, 'before:' + before, CONFIG.EXTRA_GMAIL_QUERY].filter(Boolean).join(' ');
    const known = new Set(records_(ctx.log, L_KEYS).map(r => r.id));
    const messages = [];
    let pageToken;
    do {
      requireTime_(deadline);
      const options = { q: q, maxResults: 100, includeSpamTrash: false };
      if (pageToken) options.pageToken = pageToken;
      const page = Gmail.Users.Messages.list('me', options);
      for (const item of page.messages || []) {
        requireTime_(deadline);
        if (known.has(item.id)) continue;
        const message = Gmail.Users.Messages.get('me', item.id, { format: 'full' });
        if ((message.labelIds || []).some(x => ['DRAFT', 'SPAM', 'TRASH'].includes(x))) continue;
        if ((message.labelIds || []).includes('SENT') && !selfReviewReturn_(ctx, message)) continue;
        const received = Number(message.internalDate);
        if (received >= start * 1000 && received < before * 1000 && collectAttachments_(message.payload).length) messages.push(message);
      }
      pageToken = page.nextPageToken;
    } while (pageToken);
    // Starejše sporočilo obravnavamo pred pozneje prispelimi fotografijami.
    messages.sort((a, b) => Number(a.internalDate) - Number(b.internalDate));
    for (const message of messages) {
      requireTime_(deadline);
      try {
        const mail = stageMessage_(ctx, message, deadline);
        routeMessage_(ctx, mail, deadline);
      } catch (error) {
        failures++;
        reportFailure_(ctx, message.id, error);
      }
    }
    refreshChoices_(ctx);
    processReviewQueue_(ctx, deadline);
    if (failures) throw new Error(failures + ' opravil ni bilo dokončanih. Naslednji zagon jih bo poskusil ponoviti.');
    ctx.props.setProperty('NN_CHECKPOINT', String(before));
    console.log('Obdelanih novih sporočil: ' + messages.length + '.');
  } finally { lock.releaseLock(); }
}

/** Namestljivi sprožilec spremembe: urednik potrdi vrstico s kljukico. */
function onPendingEdit(event) {
  if (!event || !event.range || !event.source) return;
  if (event.source.getId() !== PropertiesService.getScriptProperties().getProperty('NN_BOOK_ID')) return;
  const range = event.range;
  const name = range.getSheet().getName();
  if (range.getRow() < 2 || ![CONFIG.PENDING_SHEET_NAME, CONFIG.SHEET_NAME].includes(name)) return;
  const column = name === CONFIG.PENDING_SHEET_NAME ? 10 : 5;
  if (range.getColumn() > column || range.getLastColumn() < column) return;
  // Tudi pri sočasnih potrditvah obdelavo prevzame naslednji časovni zagon.
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    const ctx = context_();
    const deadline = Date.now() + CONFIG.MAX_RUN_MS;
    if (name === CONFIG.PENDING_SHEET_NAME) {
      const rows = records_(ctx.pending, P_KEYS).filter(r => r._row >= range.getRow() && r._row <= range.getLastRow() && r.confirm === true);
      for (const row of rows) confirmPending_(ctx, row.id, deadline);
    } else {
      const rows = records_(ctx.articles, A_KEYS).filter(r => r._row >= range.getRow() && r._row <= range.getLastRow() && r.status === 'pripravljeno za lekturo');
      for (const row of rows) sendForReview_(ctx, row.id, deadline);
    }
  } finally { lock.releaseLock(); }
}

/** Shrani izvirno e-pošto (.eml) in vse podprte priponke v začasno mapo. */
function stageMessage_(ctx, message, deadline) {
  let mail = find_(ctx.log, L_KEYS, message.id);
  if (mail && mail.state !== STATE.STAGING) return mail;
  const subject = decodeHeader_(header_(message.payload, 'Subject')) || '(brez zadeve)';
  const sender = decodeHeader_(header_(message.payload, 'From'));
  const email = emailAddress_(sender);
  const body = bodyText_(message.payload).slice(0, 12000);
  if (!mail) {
    const folder = oneFolder_(ctx.stage, 'posta_' + message.id);
    mail = { id: message.id, thread: message.threadId, internetId: header_(message.payload, 'Message-ID'),
      sender: sender, email: email, org: organization_(ctx, email, subject + '\n' + body),
      subject: subject, hint: titleHint_(subject, body), received: new Date(Number(message.internalDate)),
      preview: body.replace(/\s+/g, ' ').slice(0, 1200), staging: folder.getId(), files: '[]',
      state: STATE.STAGING, target: '', reason: '', linked: '', error: '', body: body,
      kind: MATERIAL.NORMAL, reviewId: '' };
    put_(ctx.log, L_KEYS, mail);
    SpreadsheetApp.flush();
  }
  const folder = DriveApp.getFolderById(mail.staging);
  const manifest = [];
  const rawName = mail.id + '_izvirno-sporočilo.eml';
  let raw = firstFile_(folder, rawName);
  if (!raw) {
    requireTime_(deadline);
    const response = Gmail.Users.Messages.get('me', message.id, { format: 'raw' });
    if (!response.raw) throw new Error('Izvirnega sporočila ni bilo mogoče prebrati.');
    raw = folder.createFile(Utilities.newBlob(Utilities.base64DecodeWebSafe(response.raw), 'message/rfc822', rawName));
  }
  manifest.push({ id: raw.getId(), kind: 'email', name: rawName });
  const attachments = collectAttachments_(message.payload);
  for (let i = 0; i < attachments.length; i++) {
    requireTime_(deadline);
    const part = attachments[i];
    const name = mail.id + '_' + String(i + 1).padStart(3, '0') + '_' + cleanName_(part.filename);
    let file = firstFile_(folder, name);
    if (!file) {
      const body = part.body || {};
      const data = body.attachmentId ? Gmail.Users.Messages.Attachments.get('me', message.id, body.attachmentId).data : body.data;
      if (data === undefined || data === null) throw new Error('Manjkajo podatki priponke: ' + name);
      file = folder.createFile(Utilities.newBlob(Utilities.base64DecodeWebSafe(data), part.mimeType || 'application/octet-stream', name));
    }
    manifest.push({ id: file.getId(), kind: fileKind_(part.mimeType, part.filename), name: name });
  }
  mail.files = JSON.stringify(manifest);
  mail.state = STATE.READY;
  mail.error = '';
  return put_(ctx.log, L_KEYS, mail);
}

/** Predlog nikoli samodejno ne poveže dopolnitve z obstoječim člankom. */
function routeMessage_(ctx, mail, deadline) {
  if (mail.state === STATE.LINKING) return finishLink_(ctx, mail, deadline);
  if (mail.state === STATE.LINKED || mail.state === STATE.LEGACY) return;
  const returnMatch = reviewedReturn_(ctx, mail);
  if (returnMatch) {
    mail.reviewId = returnMatch.request.id;
    mail.kind = returnMatch.current ? MATERIAL.REVIEWED : MATERIAL.OLD_REVIEW;
    return beginLink_(ctx, mail, returnMatch.article.id, 'Samodejna povezava po ID-ju članka, krogu lekture, naslovu lektorja in imenu datoteke.', deadline);
  }
  const suggestions = matchArticles_(ctx, mail);
  const files = JSON.parse(mail.files);
  const imagesOnly = files.some(f => f.kind === 'image') && !files.some(f => f.kind === 'document');
  const followup = followupCue_(mail.subject + '\n' + mail.body);
  const subjectId = subjectArticleId_(mail.subject);
  if (suggestions.length || imagesOnly || followup || subjectId || mail.state === STATE.PENDING) {
    mail.state = STATE.PENDING;
    mail.reason = suggestions.length ? suggestions.map(s => s.reasons.join(', ')).join(' | ')
      : imagesOnly ? 'Prejete so samo slike; članek mora izbrati urednik.'
        : 'Besedilo nakazuje dopolnitev; članek mora izbrati urednik.';
    put_(ctx.log, L_KEYS, mail);
    const existing = find_(ctx.pending, P_KEYS, mail.id);
    put_(ctx.pending, P_KEYS, { received: mail.received, sender: mail.sender, org: mail.org,
      subject: mail.subject, preview: mail.preview,
      suggestions: suggestions.map(s => articleLabel_(s.article)).join('\n'), reason: mail.reason,
      stagingUrl: DriveApp.getFolderById(mail.staging).getUrl(), selection: existing ? existing.selection : '',
      confirm: existing ? existing.confirm : false, error: existing ? existing.error : '', id: mail.id,
      kind: existing ? existing.kind : MATERIAL.NORMAL });
    pendingControls_(ctx);
  } else {
    beginLink_(ctx, mail, 'NN-' + mail.id, 'Samodejno ustvarjen nov članek.', deadline);
  }
}

/** Po izbiri članka shrani cilj, nato premakne gradivo in zaključi povezavo. */
function confirmPending_(ctx, messageId, deadline) {
  const pending = find_(ctx.pending, P_KEYS, messageId);
  if (!pending || pending.confirm !== true) return;
  try {
    let mail = find_(ctx.log, L_KEYS, messageId);
    if (!mail) throw new Error('Sporočilo ni najdeno v dnevniku pošte.');
    if (mail.state === STATE.LINKED || mail.state === STATE.LEGACY) return cleanupLinked_(ctx, mail);
    if (mail.state === STATE.LINKING) return finishLink_(ctx, mail, deadline);
    if (mail.state !== STATE.PENDING) throw new Error('Gradivo še ni pripravljeno za povezavo.');
    if (!pending.selection) throw new Error('Najprej izberite članek v stolpcu »Izbrani članek«.');
    let target;
    if (pending.selection === NEW_ARTICLE) target = 'NN-' + mail.id;
    else {
      const match = String(pending.selection).match(/\[(NN-[a-zA-Z0-9_-]+)\]$/);
      target = match && match[1];
      if (!target || !find_(ctx.articles, A_KEYS, target)) throw new Error('Izbrani članek ne obstaja več. Ponovno ga izberite.');
    }
    mail.kind = pending.kind === MATERIAL.REVIEWED ? MATERIAL.REVIEWED : MATERIAL.NORMAL;
    const reviewMatch = String(mail.subject).match(/\[(NN-[a-zA-Z0-9_-]+)\]\[L(\d+)\]/i);
    if (mail.kind === MATERIAL.REVIEWED && reviewMatch && reviewMatch[1] === target) {
      const request = find_(ctx.reviews, R_KEYS, target + '-L' + Number(reviewMatch[2]));
      if (request) {
        mail.reviewId = request.id;
        const article = find_(ctx.articles, A_KEYS, target);
        if (!reviewIsCurrent_(article, request)) mail.kind = MATERIAL.OLD_REVIEW;
      }
    }
    beginLink_(ctx, mail, target, 'Povezavo je potrdil urednik.', deadline);
  } catch (error) {
    reportFailure_(ctx, messageId, error);
    const row = find_(ctx.pending, P_KEYS, messageId);
    if (row) put_(ctx.pending, P_KEYS, Object.assign(row, { confirm: false, error: String(error.message || error) }));
  }
}

function beginLink_(ctx, mail, target, reason, deadline) {
  mail.state = STATE.LINKING;
  mail.target = target;
  mail.reason = reason;
  mail.error = '';
  put_(ctx.log, L_KEYS, mail);
  SpreadsheetApp.flush(); // Cilj ostane shranjen tudi ob prekinitvi prenosa.
  finishLink_(ctx, mail, deadline);
}

function finishLink_(ctx, mail, deadline) {
  let article = find_(ctx.articles, A_KEYS, mail.target);
  if (!article && mail.target === 'NN-' + mail.id) article = createArticle_(ctx, mail);
  if (!article) throw new Error('Ciljni članek ni najden: ' + mail.target);
  if (mail.kind === MATERIAL.REVIEWED && mail.reviewId && article.materialId !== mail.id) {
    const request = find_(ctx.reviews, R_KEYS, mail.reviewId);
    if (request && !reviewIsCurrent_(article, request)) {
      mail.kind = MATERIAL.OLD_REVIEW;
      put_(ctx.log, L_KEYS, mail);
    }
  }
  const folder = DriveApp.getFolderById(article.folderId);
  if (folder.isTrashed()) throw new Error('Mapa izbranega članka je v košu. Najprej jo obnovite.');
  for (const item of JSON.parse(mail.files)) {
    requireTime_(deadline);
    const file = DriveApp.getFileById(item.id);
    if (file.isTrashed()) throw new Error('Datoteka je v košu: ' + item.name);
    // Premik ohrani ID; ponovljen zagon ne naredi nove kopije datoteke.
    let alreadyThere = false;
    const parents = file.getParents();
    while (parents.hasNext()) if (parents.next().getId() === folder.getId()) alreadyThere = true;
    if (!alreadyThere) file.moveTo(folder);
  }
  updateArticle_(ctx, mail.target, mail.id);
  if (mail.reviewId) {
    const request = find_(ctx.reviews, R_KEYS, mail.reviewId);
    if (request) {
      request.state = mail.kind === MATERIAL.OLD_REVIEW ? 'Vrnjena starejša različica' : 'Lekturirano';
      request.responseId = mail.id;
      request.returned = mail.received;
      request.error = '';
      put_(ctx.reviews, R_KEYS, request);
    }
  }
  mail.state = STATE.LINKED;
  mail.linked = new Date(Date.now());
  mail.error = '';
  put_(ctx.log, L_KEYS, mail);
  SpreadsheetApp.flush();
  cleanupLinked_(ctx, mail);
}

/** Prazno začasno mapo odstrani šele po uspešnem prenosu in vpisu evidence. */
function cleanupLinked_(ctx, mail) {
  const row = find_(ctx.pending, P_KEYS, mail.id);
  if (row) ctx.pending.deleteRow(row._row);
  if (mail.staging) {
    const folder = DriveApp.getFolderById(mail.staging);
    if (!folder.isTrashed() && !folder.getFiles().hasNext() && !folder.getFolders().hasNext()) folder.setTrashed(true);
  }
}

/** Nedokončana opravila nadaljuje tudi po izpadu ali sočasnih potrditvah. */
function recover_(ctx, deadline) {
  let failures = 0;
  for (let mail of records_(ctx.log, L_KEYS)) {
    requireTime_(deadline);
    try {
      if (mail.state === STATE.STAGING) {
        const message = Gmail.Users.Messages.get('me', mail.id, { format: 'full' });
        mail = stageMessage_(ctx, message, deadline);
        routeMessage_(ctx, mail, deadline);
      } else if (mail.state === STATE.READY) routeMessage_(ctx, mail, deadline);
      else if (mail.state === STATE.LINKING) finishLink_(ctx, mail, deadline);
    } catch (error) { failures++; reportFailure_(ctx, mail.id, error); }
  }
  for (const pending of records_(ctx.pending, P_KEYS)) {
    requireTime_(deadline);
    const mail = find_(ctx.log, L_KEYS, pending.id);
    if (mail && (mail.state === STATE.LINKED || mail.state === STATE.LEGACY)) cleanupLinked_(ctx, mail);
    else if (pending.confirm === true) confirmPending_(ctx, pending.id, deadline);
  }
  // Obnovi vrstico, če je bil zagon prekinjen med vpisom dnevnika in čakalnice.
  for (const mail of records_(ctx.log, L_KEYS)) {
    if (mail.state === STATE.PENDING && !find_(ctx.pending, P_KEYS, mail.id)) routeMessage_(ctx, mail, deadline);
  }
  return failures;
}

function createArticle_(ctx, mail) {
  const id = 'NN-' + mail.id;
  const existing = find_(ctx.articles, A_KEYS, id);
  if (existing) return existing;
  const title = mail.hint || mail.subject;
  const name = cleanName_(title).slice(0, 180);
  const marker = 'ID članka: ' + id;
  const candidates = ctx.root.getFoldersByName(name);
  let folder;
  let collision = false;
  while (candidates.hasNext()) {
    const item = candidates.next();
    if (item.getDescription() === marker) folder = item;
    collision = true;
  }
  if (!folder) {
    const targetName = collision ? name + ' [' + id + ']' : name;
    const retry = ctx.root.getFoldersByName(targetName);
    while (retry.hasNext()) {
      const item = retry.next();
      if (item.getDescription() === marker) folder = item;
    }
    if (!folder) folder = ctx.root.createFolder(targetName).setDescription(marker);
  }
  const article = { subject: mail.subject, sender: mail.sender, received: mail.received,
    folderUrl: folder.getUrl(), status: 'Prenos v teku', sourceId: mail.id, id: id, org: mail.org,
    emails: mail.email, title: title, images: 0, documents: 0, last: mail.received,
    update: 'Prvi prejem gradiva', folderId: folder.getId(), reviewer: CONFIG.DEFAULT_REVIEWER };
  put_(ctx.articles, A_KEYS, article);
  refreshChoices_(ctx);
  articleControls_(ctx, article.id);
  return article;
}

/** Števce preračuna iz dejanske mape, zato jih ponovitve ne povečujejo. */
function updateArticle_(ctx, articleId, incomingId) {
  const article = find_(ctx.articles, A_KEYS, articleId);
  const files = DriveApp.getFolderById(article.folderId).getFiles();
  let images = 0;
  let documents = 0;
  while (files.hasNext()) {
    const file = files.next();
    const kind = fileKind_(file.getMimeType(), file.getName());
    if (kind === 'image') images++;
    if (kind === 'document') documents++;
  }
  const linked = records_(ctx.log, L_KEYS).filter(m => m.target === articleId &&
    [STATE.LINKING, STATE.LINKED, STATE.LEGACY].includes(m.state))
    .sort((a, b) => Number(new Date(a.received)) - Number(new Date(b.received)));
  const latest = linked[linked.length - 1];
  const emails = new Set(String(article.emails || '').split(/[;,\s]+/).filter(Boolean));
  linked.forEach(m => { if (m.email) emails.add(m.email); });
  article.emails = Array.from(emails).join('; ');
  article.images = images;
  article.documents = documents;
  if (article.status === 'Prenos v teku') article.status = 'Prejeto';
  if (!article.reviewer) article.reviewer = CONFIG.DEFAULT_REVIEWER;
  if (!article.org) article.org = (linked.find(m => m.org) || {}).org || '';
  if (latest) {
    article.last = latest.received;
    if (latest.id === article.sourceId) article.update = 'Prvi prejem gradiva';
    else {
      const manifest = JSON.parse(latest.files || '[]');
      const newImages = manifest.filter(f => f.kind === 'image').length;
      const newDocs = manifest.filter(f => f.kind === 'document').length;
      const parts = [];
      if (newImages) parts.push('Prejete nove slike (' + newImages + ')');
      if (newDocs) parts.push(correctionCue_(latest.subject + '\n' + latest.body) ? 'Prejeta popravljena različica (' + newDocs + ')' : 'Prejeti novi dokumenti (' + newDocs + ')');
      article.update = parts.join('; ') || 'Prejeta dopolnitev';
    }
  }
  // Izbiro dokumenta določa potrjena povezava. Fotografije z novejšim datumom
  // ne smejo preprečiti poznejšega ročnega povezovanja starejšega popravka.
  const incoming = incomingId ? linked.find(m => m.id === incomingId)
    : !article.materialId ? linked.slice().reverse().find(m => JSON.parse(m.files || '[]').some(f => f.kind === 'document')) : null;
  if (incoming) {
    const latestDocs = JSON.parse(incoming.files || '[]').filter(f => f.kind === 'document');
    if (article.materialId !== incoming.id && latestDocs.length) {
      if (incoming.kind === MATERIAL.REVIEWED) {
        article.status = 'Lekturirano';
        article.reviewStatus = 'Prejeta in shranjena lekturirana različica.';
        article.update = 'Prejeta lekturirana različica';
      } else if (incoming.kind === MATERIAL.OLD_REVIEW) {
        article.reviewStatus = 'Prejeta lektura starejše različice. Novejši dokument še ni lekturiran.';
        article.update = 'Prejeta lektura prejšnje različice';
      } else if (article.currentReview && ['V lekturi', 'Lekturirano'].includes(article.status)) {
        article.status = 'Potrebna nova lektura';
        article.reviewStatus = 'Prejet je novejši dokument; obstoječa lektura velja za prejšnjo različico.';
      }
      if (incoming.kind !== MATERIAL.OLD_REVIEW) {
        const chosen = incoming.kind === MATERIAL.REVIEWED
          ? latestDocs.find(f => /_lekturirano\.(docx?|docm|pdf)$/i.test(f.name)) || latestDocs[0]
          : latestDocs.find(f => /\.(docx?|docm)$/i.test(f.name)) || latestDocs[0];
        article.reviewDoc = chosen.name + ' [' + chosen.id + ']';
        article.materialId = incoming.id;
      }
    }
  }
  put_(ctx.articles, A_KEYS, article);
  articleControls_(ctx, article.id);
}

/** Razložljivi predlogi: naslov, telo sporočila, pošiljatelj in organizacija. */
function matchArticles_(ctx, mail) {
  const corpus = normalize_(mail.subject + ' ' + mail.hint + ' ' + mail.body);
  const subject = normalize_(titleHint_(mail.subject, ''));
  const logs = records_(ctx.log, L_KEYS);
  const result = [];
  for (const article of records_(ctx.articles, A_KEYS)) {
    let score = 0;
    const reasons = [];
    if (subjectArticleId_(mail.subject) === article.id) { score += 150; reasons.push('ID članka v zadevi'); }
    const senders = String(article.emails || emailAddress_(article.sender)).toLowerCase().split(/[;,\s]+/);
    if (mail.email && senders.includes(mail.email)) { score += 30; reasons.push('isti e-naslov pošiljatelja'); }
    if (mail.org && article.org && normalize_(mail.org) === normalize_(article.org)) { score += 35; reasons.push('ista organizacija'); }
    const title = normalize_(article.title || article.subject);
    if (title.length > 4 && (subject === title || (' ' + corpus + ' ').includes(' ' + title + ' '))) {
      score += 70; reasons.push('ujemanje naslova članka');
    } else {
      const tokens = titleTokens_(title);
      const bodyTokens = new Set(titleTokens_(corpus));
      const common = tokens.filter(t => bodyTokens.has(t));
      if (common.length >= 2 && common.length / Math.max(tokens.length, 1) >= 0.4) {
        score += 25 + Math.round(25 * common.length / tokens.length);
        reasons.push('skupne besede naslova: ' + common.slice(0, 5).join(', '));
      }
    }
    if (logs.some(m => m.target === article.id && [STATE.LINKED, STATE.LEGACY].includes(m.state) &&
      ((mail.thread && m.thread === mail.thread) || (m.internetId && mail.body.includes(m.internetId))))) {
      score += 80; reasons.push('isti e-poštni pogovor');
    }
    if (score >= CONFIG.MIN_MATCH_SCORE) result.push({ article: article, score: score, reasons: reasons });
  }
  return result.sort((a, b) => b.score - a.score).slice(0, 3);
}

function organization_(ctx, email, text) {
  const contacts = records_(ctx.contacts, ['email', 'org']);
  const contact = contacts.find(c => String(c.email).trim().toLowerCase() === email);
  if (contact) return contact.org;
  const corpus = ' ' + normalize_(text) + ' ';
  const organizations = Array.from(new Set(records_(ctx.articles, A_KEYS).map(a => a.org).filter(Boolean)))
    .sort((a, b) => b.length - a.length);
  return organizations.find(org => normalize_(org).length >= 4 && corpus.includes(' ' + normalize_(org) + ' ')) || '';
}

function titleHint_(subject, body) {
  const explicit = String(body).match(/(?:naslov\s+članka|naslov\s+prispevka|article\s+title|titel)\s*:\s*([^\r\n]{3,180})/i);
  if (explicit) return explicit[1].trim();
  return String(subject).replace(/^(?:(?:re|fw|fwd|odg|posr)\s*:\s*)+/gi, '').trim();
}

function followupCue_(text) {
  return /^(?:re|fw|fwd|odg|posr)\s*:/i.test(text) ||
    /\b(?:photos?|pictures?|corrected|correction|revised|revision|updated|fotograf\w*|slike|fotke|poprav\w*|dopoln\w*|isprav\w*|dopun\w*|berichtig\w*|korrektur\w*|nachtrag\w*|fotos?)\b/i.test(normalize_(text));
}

function correctionCue_(text) {
  return /\b(?:corrected|correction|revised|revision|updated|poprav\w*|isprav\w*|berichtig\w*|korrektur\w*)\b/i.test(normalize_(text));
}

function normalize_(text) {
  return String(text || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, ' ').trim().replace(/\s+/g, ' ');
}

function titleTokens_(text) {
  const stop = new Set('the and for our with this that here are article photos photo pictures corrected new clanek prispevek fotografije slike popravek dopolnitev za in na pri od je so ter se nase nas der die das und fur mit'.split(' '));
  return Array.from(new Set(normalize_(text).split(' ').filter(t => t.length >= 3 && !stop.has(t))));
}

function emailAddress_(value) {
  const match = String(value).match(/[a-z0-9.!#$%&'*+\/=?^_`{|}~-]+@[a-z0-9.-]+\.[a-z]{2,}/i);
  return match ? match[0].toLowerCase() : '';
}

function bodyText_(payload) {
  let plain = '';
  let html = '';
  function visit(part) {
    if (!part || part.filename) return;
    if (part.body && part.body.data && /^text\/(plain|html)$/i.test(part.mimeType || '')) {
      const contentType = header_(part, 'Content-Type');
      const charset = (contentType.match(/charset=["']?([^;"'\s]+)/i) || [])[1] || 'UTF-8';
      let text;
      try { text = Utilities.newBlob(Utilities.base64DecodeWebSafe(part.body.data)).getDataAsString(charset); }
      catch (error) { text = Utilities.newBlob(Utilities.base64DecodeWebSafe(part.body.data)).getDataAsString('UTF-8'); }
      if (part.mimeType.toLowerCase() === 'text/plain') plain += '\n' + text;
      else html += '\n' + text;
    }
    (part.parts || []).forEach(visit);
  }
  visit(payload);
  return (plain || html.replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, ' ').replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, ' ')
    .replace(/<[^>]*>/g, ' ').replace(/&nbsp;/gi, ' ').replace(/&amp;/gi, '&').replace(/&lt;/gi, '<').replace(/&gt;/gi, '>')).trim();
}

function fileKind_(mime, name) {
  if (/^image\//i.test(mime || '') || /\.(jpe?g|png|gif|webp|heic|heif|tiff?|bmp|avif|svg)$/i.test(name || '')) return 'image';
  if (/application\/(pdf|msword|vnd\..*word)/i.test(mime || '') || /\.(pdf|doc|docx|docm)$/i.test(name || '')) return 'document';
  return 'other';
}

/** Vsak članek ima trajen ID; besedilo pred njim se lahko med odgovori spremeni. */
function subjectArticleId_(subject) {
  const match = String(subject).match(/\[(NN-[a-zA-Z0-9_-]+)\]/i);
  return match ? match[1] : '';
}

function documentId_(label) {
  const match = String(label || '').match(/\[([a-zA-Z0-9_-]+)\]$/);
  return match ? match[1] : '';
}

function documentChoices_(article) {
  const files = DriveApp.getFolderById(article.folderId).getFiles();
  const result = [];
  while (files.hasNext()) {
    const file = files.next();
    if (fileKind_(file.getMimeType(), file.getName()) === 'document') result.push(file.getName() + ' [' + file.getId() + ']');
  }
  return result.sort();
}

/** Spustni seznam stanj in izbira konkretne različice dokumenta. */
function articleControls_(ctx, onlyArticleId) {
  const count = ctx.articles.getLastRow() - 1;
  if (count < 1) return;
  ctx.articles.getRange(2, 5, count, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(ARTICLE_STATUSES, true).setAllowInvalid(true).build());
  for (const article of records_(ctx.articles, A_KEYS)) {
    if (onlyArticleId && article.id !== onlyArticleId) continue;
    if (!article.folderId) continue;
    const choices = documentChoices_(article);
    if (choices.length) {
      ctx.articles.getRange(article._row, 17).setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInList(choices, true).setAllowInvalid(false)
        .setHelpText('Pred pošiljanjem preverite, katero različico dokumenta bo prejel lektor.').build());
      if (!article.reviewDoc) {
        article.reviewDoc = choices[choices.length - 1];
        put_(ctx.articles, A_KEYS, article);
      }
    }
  }
}

function processReviewQueue_(ctx, deadline) {
  for (const article of records_(ctx.articles, A_KEYS)) {
    requireTime_(deadline);
    if (article.status === 'pripravljeno za lekturo') sendForReview_(ctx, article.id, deadline);
  }
}

/** Izbira stanja je urednikovo navodilo za pošiljanje dokumenta lektorju. */
function sendForReview_(ctx, articleId, deadline) {
  let article = find_(ctx.articles, A_KEYS, articleId);
  let request;
  let attempted = false;
  try {
    requireTime_(deadline);
    const recipient = String(article.reviewer || CONFIG.DEFAULT_REVIEWER).trim().toLowerCase();
    if (!recipient || emailAddress_(recipient) !== recipient) throw new Error('V stolpec »Lektor« vnesite en veljaven e-naslov.');
    const docId = documentId_(article.reviewDoc);
    if (!docId) throw new Error('Najprej izberite »Dokument za lekturo«.');
    const file = DriveApp.getFileById(docId);
    const parents = file.getParents();
    let inArticle = false;
    while (parents.hasNext()) if (parents.next().getId() === article.folderId) inArticle = true;
    if (!inArticle || file.isTrashed() || fileKind_(file.getMimeType(), file.getName()) !== 'document') {
      throw new Error('Izbrani dokument mora biti Word ali PDF v mapi tega članka.');
    }
    const history = records_(ctx.reviews, R_KEYS).filter(r => r.article === article.id);
    request = history.slice().reverse().find(r => r.docId === docId && r.recipient === recipient &&
      ['Pripravljeno', 'Pošiljanje', 'Poslano', 'Preverite pošiljanje', 'Napaka pred pošiljanjem', 'Ni bilo poslano'].includes(r.state));
    if (!request) {
      const round = Math.max(0, ...history.map(r => Number(r.round) || 0)) + 1;
      request = { id: article.id + '-L' + round, article: article.id, round: round, recipient: recipient,
        subject: '[' + article.id + '][L' + round + '] ' + article.title,
        docId: docId, docName: file.getName(), created: new Date(Date.now()), sent: '',
        state: 'Pripravljeno', responseId: '', returned: '', error: '' };
      put_(ctx.reviews, R_KEYS, request);
    }
    article.reviewer = recipient;
    article.reviewRound = request.round;
    article.reviewSubject = request.subject;
    article.currentReview = request.id;
    if (request.state === 'Poslano') {
      article.status = 'V lekturi'; article.reviewSent = request.sent;
      article.reviewStatus = 'Ta dokument je bil temu lektorju že poslan. Ponovno pošiljanje ni bilo izvedeno.';
      put_(ctx.articles, A_KEYS, article); return;
    }
    if (['Pošiljanje', 'Preverite pošiljanje'].includes(request.state)) {
      article.status = 'Preverite pošiljanje';
      article.reviewStatus = 'Izid prejšnjega pošiljanja ni potrjen. Preverite pošto; v zavihku »Lekture« označite »Poslano« ali »Ni bilo poslano« in ponovno izberite pripravo za lekturo.';
      put_(ctx.articles, A_KEYS, article); return;
    }
    const extension = (file.getName().match(/\.(docx?|docm|pdf)$/i) || [])[0];
    if (!extension) throw new Error('Dokument za lekturo mora imeti končnico .doc, .docx, .docm ali .pdf.');
    const outgoingName = article.id + '_L' + request.round + extension;
    const returnedName = article.id + '_L' + request.round + '_lekturirano' + extension;
    const blob = file.getBlob().setName(outgoingName);
    if (MailApp.getRemainingDailyQuota() < 1) throw new Error('Dnevna omejitev pošiljanja je dosežena. Poskusite pozneje.');
    const body = 'Pozdravljeni,\n\nprosim za lekturo priloženega članka.\n\n' +
      'Naslov članka: ' + article.title + '\nID članka: ' + article.id + '\nKrog lekture: ' + request.round +
      '\n\nOdgovorite na to sporočilo in ohranite isto zadevo, vključno z ID-jem in oznako kroga. ' +
      'Običajna predpona »Re:« je dovoljena.\nPopravljeni dokument shranite z imenom:\n' + returnedName +
      '\n\nPripnite popravljeni dokument odgovoru. Samo besedilo odgovora brez priponke se ne obravnava kot vrnjena lektura.' +
      '\n\nHvala in lep pozdrav,\nUredništvo Naših novin';
    article.reviewStatus = 'Pošiljanje dokumenta lektorju.';
    put_(ctx.articles, A_KEYS, article);
    request.state = 'Pošiljanje'; request.error = '';
    put_(ctx.reviews, R_KEYS, request);
    SpreadsheetApp.flush();
    attempted = true;
    MailApp.sendEmail({ to: recipient, subject: request.subject, body: body, attachments: [blob],
      name: 'Naše novine – uredništvo', replyTo: ctx.props.getProperty('NN_MAILBOX') });
    request.state = 'Poslano'; request.sent = new Date(Date.now());
    put_(ctx.reviews, R_KEYS, request);
    article = find_(ctx.articles, A_KEYS, articleId);
    article.status = 'V lekturi'; article.reviewSent = request.sent;
    article.reviewStatus = 'Dokument je poslan lektorju: ' + recipient;
    put_(ctx.articles, A_KEYS, article);
  } catch (error) {
    const text = String(error.message || error);
    if (request) {
      request.state = attempted ? 'Preverite pošiljanje' : 'Napaka pred pošiljanjem';
      request.error = text;
      put_(ctx.reviews, R_KEYS, request);
    }
    article = find_(ctx.articles, A_KEYS, articleId);
    article.status = attempted ? 'Preverite pošiljanje' : 'Napaka pri lekturi';
    article.reviewStatus = text + (attempted ? ' Ponovno pošiljanje je zadržano do ročnega preverjanja.' : ' Odpravite napako in ponovno izberite pripravo za lekturo.');
    put_(ctx.articles, A_KEYS, article);
    console.error('Lektura ' + articleId + ': ' + text);
  }
}

/** Samodejno sprejmemo samo dogovorjen odgovor lektorja z ustrezno datoteko. */
function reviewedReturn_(ctx, mail) {
  const match = String(mail.subject).match(/\[(NN-[a-zA-Z0-9_-]+)\]\[L(\d+)\]/i);
  if (!match) return null;
  const article = find_(ctx.articles, A_KEYS, match[1]);
  const request = find_(ctx.reviews, R_KEYS, match[1] + '-L' + Number(match[2]));
  if (!article || !request || request.recipient !== mail.email) return null;
  if (!['Poslano', 'Pošiljanje', 'Preverite pošiljanje'].includes(request.state)) return null;
  const documents = JSON.parse(mail.files).filter(f => f.kind === 'document');
  if (!documents.some(f => /_lekturirano\.(docx?|docm|pdf)$/i.test(f.name))) return null;
  return { article: article, request: request, current: reviewIsCurrent_(article, request) };
}

function reviewIsCurrent_(article, request) {
  return article && article.currentReview === request.id && documentId_(article.reviewDoc) === request.docId;
}

/** Izjema za primer, ko isti račun prejema članke in tudi opravlja lekturo. */
function selfReviewReturn_(ctx, message) {
  const mailbox = ctx.props.getProperty('NN_MAILBOX');
  const sender = emailAddress_(decodeHeader_(header_(message.payload, 'From')));
  if (sender !== mailbox || emailAddress_(header_(message.payload, 'To')) !== mailbox) return false;
  const subject = decodeHeader_(header_(message.payload, 'Subject'));
  const match = subject.match(/\[(NN-[a-zA-Z0-9_-]+)\]\[L(\d+)\]/i);
  if (!match) return false;
  const request = find_(ctx.reviews, R_KEYS, match[1] + '-L' + Number(match[2]));
  return Boolean(request && request.recipient === mailbox && collectAttachments_(message.payload)
    .some(f => /_lekturirano\.(docx?|docm|pdf)$/i.test(f.filename)));
}

function context_() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('NN_BOOK_ID') || !props.getProperty('NN_STAGING_ID')) throw new Error('Najprej zaženite setup().');
  const ctx = sheets_(SpreadsheetApp.openById(props.getProperty('NN_BOOK_ID')), false);
  return Object.assign(ctx, { props: props, root: DriveApp.getFolderById(props.getProperty('NN_FOLDER_ID')),
    stage: DriveApp.getFolderById(props.getProperty('NN_STAGING_ID')) });
}

function sheets_(book, install) {
  function ensure(name, headers, legacy) {
    let sheet = book.getSheetByName(name);
    if (!sheet && install) sheet = book.insertSheet(name);
    if (!sheet) throw new Error('Manjka zavihek »' + name + '«. Ponovno zaženite setup().');
    if (install) {
      if (sheet.getMaxColumns() < headers.length) sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
      const existing = sheet.getLastRow() ? sheet.getRange(1, 1, 1, Math.min(headers.length, 6)).getDisplayValues()[0] : [];
      const oldHeaders = ['Email subject', 'Sender', 'Received date', 'Drive folder link', 'Status', 'Gmail message ID'];
      const validLegacy = legacy && existing.every((v, i) => v === oldHeaders[i]);
      if (existing.length && !validLegacy && existing.some((v, i) => v !== headers[i])) throw new Error('Nepričakovane glave v zavihku »' + name + '«. Podatki niso bili spremenjeni.');
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    const actual = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
    if (actual.some((v, i) => v !== headers[i])) throw new Error('Spremenjene glave v zavihku »' + name + '«. Obnovite vrstni red stolpcev.');
    return sheet;
  }
  const ctx = { book: book, articles: ensure(CONFIG.SHEET_NAME, A_HEADERS, true),
    pending: ensure(CONFIG.PENDING_SHEET_NAME, P_HEADERS), log: ensure(CONFIG.LOG_SHEET_NAME, L_HEADERS),
    contacts: ensure(CONFIG.CONTACTS_SHEET_NAME, ['E-naslov', 'Organizacija']),
    choices: ensure(CONFIG.CHOICES_SHEET_NAME, ['Članek za povezavo']),
    reviews: ensure(CONFIG.REVIEW_SHEET_NAME, R_HEADERS) };
  if (install) {
    ctx.articles.hideColumns(6); ctx.articles.hideColumns(15);
    ctx.articles.hideColumns(22, 2);
    ctx.pending.hideColumns(12); ctx.log.hideColumns(11, 2); ctx.log.hideColumns(18);
    ctx.choices.hideSheet();
    ctx.articles.setColumnWidth(10, 300); ctx.articles.setColumnWidth(14, 320);
    ctx.articles.setColumnWidth(7, 230); ctx.articles.setColumnWidth(16, 220);
    ctx.articles.setColumnWidth(17, 380); ctx.articles.setColumnWidth(20, 400); ctx.articles.setColumnWidth(21, 380);
    ctx.pending.setColumnWidth(5, 350); ctx.pending.setColumnWidth(6, 350); ctx.pending.setColumnWidth(9, 350);
    ctx.pending.getRange(1, 10).setNote('Izberite članek v stolpcu I in nato označite to polje. Povezana vrstica bo odstranjena iz čakalnice.');
  }
  return ctx;
}

/** Ohranitev obstoječih šestih stolpcev in podatkov prve različice skripta. */
function migrateArticles_(ctx) {
  for (const article of records_(ctx.articles, A_KEYS)) {
    if (!article.sourceId) continue;
    article.id = article.id || 'NN-' + article.sourceId;
    article.title = article.title || article.subject;
    article.emails = article.emails || emailAddress_(article.sender);
    article.folderId = article.folderId || folderId_(article.folderUrl);
    if (article.status === 'Received') article.status = 'Prejeto';
    article.last = article.last || article.received;
    article.update = article.update || 'Prvi prejem gradiva';
    article.reviewer = article.reviewer || CONFIG.DEFAULT_REVIEWER;
    put_(ctx.articles, A_KEYS, article);
    if (!find_(ctx.log, L_KEYS, article.sourceId)) {
      put_(ctx.log, L_KEYS, { id: article.sourceId, sender: article.sender, email: emailAddress_(article.sender),
        org: article.org, subject: article.subject, received: article.received, files: '[]',
        state: STATE.LEGACY, target: article.id, reason: 'Preneseno iz prejšnje evidence.' });
    }
    updateArticle_(ctx, article.id);
  }
}

function refreshChoices_(ctx) {
  const labels = records_(ctx.articles, A_KEYS).filter(a => a.id).map(articleLabel_).sort();
  const values = [[NEW_ARTICLE]].concat(labels.map(label => [safeCell_(label)]));
  const oldCount = ctx.choices.getLastRow() - 1;
  ensureRows_(ctx.choices, values.length + 1);
  if (oldCount > 0) ctx.choices.getRange(2, 1, oldCount, 1).clearContent();
  ctx.choices.getRange(2, 1, values.length, 1).setNumberFormat('@').setValues(values);
  pendingControls_(ctx);
}

function pendingControls_(ctx) {
  const count = ctx.pending.getLastRow() - 1;
  if (count < 1) return;
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ctx.choices.getRange(2, 1, Math.max(1, ctx.choices.getLastRow() - 1), 1), true)
    .setAllowInvalid(false).setHelpText('Izberite obstoječi članek ali možnost za nov članek.').build();
  ctx.pending.getRange(2, 9, count, 1).setDataValidation(rule);
  // V nasprotju z insertCheckboxes to ne ponastavi že označenih potrditvenih polj.
  const check = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  ctx.pending.getRange(2, 10, count, 1).setDataValidation(check);
  ctx.pending.getRange(2, 13, count, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList([MATERIAL.NORMAL, MATERIAL.REVIEWED], true).setAllowInvalid(false).build());
}

function articleLabel_(article) { return (article.title || article.subject) + ' [' + article.id + ']'; }

function records_(sheet, keys) {
  const count = sheet.getLastRow() - 1;
  if (count < 1) return [];
  return sheet.getRange(2, 1, count, keys.length).getValues().map((values, i) => {
    const record = { _row: i + 2 };
    keys.forEach((key, j) => record[key] = values[j] === undefined ? '' : values[j]);
    return record;
  }).filter(r => r[keys[0]] !== '');
}

function find_(sheet, keys, id) {
  return records_(sheet, keys).find(r => String(r.id) === String(id));
}

function put_(sheet, keys, record) {
  const existing = find_(sheet, keys, record.id) || (keys === A_KEYS && record.sourceId
    ? records_(sheet, keys).find(r => r.sourceId === record.sourceId) : null);
  const row = existing ? existing._row : sheet.getLastRow() + 1;
  ensureRows_(sheet, row);
  const values = keys.map(key => typeof record[key] === 'string' ? safeCell_(record[key]) : record[key] ?? '');
  // Besedilne ID-je vedno ohranimo, tudi če so sestavljeni samo iz številk.
  keys.forEach((key, i) => {
    if (['id', 'sourceId', 'thread', 'target', 'internetId', 'folderId', 'staging'].includes(key)) sheet.getRange(row, i + 1).setNumberFormat('@');
    if (['received', 'linked', 'last', 'reviewSent', 'created', 'sent', 'returned'].includes(key)) sheet.getRange(row, i + 1).setNumberFormat('dd. mm. yyyy hh:mm');
  });
  sheet.getRange(row, 1, 1, keys.length).setValues([values]);
  return Object.assign({}, record, { _row: row });
}

function ensureRows_(sheet, row) {
  if (row > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), Math.max(100, row - sheet.getMaxRows()));
}

function folderId_(url) {
  const match = String(url || '').match(/\/folders\/([^/?#]+)/);
  if (!match) throw new Error('Povezava do mape članka ni veljavna: ' + url);
  return match[1];
}

function oneFolder_(parent, name) {
  const matches = parent.getFoldersByName(name);
  const folder = matches.hasNext() ? matches.next() : parent.createFolder(name);
  if (matches.hasNext()) throw new Error('Več map ima ime »' + name + '«. Odpravite podvojitev.');
  return folder;
}

function firstFile_(folder, name) {
  const matches = folder.getFilesByName(name);
  return matches.hasNext() ? matches.next() : null;
}

function reportFailure_(ctx, id, error) {
  const text = String(error.message || error);
  console.error('Sporočilo ' + id + ': ' + text);
  const mail = find_(ctx.log, L_KEYS, id);
  if (mail) { mail.error = text; put_(ctx.log, L_KEYS, mail); }
  const pending = find_(ctx.pending, P_KEYS, id);
  if (pending) { pending.error = text; put_(ctx.pending, P_KEYS, pending); }
}

function requireTime_(deadline) {
  if (Date.now() >= deadline) throw new Error('Čas za izvajanje je potekel. Naslednji zagon bo nadaljeval nedokončano delo.');
}

function stopAutomation() {
  ScriptApp.getProjectTriggers().filter(t => ['processSubmissions', 'onPendingEdit'].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));
  console.log('Samodejna obdelava in potrjevanje sta ustavljena. Za nadaljevanje zaženite setup().');
}
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
      const filename = part.filename || 'priponka.' + (mimeExtensions[type] || type.split('/')[1]);
      const disposition = header_(part, 'Content-Disposition');
      const inline = /^inline\b/i.test(disposition) ||
        (!/^attachment\b/i.test(disposition) && Boolean(header_(part, 'Content-ID')));
      const image = type.startsWith('image/') || /\.(jpe?g|png|gif|webp|heic|heif|tiff?|bmp|avif|svg)$/i.test(filename);
      const document = Boolean(mimeExtensions[type]) || /\.(pdf|doc|docx|docm)$/i.test(filename);
      if ((image || document) && (!inline || CONFIG.INCLUDE_INLINE_IMAGES)) {
        result.push(Object.assign({}, part, { filename: filename }));
      }
      return; // ZIP datotek in priloženih sporočil .eml ne razširjamo.
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

/** Razkodiranje naslovov in imen po RFC 2047, tudi šumnikov. */
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
  return String(value).replace(/[\\/\u0000-\u001f\u007f]/g, ' ').trim() || 'brez imena';
}

function safeCell_(value) {
  // Podatke iz e-pošte zapišemo kot besedilo, nikoli kot formule.
  const text = String(value);
  return /^[\s]*[=+@-]/.test(text) ? "'" + text : text;
}


