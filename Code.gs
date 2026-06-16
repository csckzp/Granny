// =============================================================================
//  Code.gs — Granny Genealogy App · Google Apps Script Backend
// =============================================================================
//
//  Communication with Vue frontend (two modes):
//
//  1. google.script.run  (preferred — no CORS, works because HTML is served
//     by this same deployment)
//
//     google.script.run
//       .withSuccessHandler(cb)
//       .withFailureHandler(errCb)
//       .execute({ action: 'getPedigreePayload', data: { person_id: 'PER0001' } });
//
//  2. REST via doGet / doPost  (useful for testing or external integrations)
//
//     GET  ?action=getPedigreePayload&person_id=PER0001
//     POST { "action": "savePerson", "data": { "first_name": "Jane", ... } }
//
//  All responses share the shape: { ok: true, ... } | { ok: false, error: "..." }
// =============================================================================

// ─── Sheet name constants ─────────────────────────────────────────────────────

const SHEETS = {
  USERS:              'Users',
  PEOPLE:             'People',
  FAMILIES:           'Families',
  FAMILY_CHILDREN:    'Family_Children',
  EVENTS:             'Events',
  EVENT_PARTICIPANTS: 'Event_Participants',
};

// ─── HTTP entry points ────────────────────────────────────────────────────────

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  // No action param → serve the Vue HTML shell
  if (!action) {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Granny')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Action param present → treat as a read-only API call (GET = reads only)
  const READ_ACTIONS = new Set([
    'getPerson', 'listPeople', 'getPedigreePayload', 'listEvents',
  ]);
  if (!READ_ACTIONS.has(action)) {
    return _jsonOut({ ok: false, error: 'Use POST for write operations.' });
  }
  return _jsonOut(route({ action, data: e.parameter }));
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    return _jsonOut(route(payload));
  } catch (err) {
    return _jsonOut({ ok: false, error: err.message });
  }
}

function _jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── Universal router ─────────────────────────────────────────────────────────
//
//  Called directly by google.script.run.execute({action, data}) from Vue,
//  or internally by doGet/doPost.

function execute(payload) {
  return _sanitizeForClient(route(payload));
}

function route(payload) {
  const { action, data } = payload || {};
  try {
    switch (action) {
      // People
      case 'listPeople':              return listPeople();
      case 'getPerson':               return getPerson(data.person_id);
      case 'getPedigreePayload':      return getPedigreePayload(data.person_id);
      case 'savePerson':              return savePerson(data);
      case 'deletePerson':            return deletePerson(data.person_id);

      // Families
      case 'saveFamily':              return saveFamily(data);
      case 'deleteFamily':            return deleteFamily(data.family_id);

      // Family_Children junction
      case 'addChild':                return addFamilyChild(data);
      case 'removeChild':             return removeFamilyChild(data);

      // Events
      case 'listEvents':              return listEvents();
      case 'saveEvent':               return saveEvent(data);
      case 'deleteEvent':             return deleteEvent(data.event_id);
      case 'addEventParticipant':     return addEventParticipant(data);
      case 'removeEventParticipant':  return removeEventParticipant(data);

      default:
        return { ok: false, error: `Unknown action: "${action}"` };
    }
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// ─── Sheet utilities ──────────────────────────────────────────────────────────

function _ss() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function _sheet(name) {
  return _ss().getSheetByName(name);
}

/**
 * Converts server values into shapes that google.script.run can reliably
 * marshal back to the browser. In particular, Sheet date cells become JS Date
 * objects when read via getValues(), and those are not safe to return directly.
 */
function _sanitizeForClient(value) {
  if (value === undefined || value === null) return null;
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return value.toISOString();
  }
  if (Array.isArray(value)) {
    return value.map(_sanitizeForClient);
  }
  if (typeof value === 'object') {
    return Object.fromEntries(
      Object.entries(value).map(([key, item]) => [key, _sanitizeForClient(item)])
    );
  }
  return value;
}

/**
 * Reads a whole sheet into an array of plain objects keyed by header row.
 * Skips rows where column A is blank (deleted / empty trailing rows).
 */
function _toObjects(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1)
    .filter(row => row[0] !== '' && row[0] !== null && row[0] !== undefined)
    .map(row => Object.fromEntries(headers.map((h, i) => [h, row[i]])));
}

/**
 * Returns the 1-based row number for a record whose first column === pkValue.
 * Returns -1 when not found.
 */
function _findRow(sheet, pkValue) {
  const col = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  for (let i = 1; i < col.length; i++) {
    if (String(col[i][0]) === String(pkValue)) return i + 1;
  }
  return -1;
}

/**
 * Generates the next sequential ID for a sheet.
 * Scans column A for highest existing suffixed number in the prefix series.
 */
function _nextId(sheet, prefix) {
  const col = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), 1).getValues();
  let max = 0;
  for (let i = 1; i < col.length; i++) {
    const cell = String(col[i][0]);
    if (cell.startsWith(prefix)) {
      const n = parseInt(cell.slice(prefix.length), 10);
      if (!isNaN(n) && n > max) max = n;
    }
  }
  return prefix + String(max + 1).padStart(4, '0');
}

/**
 * Writes a plain-object row according to the sheet's header order.
 * Missing keys become empty strings.
 */
function _objectToRow(headers, obj) {
  return headers.map(h => (h in obj ? obj[h] : ''));
}

/**
 * Acquires a script-level lock (serialises all writes) and returns it.
 * Throws if the lock cannot be obtained within 15 seconds.
 */
function _acquireLock() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (_) {
    throw new Error('Server is busy — another write is in progress. Please retry.');
  }
  return lock;
}

// ─── Auth helpers ─────────────────────────────────────────────────────────────

function _currentUser() {
  const email = Session.getActiveUser().getEmail();
  if (!email) return null;
  const users = _toObjects(_sheet(SHEETS.USERS));
  return users.find(u => u.email === email) || null;
}

function _requireRole(...allowed) {
  const user = _currentUser();
  if (!user || !allowed.includes(user.role)) {
    throw new Error('Permission denied.');
  }
  return user;
}

// ─── Pedigree Payload ─────────────────────────────────────────────────────────

/**
 * The primary read function for the Vue pedigree view.
 *
 * Returns a single deeply-nested object:
 * {
 *   ok: true,
 *   person:  { ...People row },
 *   parents: [
 *     { person: {...}, relationship_type: "biological", family_id: "FAM0001" }
 *   ],
 *   spouses: [
 *     {
 *       family:   { ...Families row },
 *       spouse:   { ...People row } | null,
 *       children: [ { person: {...}, relationship_type: "biological" } ]
 *     }
 *   ],
 *   events: [
 *     { ...Events row, role: "subject", co_participants: [ { person:{...}, role:"spouse" } ] }
 *   ]
 * }
 *
 * All six sheets are loaded once into memory so there are no repeated
 * sheet reads — each resolution step is an in-memory array scan or
 * object-index lookup.
 */
function getPedigreePayload(personId) {
  if (!personId) return { ok: false, error: 'person_id is required.' };

  const ss = _ss();
  const people      = _toObjects(ss.getSheetByName(SHEETS.PEOPLE));
  const families    = _toObjects(ss.getSheetByName(SHEETS.FAMILIES));
  const famChildren = _toObjects(ss.getSheetByName(SHEETS.FAMILY_CHILDREN));
  const events      = _toObjects(ss.getSheetByName(SHEETS.EVENTS));
  const evtParts    = _toObjects(ss.getSheetByName(SHEETS.EVENT_PARTICIPANTS));

  // Build a people index once so every subsequent lookup is O(1)
  const byId = Object.fromEntries(people.map(p => [p.person_id, p]));

  const person = byId[personId];
  if (!person) return { ok: false, error: `No person with id "${personId}".` };

  // ── Parents ──────────────────────────────────────────────
  // Find every Family_Children row where this person is the child,
  // then resolve both spouses of that family as "parent" entries.
  const parentLinks = famChildren.filter(fc => String(fc.child_id) === String(personId));
  const parents = parentLinks.flatMap(link => {
    const fam = families.find(f => String(f.family_id) === String(link.family_id));
    if (!fam) return [];
    const entry = { relationship_type: link.relationship_type, family_id: fam.family_id };
    const result = [];
    if (fam.spouse1_id && byId[fam.spouse1_id]) {
      result.push({ ...entry, person: byId[fam.spouse1_id] });
    }
    if (fam.spouse2_id && byId[fam.spouse2_id]) {
      result.push({ ...entry, person: byId[fam.spouse2_id] });
    }
    return result;
  });

  // ── Spouses / children ───────────────────────────────────
  // Find every Families row where this person is a spouse,
  // then resolve the partner and that family's children.
  const spouseFamilies = families.filter(
    f => String(f.spouse1_id) === String(personId) ||
         String(f.spouse2_id) === String(personId)
  );
  const spouses = spouseFamilies.map(fam => {
    const partnerId = String(fam.spouse1_id) === String(personId)
      ? fam.spouse2_id
      : fam.spouse1_id;

    const children = famChildren
      .filter(fc => String(fc.family_id) === String(fam.family_id))
      .map(fc => ({ person: byId[fc.child_id] || null, relationship_type: fc.relationship_type }))
      .filter(c => c.person !== null);

    return {
      family:   fam,
      spouse:   partnerId ? (byId[partnerId] || null) : null,
      children,
    };
  });

  // ── Events ───────────────────────────────────────────────
  const participations = evtParts.filter(ep => String(ep.person_id) === String(personId));
  const resolvedEvents = participations.map(ep => {
    const event = events.find(e => String(e.event_id) === String(ep.event_id));
    if (!event) return null;
    const coParticipants = evtParts
      .filter(x => String(x.event_id) === String(ep.event_id) &&
                   String(x.person_id) !== String(personId))
      .map(x => ({ person: byId[x.person_id] || null, role: x.role }))
      .filter(x => x.person !== null);
    return { ...event, role: ep.role, co_participants: coParticipants };
  }).filter(Boolean);

  return { ok: true, person, parents, spouses, events: resolvedEvents };
}

// ─── People CRUD ──────────────────────────────────────────────────────────────

function listPeople() {
  const data = _toObjects(_sheet(SHEETS.PEOPLE));
  return { ok: true, data };
}

function getPerson(personId) {
  if (!personId) return { ok: false, error: 'person_id required.' };
  const person = _toObjects(_sheet(SHEETS.PEOPLE)).find(
    p => String(p.person_id) === String(personId)
  );
  return person ? { ok: true, data: person } : { ok: false, error: 'Person not found.' };
}

function savePerson(data) {
  _requireRole('admin', 'editor');
  const lock = _acquireLock();
  try {
    const sheet = _sheet(SHEETS.PEOPLE);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const now = new Date().toISOString();

    if (data.person_id) {
      // ── UPDATE ──
      const rowNum = _findRow(sheet, data.person_id);
      if (rowNum === -1) return { ok: false, error: 'Person not found.' };
      data.updated_at = now;
      sheet.getRange(rowNum, 1, 1, headers.length)
           .setValues([_objectToRow(headers, data)]);
      return { ok: true, person_id: data.person_id };
    } else {
      // ── INSERT ──
      const user = _currentUser();
      data.person_id  = _nextId(sheet, 'PER');
      data.created_by = user ? user.user_id : '';
      data.created_at = now;
      data.updated_at = now;
      sheet.appendRow(_objectToRow(headers, data));
      return { ok: true, person_id: data.person_id };
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * Deletes a person and cascades:
 *  - Removes their rows from Family_Children (as child)
 *  - Removes their rows from Event_Participants
 *  - Nulls out spouse1_id / spouse2_id in any Families rows (does not delete
 *    the family, since other members remain)
 */
function deletePerson(personId) {
  _requireRole('admin');
  if (!personId) return { ok: false, error: 'person_id required.' };
  const lock = _acquireLock();
  try {
    const sheet = _sheet(SHEETS.PEOPLE);
    const rowNum = _findRow(sheet, personId);
    if (rowNum === -1) return { ok: false, error: 'Person not found.' };
    sheet.deleteRow(rowNum);

    // Cascade: remove from junction tables
    _deleteJunctionRows(SHEETS.FAMILY_CHILDREN,    'child_id',  personId);
    _deleteJunctionRows(SHEETS.EVENT_PARTICIPANTS,  'person_id', personId);

    // Cascade: null out spouse references in Families (don't delete the family)
    _nullifyColumn(SHEETS.FAMILIES, 'spouse1_id', personId);
    _nullifyColumn(SHEETS.FAMILIES, 'spouse2_id', personId);

    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// ─── Families CRUD ────────────────────────────────────────────────────────────

function saveFamily(data) {
  _requireRole('admin', 'editor');
  const lock = _acquireLock();
  try {
    const sheet = _sheet(SHEETS.FAMILIES);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const now = new Date().toISOString();

    if (data.family_id) {
      const rowNum = _findRow(sheet, data.family_id);
      if (rowNum === -1) return { ok: false, error: 'Family not found.' };
      data.updated_at = now;
      sheet.getRange(rowNum, 1, 1, headers.length)
           .setValues([_objectToRow(headers, data)]);
      return { ok: true, family_id: data.family_id };
    } else {
      const user = _currentUser();
      data.family_id  = _nextId(sheet, 'FAM');
      data.created_by = user ? user.user_id : '';
      data.created_at = now;
      data.updated_at = now;
      sheet.appendRow(_objectToRow(headers, data));
      return { ok: true, family_id: data.family_id };
    }
  } finally {
    lock.releaseLock();
  }
}

function deleteFamily(familyId) {
  _requireRole('admin');
  if (!familyId) return { ok: false, error: 'family_id required.' };
  const lock = _acquireLock();
  try {
    const sheet = _sheet(SHEETS.FAMILIES);
    const rowNum = _findRow(sheet, familyId);
    if (rowNum === -1) return { ok: false, error: 'Family not found.' };
    sheet.deleteRow(rowNum);
    _deleteJunctionRows(SHEETS.FAMILY_CHILDREN, 'family_id', familyId);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// ─── Family_Children junction ─────────────────────────────────────────────────

function addFamilyChild(data) {
  _requireRole('admin', 'editor');
  const lock = _acquireLock();
  try {
    const sheet = _sheet(SHEETS.FAMILY_CHILDREN);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    // Guard against duplicate (family_id, child_id, relationship_type)
    const exists = _toObjects(sheet).some(
      r => String(r.family_id)         === String(data.family_id) &&
           String(r.child_id)          === String(data.child_id) &&
           String(r.relationship_type) === String(data.relationship_type)
    );
    if (exists) return { ok: false, error: 'This child link already exists.' };
    sheet.appendRow(_objectToRow(headers, data));
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function removeFamilyChild(data) {
  _requireRole('admin', 'editor');
  const lock = _acquireLock();
  try {
    const sheet = _sheet(SHEETS.FAMILY_CHILDREN);
    const values = sheet.getDataRange().getValues();
    const h = values[0];
    const famIdx = h.indexOf('family_id');
    const chiIdx = h.indexOf('child_id');
    const relIdx = h.indexOf('relationship_type');
    // Scan bottom-up so row-index drift from deleteRow doesn't affect earlier rows
    for (let i = values.length - 1; i >= 1; i--) {
      if (String(values[i][famIdx]) === String(data.family_id) &&
          String(values[i][chiIdx]) === String(data.child_id) &&
          String(values[i][relIdx]) === String(data.relationship_type)) {
        sheet.deleteRow(i + 1);
        return { ok: true };
      }
    }
    return { ok: false, error: 'Child link not found.' };
  } finally {
    lock.releaseLock();
  }
}

// ─── Events CRUD ──────────────────────────────────────────────────────────────

function listEvents() {
  return { ok: true, data: _toObjects(_sheet(SHEETS.EVENTS)) };
}

function saveEvent(data) {
  _requireRole('admin', 'editor');
  const lock = _acquireLock();
  try {
    const sheet = _sheet(SHEETS.EVENTS);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const now = new Date().toISOString();

    if (data.event_id) {
      const rowNum = _findRow(sheet, data.event_id);
      if (rowNum === -1) return { ok: false, error: 'Event not found.' };
      data.updated_at = now;
      sheet.getRange(rowNum, 1, 1, headers.length)
           .setValues([_objectToRow(headers, data)]);
      return { ok: true, event_id: data.event_id };
    } else {
      const user = _currentUser();
      data.event_id   = _nextId(sheet, 'EVT');
      data.created_by = user ? user.user_id : '';
      data.created_at = now;
      data.updated_at = now;
      sheet.appendRow(_objectToRow(headers, data));
      return { ok: true, event_id: data.event_id };
    }
  } finally {
    lock.releaseLock();
  }
}

function deleteEvent(eventId) {
  _requireRole('admin');
  if (!eventId) return { ok: false, error: 'event_id required.' };
  const lock = _acquireLock();
  try {
    const sheet = _sheet(SHEETS.EVENTS);
    const rowNum = _findRow(sheet, eventId);
    if (rowNum === -1) return { ok: false, error: 'Event not found.' };
    sheet.deleteRow(rowNum);
    _deleteJunctionRows(SHEETS.EVENT_PARTICIPANTS, 'event_id', eventId);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// ─── Event_Participants junction ──────────────────────────────────────────────

function addEventParticipant(data) {
  _requireRole('admin', 'editor');
  const lock = _acquireLock();
  try {
    const sheet = _sheet(SHEETS.EVENT_PARTICIPANTS);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const exists = _toObjects(sheet).some(
      r => String(r.event_id)  === String(data.event_id) &&
           String(r.person_id) === String(data.person_id)
    );
    if (exists) return { ok: false, error: 'Person is already a participant.' };
    sheet.appendRow(_objectToRow(headers, data));
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function removeEventParticipant(data) {
  _requireRole('admin', 'editor');
  const lock = _acquireLock();
  try {
    const sheet = _sheet(SHEETS.EVENT_PARTICIPANTS);
    const values = sheet.getDataRange().getValues();
    const h = values[0];
    const evtIdx  = h.indexOf('event_id');
    const persIdx = h.indexOf('person_id');
    for (let i = values.length - 1; i >= 1; i--) {
      if (String(values[i][evtIdx])  === String(data.event_id) &&
          String(values[i][persIdx]) === String(data.person_id)) {
        sheet.deleteRow(i + 1);
        return { ok: true };
      }
    }
    return { ok: false, error: 'Participant link not found.' };
  } finally {
    lock.releaseLock();
  }
}

// ─── Private cascade helpers ──────────────────────────────────────────────────

/**
 * Deletes every row in sheetName where the column named `colName`
 * equals `value`. Scans bottom-up to avoid row-index drift.
 */
function _deleteJunctionRows(sheetName, colName, value) {
  const sheet = _sheet(sheetName);
  const values = sheet.getDataRange().getValues();
  const colIdx = values[0].indexOf(colName);
  if (colIdx === -1) return;
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][colIdx]) === String(value)) {
      sheet.deleteRow(i + 1);
    }
  }
}

/**
 * Sets every cell in `colName` to '' where the current value === `value`.
 * Used to null-out FK references without deleting the parent row.
 */
function _nullifyColumn(sheetName, colName, value) {
  const sheet = _sheet(sheetName);
  const values = sheet.getDataRange().getValues();
  const colIdx = values[0].indexOf(colName);
  if (colIdx === -1) return;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][colIdx]) === String(value)) {
      sheet.getRange(i + 1, colIdx + 1).setValue('');
    }
  }
}
