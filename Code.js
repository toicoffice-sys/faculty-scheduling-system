// ─── DLSL Faculty Timetabler ───────────────────────────────────────────────
// Spreadsheet: https://docs.google.com/spreadsheets/d/1xIzCxULajIcIIqKVIoRmA8xdRGHKksEkDIkja5bObwo
// Apps Script: https://script.google.com/u/0/home/projects/1SUtVOEohLWRotcOZuWKQxuJscFFleuWL9LBw8-SsLya28vKcuQH-PdGx/edit

const SS_ID = '1xIzCxULajIcIIqKVIoRmA8xdRGHKksEkDIkja5bObwo';

const SH = {
  CONFIG:      'Config',
  USERS:       'Users',
  FACULTY:     'Faculty',
  BUILDINGS:   'Buildings',
  ROOMS:       'Rooms',
  SUBJECTS:    'Subjects',
  SECTIONS:    'Sections',
  PERIODS:     'Periods',
  LESSONS:     'Lessons',
  SCHEDULE:    'Schedule',
  CONSTRAINTS: 'Constraints',
};

const DLSL_DEPTS = ['SDA', 'SMIT', 'SHRIM', 'STHM', 'GEd'];
const DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];

// ─── Entry Point ───────────────────────────────────────────────────────────

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('DLSL Faculty Timetabler')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(fn) {
  return HtmlService.createHtmlOutputFromFile(fn).getContent();
}

// ─── Spreadsheet Utilities ─────────────────────────────────────────────────

// Cache spreadsheet object within a single execution (avoids 11× openById calls)
let _ss = null;
function ss_() {
  if (!_ss) _ss = SpreadsheetApp.openById(SS_ID);
  return _ss;
}

function getSheet(name) {
  const s = ss_();
  return s.getSheetByName(name) || s.insertSheet(name);
}

function sheetRows(name) {
  const sh = getSheet(name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
}

function upsertRow(sheetName, row) {
  const sh = getSheet(sheetName);
  const vals = sh.getDataRange().getValues();
  const id = row[0].toString();
  for (let i = 1; i < vals.length; i++) {
    if (vals[i][0] != null && vals[i][0].toString() === id) {
      sh.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return;
    }
  }
  sh.appendRow(row);
}

function deleteById(sheetName, id) {
  const sh = getSheet(sheetName);
  const vals = sh.getDataRange().getValues();
  for (let i = vals.length - 1; i >= 1; i--) {
    if (vals[i][0] != null && vals[i][0].toString() === id.toString()) {
      sh.deleteRow(i + 1);
    }
  }
}

function tryParse(v, fallback) {
  try { return (v && v !== '') ? JSON.parse(v.toString()) : fallback; }
  catch (e) { return fallback; }
}

function newId(prefix) {
  return (prefix || 'id') + '_' + Utilities.getUuid().replace(/-/g, '').substr(0, 10);
}

// ─── Auth ──────────────────────────────────────────────────────────────────

const ALLOWED_DOMAIN = 'dlsl.edu.ph';

function getCurrentUser() {
  const email = Session.getActiveUser().getEmail();
  if (!email) return null;

  // Restrict to DLSL email accounts only
  if (!email.toLowerCase().endsWith('@' + ALLOWED_DOMAIN)) {
    return { error: 'domain_restricted', email };
  }

  const rows = sheetRows(SH.USERS);
  for (const r of rows) {
    if (r[0] && r[0].toString().toLowerCase() === email.toLowerCase()) {
      return { email, role: r[1] || 'viewer', name: r[2] || email.split('@')[0] };
    }
  }

  // First user ever (first @dlsl.edu.ph account) becomes admin
  const sh = getSheet(SH.USERS);
  if (sh.getLastRow() <= 1) {
    sh.appendRow([email, 'admin', email.split('@')[0]]);
    return { email, role: 'admin', name: email.split('@')[0] };
  }

  return null; // not in Users sheet — needs admin to add them
}

function requireRole(minRole) {
  const user = getCurrentUser();
  if (!user) throw new Error('Not authorized — log in with your DLSL Google account.');
  const levels = { viewer: 0, editor: 1, admin: 2 };
  if ((levels[user.role] || 0) < (levels[minRole] || 0)) {
    throw new Error('Insufficient permissions. Required: ' + minRole);
  }
  return user;
}

// ─── Cache ─────────────────────────────────────────────────────────────────

const CACHE_KEY = 'dlsl_tt_boot_v1';
const CACHE_TTL = 90; // seconds — balance between freshness and speed

function getSharedData_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(CACHE_KEY);
  if (hit) {
    try { return JSON.parse(hit); } catch(e) {}
  }
  const data = {
    faculty:     readFaculty(),
    buildings:   readBuildings(),
    rooms:       readRooms(),
    subjects:    readSubjects(),
    sections:    readSections(),
    periods:     readPeriods(),
    lessons:     readLessons(),
    schedule:    readSchedule(),
    constraints: readConstraints(),
    config:      readConfig(),
  };
  try { cache.put(CACHE_KEY, JSON.stringify(data), CACHE_TTL); } catch(e) {}
  return data;
}

function invalidateCache_() {
  try { CacheService.getScriptCache().remove(CACHE_KEY); } catch(e) {}
}

// ─── Boot Data ─────────────────────────────────────────────────────────────

function getBootData() {
  const user = getCurrentUser();
  if (!user) return { error: 'not_authorized' };
  if (user.error === 'domain_restricted') return { error: 'domain_restricted', email: user.email };
  const shared = getSharedData_();
  return { user, ...shared };
}

// ─── Config ────────────────────────────────────────────────────────────────

function readConfig() {
  const cfg = {
    schoolName: 'De La Salle Lipa',
    academicYear: '2025-2026',
    term: '1st Semester',
  };
  sheetRows(SH.CONFIG).forEach(r => { if (r[0]) cfg[r[0].toString()] = r[1]; });
  return cfg;
}

function saveConfig(key, value) {
  requireRole('admin');
  const sh = getSheet(SH.CONFIG);
  const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (vals[i][0] === key) { sh.getRange(i + 1, 2).setValue(value); invalidateCache_(); return { ok: true }; }
  }
  sh.appendRow([key, value]);
  invalidateCache_();
  return { ok: true };
}

// ─── Faculty ───────────────────────────────────────────────────────────────

function readFaculty() {
  return sheetRows(SH.FACULTY).filter(r => r[0]).map(r => ({
    id: r[0].toString(),
    firstName: r[1] || '',
    lastName:  r[2] || '',
    email:     r[3] || '',
    department: r[4] || '',
    track:     r[5] || '',
    role:      r[6] || 'full-time',
    maxPerDay: Number(r[7]) || 5,
    maxPerWeek: Number(r[8]) || 22,
    subjects:     tryParse(r[9],  []),
    availability: tryParse(r[10], []),
    isAdviser: r[11] === true || r[11] === 'TRUE',
    notes: r[12] || '',
  }));
}

function saveFaculty(data) {
  requireRole('editor');
  if (!data.id) data.id = newId('fac');
  upsertRow(SH.FACULTY, [
    data.id, data.firstName || '', data.lastName || '', data.email || '',
    data.department || '', data.track || '', data.role || 'full-time',
    data.maxPerDay || 5, data.maxPerWeek || 22,
    JSON.stringify(data.subjects || []), JSON.stringify(data.availability || []),
    data.isAdviser ? 'TRUE' : 'FALSE', data.notes || '',
  ]);
  invalidateCache_();
  return { ok: true, id: data.id };
}

function deleteFaculty(id) {
  requireRole('editor');
  deleteById(SH.FACULTY, id);
  invalidateCache_();
  return { ok: true };
}

// ─── Buildings ─────────────────────────────────────────────────────────────

function readBuildings() {
  return sheetRows(SH.BUILDINGS).filter(r => r[0]).map(r => ({
    id: r[0].toString(), name: r[1] || '', address: r[2] || '',
  }));
}

function saveBuilding(data) {
  requireRole('editor');
  if (!data.id) data.id = newId('bldg');
  upsertRow(SH.BUILDINGS, [data.id, data.name, data.address || '']);
  invalidateCache_();
  return { ok: true, id: data.id };
}

function deleteBuilding(id) {
  requireRole('admin');
  deleteById(SH.BUILDINGS, id);
  invalidateCache_();
  return { ok: true };
}

// ─── Rooms ─────────────────────────────────────────────────────────────────

function readRooms() {
  return sheetRows(SH.ROOMS).filter(r => r[0]).map(r => ({
    id: r[0].toString(), name: r[1] || '', code: r[2] || '',
    capacity: Number(r[3]) || 35, type: r[4] || 'classroom',
    buildingId: r[5] ? r[5].toString() : '', features: tryParse(r[6], []),
  }));
}

function saveRoom(data) {
  requireRole('editor');
  if (!data.id) data.id = newId('room');
  upsertRow(SH.ROOMS, [
    data.id, data.name || '', data.code || '', data.capacity || 35,
    data.type || 'classroom', data.buildingId || '',
    JSON.stringify(data.features || []),
  ]);
  invalidateCache_();
  return { ok: true, id: data.id };
}

function deleteRoom(id) {
  requireRole('editor');
  deleteById(SH.ROOMS, id);
  invalidateCache_();
  return { ok: true };
}

// ─── Subjects ──────────────────────────────────────────────────────────────

function readSubjects() {
  return sheetRows(SH.SUBJECTS).filter(r => r[0]).map(r => ({
    id: r[0].toString(), name: r[1] || '', code: r[2] || '',
    department: r[3] || '', requiresRoom: r[4] || '', color: r[5] || '#2E7D32',
  }));
}

function saveSubject(data) {
  requireRole('editor');
  if (!data.id) data.id = newId('subj');
  upsertRow(SH.SUBJECTS, [
    data.id, data.name || '', data.code || '', data.department || '',
    data.requiresRoom || '', data.color || '#2E7D32',
  ]);
  invalidateCache_();
  return { ok: true, id: data.id };
}

function deleteSubject(id) {
  requireRole('editor');
  deleteById(SH.SUBJECTS, id);
  invalidateCache_();
  return { ok: true };
}

// ─── Sections ──────────────────────────────────────────────────────────────

function readSections() {
  return sheetRows(SH.SECTIONS).filter(r => r[0]).map(r => ({
    id: r[0].toString(), name: r[1] || '', year: Number(r[2]) || 1,
    size: Number(r[3]) || 35, program: r[4] || '',
  }));
}

function saveSection(data) {
  requireRole('editor');
  if (!data.id) data.id = newId('sec');
  upsertRow(SH.SECTIONS, [
    data.id, data.name || '', data.year || 1, data.size || 35, data.program || '',
  ]);
  invalidateCache_();
  return { ok: true, id: data.id };
}

function deleteSection(id) {
  requireRole('editor');
  deleteById(SH.SECTIONS, id);
  invalidateCache_();
  return { ok: true };
}

// ─── Periods ───────────────────────────────────────────────────────────────

function readPeriods() {
  const rows = sheetRows(SH.PERIODS).filter(r => r[0]);
  if (rows.length === 0) return seedPeriods();
  return rows.map(r => ({
    id: r[0].toString(), name: r[1] || '', day: r[2] || '',
    startTime: r[3] || '', endTime: r[4] || '',
    isBreak: r[5] === true || r[5] === 'TRUE',
    isLunch: r[6] === true || r[6] === 'TRUE',
  }));
}

function seedPeriods() {
  const slots = [
    { name: 'P1',    start: '07:30', end: '09:00', isLunch: false },
    { name: 'P2',    start: '09:00', end: '10:30', isLunch: false },
    { name: 'P3',    start: '10:30', end: '12:00', isLunch: false },
    { name: 'Lunch', start: '12:00', end: '13:00', isLunch: true  },
    { name: 'P4',    start: '13:00', end: '14:30', isLunch: false },
    { name: 'P5',    start: '14:30', end: '16:00', isLunch: false },
    { name: 'P6',    start: '16:00', end: '17:30', isLunch: false },
    { name: 'P7',    start: '17:30', end: '19:00', isLunch: false },
  ];
  const periods = [];
  let n = 1;
  DAYS.forEach(day => {
    slots.forEach(s => {
      const id = 'p' + (n++);
      upsertRow(SH.PERIODS, [id, s.name, day, s.start, s.end, false, s.isLunch]);
      periods.push({ id, name: s.name, day, startTime: s.start, endTime: s.end, isBreak: false, isLunch: s.isLunch });
    });
  });
  return periods;
}

function savePeriod(data) {
  requireRole('admin');
  if (!data.id) data.id = newId('per');
  upsertRow(SH.PERIODS, [
    data.id, data.name, data.day, data.startTime, data.endTime,
    data.isBreak || false, data.isLunch || false,
  ]);
  return { ok: true, id: data.id };
}

function deletePeriod(id) {
  requireRole('admin');
  deleteById(SH.PERIODS, id);
  return { ok: true };
}

// ─── Lessons ───────────────────────────────────────────────────────────────

function readLessons() {
  return sheetRows(SH.LESSONS).filter(r => r[0]).map(r => ({
    id: r[0].toString(),
    subjectId:       r[1] ? r[1].toString() : '',
    teacherId:       r[2] ? r[2].toString() : '',
    sectionId:       r[3] ? r[3].toString() : '',
    preferredRoomId: r[4] ? r[4].toString() : '',
    periodsRequired: Number(r[5]) || 1,
    isDouble: r[6] === true || r[6] === 'TRUE',
  }));
}

function saveLesson(data) {
  requireRole('editor');
  if (!data.id) data.id = newId('les');
  upsertRow(SH.LESSONS, [
    data.id, data.subjectId || '', data.teacherId || '', data.sectionId || '',
    data.preferredRoomId || '', data.periodsRequired || 1, data.isDouble ? 'TRUE' : 'FALSE',
  ]);
  return { ok: true, id: data.id };
}

function deleteLesson(id) {
  requireRole('editor');
  deleteById(SH.LESSONS, id);
  const sh = getSheet(SH.SCHEDULE);
  const vals = sh.getDataRange().getValues();
  for (let i = vals.length - 1; i >= 1; i--) {
    if (vals[i][1] && vals[i][1].toString() === id) sh.deleteRow(i + 1);
  }
  return { ok: true };
}

// ─── Schedule ──────────────────────────────────────────────────────────────

function readSchedule() {
  return sheetRows(SH.SCHEDULE).filter(r => r[0]).map(r => ({
    id: r[0].toString(),
    lessonId: r[1] ? r[1].toString() : '',
    periodId: r[2] ? r[2].toString() : '',
    roomId:   r[3] ? r[3].toString() : '',
    status:   r[4] || 'placed',
    lockedBy: r[5] || '',
  }));
}

function saveScheduleEntry(data) {
  requireRole('editor');
  if (!data.id) data.id = newId('sch');
  upsertRow(SH.SCHEDULE, [
    data.id, data.lessonId, data.periodId, data.roomId,
    data.status || 'placed', data.lockedBy || '',
  ]);
  invalidateCache_();
  return { ok: true, id: data.id };
}

function deleteScheduleEntry(id) {
  requireRole('editor');
  deleteById(SH.SCHEDULE, id);
  invalidateCache_();
  return { ok: true };
}

function clearSchedule(keepLocked) {
  requireRole('admin');
  const sh = getSheet(SH.SCHEDULE);
  const vals = sh.getDataRange().getValues();
  for (let i = vals.length - 1; i >= 1; i--) {
    if (!keepLocked || vals[i][4] !== 'locked') sh.deleteRow(i + 1);
  }
  invalidateCache_();
  return { ok: true };
}

// ─── Constraints ───────────────────────────────────────────────────────────

function readConstraints() {
  return sheetRows(SH.CONSTRAINTS).filter(r => r[0]).map(r => ({
    id: r[0].toString(), type: r[1] || '', severity: r[2] || 'soft',
    description: r[3] || '', teacherId: r[4] ? r[4].toString() : '',
    sectionId: r[5] ? r[5].toString() : '', subjectId: r[6] ? r[6].toString() : '',
    value: r[7] || '',
  }));
}

function saveConstraint(data) {
  requireRole('editor');
  if (!data.id) data.id = newId('con');
  upsertRow(SH.CONSTRAINTS, [
    data.id, data.type || '', data.severity || 'soft', data.description || '',
    data.teacherId || '', data.sectionId || '', data.subjectId || '', data.value || '',
  ]);
  invalidateCache_();
  return { ok: true, id: data.id };
}

function deleteConstraint(id) {
  requireRole('editor');
  deleteById(SH.CONSTRAINTS, id);
  invalidateCache_();
  return { ok: true };
}

// ─── Users ─────────────────────────────────────────────────────────────────

function getUsers() {
  requireRole('admin');
  return sheetRows(SH.USERS).filter(r => r[0]).map(r => ({
    email: r[0], role: r[1] || 'viewer', name: r[2] || r[0],
  }));
}

function saveUser(data) {
  requireRole('admin');
  const sh = getSheet(SH.USERS);
  const vals = sh.getDataRange().getValues();
  const email = data.email.toLowerCase();
  for (let i = 1; i < vals.length; i++) {
    if (vals[i][0] && vals[i][0].toString().toLowerCase() === email) {
      sh.getRange(i + 1, 1, 1, 3).setValues([[data.email, data.role, data.name || data.email.split('@')[0]]]);
      return { ok: true };
    }
  }
  sh.appendRow([data.email, data.role || 'viewer', data.name || data.email.split('@')[0]]);
  return { ok: true };
}

function removeUser(email) {
  const me = requireRole('admin');
  if (me.email.toLowerCase() === email.toLowerCase()) throw new Error('Cannot remove your own account.');
  deleteById(SH.USERS, email);
  return { ok: true };
}

// ─── Auto-Scheduler ────────────────────────────────────────────────────────

function autoSchedule() {
  requireRole('editor');

  const faculty     = readFaculty();
  const rooms       = readRooms();
  const periods     = readPeriods();
  const lessons     = readLessons();
  const existing    = readSchedule();
  const constraints = readConstraints();

  const facMap  = {}, roomMap = {};
  faculty.forEach(f => { facMap[f.id] = f; });
  rooms.forEach(r => { roomMap[r.id] = r; });

  const teaching = periods.filter(p => !p.isBreak && !p.isLunch);

  // Build occupation index: periodId → { teachers, rooms, sections }
  const occ = {};
  teaching.forEach(p => { occ[p.id] = { teachers: {}, rooms: {}, sections: {} }; });

  const locked = existing.filter(e => e.status === 'locked');
  const lockedLessonIds = new Set(locked.map(e => e.lessonId));

  locked.forEach(e => {
    const les = lessons.find(l => l.id === e.lessonId);
    if (!les || !occ[e.periodId]) return;
    occ[e.periodId].teachers[les.teacherId] = true;
    occ[e.periodId].rooms[e.roomId] = true;
    occ[e.periodId].sections[les.sectionId] = true;
  });

  // Track weekly load per teacher
  const weeklyLoad = {};
  faculty.forEach(f => { weeklyLoad[f.id] = 0; });
  locked.forEach(e => {
    const les = lessons.find(l => l.id === e.lessonId);
    if (les) weeklyLoad[les.teacherId] = (weeklyLoad[les.teacherId] || 0) + 1;
  });

  // Priority: fewer available slots → schedule first
  const unplaced = lessons.filter(l => !lockedLessonIds.has(l.id));
  unplaced.sort((a, b) => {
    const fa = facMap[a.teacherId], fb = facMap[b.teacherId];
    const avA = fa ? fa.availability.reduce((n, s) => n + (s.periodIds ? s.periodIds.length : 0), 0) : 999;
    const avB = fb ? fb.availability.reduce((n, s) => n + (s.periodIds ? s.periodIds.length : 0), 0) : 999;
    if (avA !== avB) return avA - avB;
    return (b.periodsRequired || 1) - (a.periodsRequired || 1);
  });

  const newEntries = [...locked];

  unplaced.forEach(lesson => {
    const teacher    = facMap[lesson.teacherId];
    const target     = lesson.periodsRequired || 1;
    let placed       = 0;
    const usedDays   = {};

    const lessonCons = constraints.filter(c =>
      c.teacherId === lesson.teacherId ||
      c.sectionId === lesson.sectionId ||
      c.subjectId === lesson.subjectId
    );
    const spreadCon     = lessonCons.find(c => c.type === 'spread_across_days');
    const mustOnDay     = lessonCons.find(c => c.type === 'must_be_on_day');
    const mustNotOnDay  = lessonCons.find(c => c.type === 'must_not_be_on_day');
    const roomTypeCon   = constraints.find(c => c.type === 'room_required' && c.subjectId === lesson.subjectId);

    for (const period of teaching) {
      if (placed >= target) break;
      const slot = occ[period.id];
      if (!slot) continue;

      // Hard conflicts
      if (slot.teachers[lesson.teacherId]) continue;
      if (slot.sections[lesson.sectionId]) continue;

      // Teacher availability
      if (teacher && teacher.availability && teacher.availability.length > 0) {
        const dayAvail = teacher.availability.find(a => a.day === period.day);
        if (!dayAvail) continue;
        const pids = dayAvail.periodIds || [];
        if (!pids.includes(period.id)) continue;
      }

      // Weekly / daily load caps
      if (teacher && (weeklyLoad[lesson.teacherId] || 0) >= (teacher.maxPerWeek || 22)) continue;
      const dayCount = newEntries.filter(e => {
        const p = periods.find(pp => pp.id === e.periodId);
        const l = lessons.find(ll => ll.id === e.lessonId);
        return l && l.teacherId === lesson.teacherId && p && p.day === period.day;
      }).length;
      if (teacher && dayCount >= (teacher.maxPerDay || 5)) continue;

      // Constraint: must_be_on_day
      if (mustOnDay && mustOnDay.value && period.day !== mustOnDay.value) continue;
      // Constraint: must_not_be_on_day
      if (mustNotOnDay && mustNotOnDay.value && period.day === mustNotOnDay.value) continue;

      // Soft: spread across days — skip if better day exists
      if (spreadCon && usedDays[period.day]) {
        const betterExists = teaching.some(pp =>
          !usedDays[pp.day] &&
          !occ[pp.id].teachers[lesson.teacherId] &&
          !occ[pp.id].sections[lesson.sectionId]
        );
        if (betterExists) continue;
      }

      // Find a room
      let roomId = lesson.preferredRoomId;
      if (!roomId || slot.rooms[roomId]) {
        const requiredType = roomTypeCon ? roomTypeCon.value : null;
        const avail = rooms.find(r => !slot.rooms[r.id] && (!requiredType || r.type === requiredType));
        if (!avail) {
          if (requiredType) continue; // no matching room — skip this slot
          const anyFree = rooms.find(r => !slot.rooms[r.id]);
          if (!anyFree) continue;
          roomId = anyFree.id;
        } else {
          roomId = avail.id;
        }
      }

      // Place!
      const entry = { id: newId('sch'), lessonId: lesson.id, periodId: period.id, roomId, status: 'placed', lockedBy: '' };
      newEntries.push(entry);
      slot.teachers[lesson.teacherId] = true;
      slot.rooms[roomId] = true;
      slot.sections[lesson.sectionId] = true;
      weeklyLoad[lesson.teacherId] = (weeklyLoad[lesson.teacherId] || 0) + 1;
      usedDays[period.day] = true;
      placed++;
    }
  });

  // Persist: clear non-locked rows, re-write all
  const sh = getSheet(SH.SCHEDULE);
  const allVals = sh.getDataRange().getValues();
  for (let i = allVals.length - 1; i >= 1; i--) {
    if (allVals[i][4] !== 'locked') sh.deleteRow(i + 1);
  }
  newEntries.filter(e => e.status !== 'locked').forEach(e => {
    sh.appendRow([e.id, e.lessonId, e.periodId, e.roomId, e.status, e.lockedBy || '']);
  });

  invalidateCache_();
  return { ok: true, placed: newEntries.filter(e => e.status !== 'locked').length, total: lessons.length, entries: newEntries };
}

// ─── Initialize Sheets ─────────────────────────────────────────────────────

function initializeSheets() {
  const spreadsheet = ss_();
  const GREEN = '#1B5E20', WHITE = '#FFFFFF';

  const defs = {
    [SH.CONFIG]:      ['Key', 'Value'],
    [SH.USERS]:       ['Email', 'Role', 'Name'],
    [SH.FACULTY]:     ['ID', 'FirstName', 'LastName', 'Email', 'Department', 'Track', 'Role', 'MaxPerDay', 'MaxPerWeek', 'Subjects', 'Availability', 'IsAdviser', 'Notes'],
    [SH.BUILDINGS]:   ['ID', 'Name', 'Address'],
    [SH.ROOMS]:       ['ID', 'Name', 'Code', 'Capacity', 'Type', 'BuildingID', 'Features'],
    [SH.SUBJECTS]:    ['ID', 'Name', 'Code', 'Department', 'RequiresRoom', 'Color'],
    [SH.SECTIONS]:    ['ID', 'Name', 'Year', 'Size', 'Program'],
    [SH.PERIODS]:     ['ID', 'Name', 'Day', 'StartTime', 'EndTime', 'IsBreak', 'IsLunch'],
    [SH.LESSONS]:     ['ID', 'SubjectID', 'TeacherID', 'SectionID', 'PreferredRoomID', 'PeriodsRequired', 'IsDouble'],
    [SH.SCHEDULE]:    ['ID', 'LessonID', 'PeriodID', 'RoomID', 'Status', 'LockedBy'],
    [SH.CONSTRAINTS]: ['ID', 'Type', 'Severity', 'Description', 'TeacherID', 'SectionID', 'SubjectID', 'Value'],
  };

  Object.entries(defs).forEach(([name, headers]) => {
    let sh = spreadsheet.getSheetByName(name);
    if (!sh) sh = spreadsheet.insertSheet(name);
    const hr = sh.getRange(1, 1, 1, headers.length);
    hr.setValues([headers]).setBackground(GREEN).setFontColor(WHITE).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 160);
  });

  const cfgSh = spreadsheet.getSheetByName(SH.CONFIG);
  if (cfgSh.getLastRow() <= 1) {
    cfgSh.appendRow(['schoolName', 'De La Salle Lipa']);
    cfgSh.appendRow(['academicYear', '2025-2026']);
    cfgSh.appendRow(['term', '1st Semester']);
  }

  readPeriods(); // seeds default periods if empty
  return { ok: true };
}

// ─── CSV Export ────────────────────────────────────────────────────────────

function exportScheduleCSV() {
  requireRole('viewer');
  const schedule  = readSchedule();
  const lessons   = readLessons();
  const faculty   = readFaculty();
  const rooms     = readRooms();
  const periods   = readPeriods();
  const subjects  = readSubjects();
  const sections  = readSections();

  const fM = {}, subM = {}, secM = {}, rM = {}, perM = {};
  faculty.forEach(f  => { fM[f.id]   = f;   });
  subjects.forEach(s => { subM[s.id] = s;   });
  sections.forEach(s => { secM[s.id] = s;   });
  rooms.forEach(r    => { rM[r.id]   = r;   });
  periods.forEach(p  => { perM[p.id] = p;   });

  const rows = ['Day,Period,Time,Subject,Code,Teacher,Section,Room,Status'];
  schedule.forEach(e => {
    const les = lessons.find(l => l.id === e.lessonId);
    if (!les) return;
    const per = perM[e.periodId] || {};
    const fac = fM[les.teacherId] || {};
    const sub = subM[les.subjectId] || {};
    const sec = secM[les.sectionId] || {};
    const room = rM[e.roomId] || {};
    rows.push([
      per.day, per.name,
      (per.startTime || '') + '-' + (per.endTime || ''),
      sub.name, sub.code,
      (fac.lastName || '') + ', ' + (fac.firstName || ''),
      sec.name, room.code, e.status,
    ].map(v => '"' + (v || '').toString().replace(/"/g, '""') + '"').join(','));
  });
  return rows.join('\n');
}
