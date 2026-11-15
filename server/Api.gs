/**
 * 대표님 개인 비서 — 데이터 서버 v1.1 (Google Apps Script, 독립 프로젝트 '비서 API')
 *
 * 설치형 앱(https://jerryjmj.github.io/biseo/)이 이 서버와 데이터를 주고받습니다.
 * 화면은 GitHub에, 데이터는 지금처럼 '비서_DB' 구글시트에 있습니다.
 *
 * 보안: 배포는 '모든 사용자'로 열지만, 모든 요청은 비밀 열쇠(API_KEY)가 맞아야만 처리합니다.
 *       열쇠는 이 스크립트의 속성에만 저장되고 코드·GitHub에는 없습니다.
 *       열쇠를 바꾸려면 rotateKey 를 실행하세요(기존 열쇠는 즉시 무효).
 */
const DB_ID = '1u6AIaOUzAVofV9047ZyqtTEdjAAbNwr9_Y7iHnxPkno'; // 비서_DB

/* ---------- 앱이 부를 수 있는 함수 목록(이 밖의 함수는 호출 불가) ---------- */
const API = {
  load: () => load(),
  events: (a, b) => events(a, b),
  calendars: () => calendars(),
  setCalHidden: (id, hide) => setCalHidden(id, hide),
  sugarman: (force) => sugarman(force),
  addTask: (t) => addTask(t), updateTask: (id, p) => updateTask(id, p), deleteTask: (id) => deleteTask(id),
  addGoal: (g) => addGoal(g), updateGoal: (id, p) => updateGoal(id, p), deleteGoal: (id) => deleteGoal(id),
  addRec: (r) => addRec(r), updateRec: (id, p) => updateRec(id, p), deleteRec: (id) => deleteRec(id),
  addIdea: (i) => addIdea(i), deleteIdea: (id) => deleteIdea(id),
  addIdeaDone: (d) => addIdeaDone(d), deleteIdeaDone: (id) => deleteIdeaDone(id),
  briefing: () => brief_(),
  ping: () => ({ ok: true, at: now_(), ver: 'v1.1' })
};

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function keyOk_(k) {
  const key = PropertiesService.getScriptProperties().getProperty('API_KEY');
  if (!key || !k || String(k).length !== key.length) return false;
  let diff = 0;                                  // 길이가 같을 때 끝까지 비교(시간차 공격 방지)
  for (let i = 0; i < key.length; i++) diff |= key.charCodeAt(i) ^ String(k).charCodeAt(i);
  return diff === 0;
}

/** 앱 → 서버: POST 본문 {k, fn, args} (Content-Type text/plain) */
function doPost(e) {
  let req = {};
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: '잘못된 요청' }); }
  if (!keyOk_(req.k)) { Utilities.sleep(800); return json_({ ok: false, error: '열쇠가 맞지 않습니다', auth: true }); }
  const f = Object.prototype.hasOwnProperty.call(API, req.fn) ? API[req.fn] : null;
  if (!f) return json_({ ok: false, error: '없는 기능: ' + req.fn });
  try {
    return json_({ ok: true, data: f.apply(null, Array.isArray(req.args) ? req.args : []) });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/** 주소를 브라우저로 직접 열었을 때 — 아무 데이터도 주지 않음 */
function doGet() { return json_({ ok: false, error: '비서 API — 앱에서만 사용합니다' }); }

/** 설치 때 1회 실행: 권한 승인 + 비밀 열쇠 생성(실행 로그에 표시) */
function setup() {
  const ss = ss_();
  sheet_(TASK); sheet_(GOAL); sheet_(REC); sheet_(IDEA); sheet_(IDEA_DONE);
  DriveApp.getRootFolder();
  CalendarApp.getDefaultCalendar();
  SpreadsheetApp.openById(SUGAR_ID).getName();
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('API_KEY')) p.setProperty('API_KEY', newKey_());
  Logger.log('비서_DB 연결 확인: ' + ss.getName());
  Logger.log('비밀 열쇠: ' + p.getProperty('API_KEY'));
  Logger.log('이 열쇠를 앱 첫 화면의 "비밀 열쇠" 칸에 붙여넣으세요. 다른 사람에게 보여주지 마세요.');
}

/** 열쇠 바꾸기 — 실행하면 새 열쇠가 로그에 나오고, 기존 열쇠는 즉시 쓸 수 없게 됩니다 */
function rotateKey() {
  const k = newKey_();
  PropertiesService.getScriptProperties().setProperty('API_KEY', k);
  Logger.log('새 비밀 열쇠: ' + k);
}

function newKey_() {
  return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');   // 64자리
}

const TZ = 'Asia/Seoul';
const SUGAR_ID = '1L2ILu4wcuSnSMUlqR9nHID5xkkBVYRAbyiCAf9emLU4'; // 슈가맨워크_계약관리_통합시트
const TAGS = ['회사', '개인', '슈가맨'];

const TASK = {
  name: '할일',
  keys: ['id', 'text', 'due', 'tag', 'status', 'created', 'done'],
  labels: ['id', '내용', '마감', '분류', '상태', '생성', '완료']
};
const GOAL = {
  name: '목표',
  keys: ['id', 'type', 'period', 'text', 'status', 'created'],
  labels: ['id', '유형', '기간', '내용', '상태', '생성']
};
const IDEA = {
  name: '아이디어수신',
  keys: ['id', 'text', 'created', 'status'],
  labels: ['id', '내용', '생성', '상태']
};
const IDEA_DONE = {
  name: '아이디어완료',
  keys: ['id', 'no', 'title', 'created', 'status'],
  labels: ['id', 'No', '제목', '요청', '상태']
};
const TOP_FILE = '비서_아이디어TOP.csv'; // 매일 아침 Claude가 아이디어 뱅크에서 골라 만들어 두는 파일
const BRIEF_FILE = '비서_브리핑.json';     // 매일 아침 7시 통합 브리핑이 만들어 두는 파일(오늘의 3가지·한마디)
const REC = {
  name: '기록',
  keys: ['id', 'kind', 'date', 'time', 'mood', 'text', 'created'],
  labels: ['id', '종류', '날짜', '시간', '기분', '내용', '생성']
};

/* ---------- 내부 도구 ---------- */
function ss_() { return SpreadsheetApp.openById(DB_ID); }

function sheet_(def) {
  const ss = ss_();
  let sh = ss.getSheetByName(def.name);
  if (!sh) {
    sh = ss.insertSheet(def.name);
    sh.getRange(1, 1, sh.getMaxRows(), def.keys.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, def.keys.length).setValues([def.labels]).setFontWeight('bold');
    sh.setFrozenRows(1);
    const wide = ['text', 'title'].map(k => def.keys.indexOf(k)).filter(i => i >= 0)[0];
    if (wide != null) sh.setColumnWidth(wide + 1, 360);
  }
  return sh;
}

function fmt_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return v == null ? '' : String(v);
}
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }

function rows_(def) {
  const sh = sheet_(def);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, def.keys.length).getValues()
    .filter(r => r[0] !== '')
    .map(r => {
      const o = {};
      def.keys.forEach((k, j) => { o[k] = fmt_(r[j]); });
      return o;
    });
}

function append_(def, o) {
  const sh = sheet_(def);
  const r = sh.getLastRow() + 1;
  if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 200);
  sh.getRange(r, 1, 1, def.keys.length)
    .setNumberFormat('@')
    .setValues([def.keys.map(k => (o[k] == null ? '' : String(o[k])))]);
  return o;
}

/** patch === null 이면 삭제 */
function update_(def, id, patch) {
  const sh = sheet_(def);
  const last = sh.getLastRow();
  if (last < 2) throw new Error('항목을 찾을 수 없습니다');
  const ids = sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0]));
  const i = ids.indexOf(String(id));
  if (i < 0) throw new Error('항목을 찾을 수 없습니다');
  if (patch === null) { sh.deleteRow(i + 2); return null; }
  const rg = sh.getRange(i + 2, 1, 1, def.keys.length);
  const cur = rg.getValues()[0];
  const o = {};
  def.keys.forEach((k, j) => { o[k] = fmt_(cur[j]); });
  Object.keys(patch).forEach(k => { if (def.keys.indexOf(k) > 0) o[k] = patch[k]; });
  rg.setNumberFormat('@').setValues([def.keys.map(k => (o[k] == null ? '' : String(o[k])))]);
  return o;
}

function locked_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function clean_(s, max) { return String(s == null ? '' : s).trim().slice(0, max || 500); }
function isDate_(s) { return /^\d{4}-\d{2}-\d{2}$/.test(s); }
function tag_(t) { return TAGS.indexOf(t) >= 0 ? t : '회사'; }

/* ---------- 데이터 함수 ---------- */
function load() {
  const goals = rows_(GOAL).map(g => {
    if (g.type === '월') g.period = g.period.slice(0, 7);
    if (g.type === '연') g.period = g.period.slice(0, 4);
    return g;
  });
  let top = { at: '', items: [], done: [] };
  try { top = ideaTop_(); } catch (e) { top.err = String(e.message || e); }
  let inbox = rows_(IDEA);
  const done = top.done.filter(id => inbox.some(r => r.id === id && r.status === '대기'));
  if (done.length) locked_(() => done.forEach(id => update_(IDEA, id, { status: '등록됨' })));
  inbox = inbox.filter(r => r.status === '대기' && done.indexOf(r.id) < 0);
  // 앱에서 완료 처리한 아이디어 — 다음 동기화 때 아이디어 뱅크 '완료 목록'으로 옮겨집니다
  let fin = rows_(IDEA_DONE);
  const finDone = top.done.filter(id => fin.some(r => r.id === id && r.status === '대기'));
  if (finDone.length) locked_(() => finDone.forEach(id => update_(IDEA_DONE, id, { status: '반영됨' })));
  fin = fin.filter(r => r.status === '대기' && finDone.indexOf(r.id) < 0);
  const items = top.items.filter(x => !fin.some(r => r.title === x.title));
  return { tasks: rows_(TASK), goals: goals, recs: rows_(REC), ideas: { top: items, at: top.at, err: top.err || '', inbox: inbox, fin: fin }, brief: brief_() };
}

/** 아침 브리핑 파일 읽기 — 가장 최근 것. 형식이 틀려도 앱이 멈추지 않도록 필요한 칸만 골라 돌려줍니다 */
function brief_() {
  try {
    const it = DriveApp.getFilesByName(BRIEF_FILE);
    let f = null;
    while (it.hasNext()) {
      const x = it.next();
      if (x.isTrashed()) continue;
      if (!f || x.getLastUpdated() > f.getLastUpdated()) f = x;
    }
    if (!f) return { date: '' };
    const o = JSON.parse(f.getBlob().getDataAsString('UTF-8').replace(/^\uFEFF/, ''));
    const s = (v, n) => clean_(v, n || 300);
    return {
      date: isDate_(o.date) ? o.date : '',
      created: s(o.created, 10) || Utilities.formatDate(f.getLastUpdated(), TZ, 'HH:mm'),
      weekly: (Array.isArray(o.weekly) ? o.weekly : []).slice(0, 3).map(x => s(x)),
      three: (Array.isArray(o.three) ? o.three : []).slice(0, 3).map(x => ({ title: s(x.title, 120), why: s(x.why), action: s(x.action), kind: s(x.kind, 10) })),
      one: s(o.one, 200),
      word: s(o.word, 300)
    };
  } catch (e) { return { date: '', err: String(e.message || e) }; }
}

/** 아이디어 TOP 파일 읽기 — 같은 이름 파일이 여럿이면 가장 최근 것 */
function ideaTop_() {
  const it = DriveApp.getFilesByName(TOP_FILE);
  let f = null;
  while (it.hasNext()) {
    const x = it.next();
    if (x.isTrashed()) continue;
    if (!f || x.getLastUpdated() > f.getLastUpdated()) f = x;
  }
  const out = { at: '', items: [], done: [] };
  if (!f) return out;
  out.at = Utilities.formatDate(f.getLastUpdated(), TZ, 'M/d HH:mm');
  const rows = Utilities.parseCsv(f.getBlob().getDataAsString('UTF-8').replace(/^\uFEFF/, ''));
  const H = rows[0].map(v => String(v).trim());
  const c = k => H.indexOf(k);
  rows.slice(1).forEach(r => {
    const kind = String(r[c('구분')] || '').trim();
    if (kind === 'TOP') out.items.push({ no: r[c('No')] || '', title: r[c('제목')] || '', gist: r[c('요지')] || '', pri: r[c('우선순위')] || '', st: r[c('상태')] || '' });
    else if (kind === '처리' && r[c('수신id')]) out.done.push(String(r[c('수신id')]).trim());
  });
  return out;
}

/* 아이디어 던지기 — 다음 날 아침 Claude가 아이디어 뱅크로 옮깁니다 */
function addIdea(i) {
  return locked_(() => {
    const text = clean_(i.text, 2000);
    if (!text) throw new Error('내용이 비었습니다');
    return append_(IDEA, { id: clean_(i.id) || Utilities.getUuid().slice(0, 8), text: text, created: clean_(i.created) || now_(), status: '대기' });
  });
}
function deleteIdea(id) { return locked_(() => update_(IDEA, id, null)); }

/* 아이디어 완료 처리 — 다음 날 아침 Claude가 아이디어 뱅크에서 완료 목록으로 옮깁니다 */
function addIdeaDone(d) {
  return locked_(() => append_(IDEA_DONE, { id: clean_(d.id) || Utilities.getUuid().slice(0, 8), no: clean_(d.no, 10), title: clean_(d.title, 300), created: clean_(d.created) || now_(), status: '대기' }));
}
function deleteIdeaDone(id) { return locked_(() => update_(IDEA_DONE, id, null)); }

/* 할 일 */
function addTask(t) {
  return locked_(() => {
    const text = clean_(t.text);
    if (!text) throw new Error('내용이 비었습니다');
    return append_(TASK, {
      id: clean_(t.id) || Utilities.getUuid().slice(0, 8),
      text: text,
      due: isDate_(t.due) ? t.due : today_(),
      tag: tag_(t.tag),
      status: t.status === '완료' ? '완료' : '진행',
      created: clean_(t.created) || now_(),
      done: t.status === '완료' ? (clean_(t.done) || now_()) : ''
    });
  });
}

function updateTask(id, p) {
  return locked_(() => {
    const patch = {};
    if (p.text != null && clean_(p.text)) patch.text = clean_(p.text);
    if (p.due != null && isDate_(p.due)) patch.due = p.due;
    if (p.tag != null) patch.tag = tag_(p.tag);
    if (p.status != null) {
      patch.status = p.status === '완료' ? '완료' : '진행';
      patch.done = patch.status === '완료' ? now_() : '';
    }
    return update_(TASK, id, patch);
  });
}

function deleteTask(id) { return locked_(() => update_(TASK, id, null)); }

/* 목표 */
function addGoal(g) {
  return locked_(() => {
    const text = clean_(g.text);
    if (!text) throw new Error('내용이 비었습니다');
    return append_(GOAL, {
      id: clean_(g.id) || Utilities.getUuid().slice(0, 8),
      type: g.type === '연' ? '연' : '월',
      period: clean_(g.period),
      text: text,
      status: g.status === '달성' ? '달성' : '진행',
      created: clean_(g.created) || now_()
    });
  });
}

function updateGoal(id, p) {
  return locked_(() => {
    const patch = {};
    if (p.text != null && clean_(p.text)) patch.text = clean_(p.text);
    if (p.status != null) patch.status = p.status === '달성' ? '달성' : '진행';
    return update_(GOAL, id, patch);
  });
}

function deleteGoal(id) { return locked_(() => update_(GOAL, id, null)); }

/* 기록 — 일기 / 슈가맨 메모 */
function addRec(r) {
  return locked_(() => {
    const text = clean_(r.text, 5000);
    if (!text) throw new Error('내용이 비었습니다');
    return append_(REC, {
      id: clean_(r.id) || Utilities.getUuid().slice(0, 8),
      kind: ['슈가맨', '사업', '한마디'].indexOf(r.kind) >= 0 ? r.kind : '일기',   // 한마디 = 브리핑 탭에서 별표한 말
      date: isDate_(r.date) ? r.date : today_(),
      time: clean_(r.time, 5) || Utilities.formatDate(new Date(), TZ, 'HH:mm'),
      mood: clean_(r.mood, 8),
      text: text,
      created: clean_(r.created) || now_()
    });
  });
}

function updateRec(id, p) {
  return locked_(() => {
    const patch = {};
    if (p.text != null && clean_(p.text, 5000)) patch.text = clean_(p.text, 5000);
    if (p.mood != null) patch.mood = clean_(p.mood, 8);
    return update_(REC, id, patch);
  });
}

function deleteRec(id) { return locked_(() => update_(REC, id, null)); }

/* 캘린더 — start~end(yyyy-MM-dd), 캘린더 앱에서 '표시' 체크된 캘린더 전체 */
function events(start, end) {
  const s = Utilities.parseDate(start + ' 00:00:00', TZ, 'yyyy-MM-dd HH:mm:ss');
  const e = Utilities.parseDate(end + ' 23:59:59', TZ, 'yyyy-MM-dd HH:mm:ss');
  const f = (d, p) => Utilities.formatDate(d, TZ, p);
  const seen = {};
  const out = [];
  const hidden = hiddenCals_();
  CalendarApp.getAllCalendars().forEach(c => {
    if (hidden.indexOf(c.getId()) >= 0) return;
    const color = c.getColor();
    c.getEvents(s, e).forEach(ev => {
      const allDay = ev.isAllDayEvent();
      const st = allDay ? ev.getAllDayStartDate() : ev.getStartTime();
      const en = allDay ? ev.getAllDayEndDate() : ev.getEndTime();
      const endIncl = new Date(Math.max(st.getTime(), en.getTime() - 1));
      const key = ev.getId() + '|' + f(st, 'yyyyMMddHHmm');
      if (seen[key]) return;
      seen[key] = 1;
      out.push({
        title: ev.getTitle() || '(제목 없음)',
        allDay: allDay,
        sd: f(st, 'yyyy-MM-dd'),
        ed: f(endIncl, 'yyyy-MM-dd'),
        st: allDay ? '' : f(st, 'HH:mm'),
        et: allDay ? '' : f(en, 'HH:mm'),
        color: color,
        calId: c.getId(),
        cal: calName_(c.getName())
      });
    });
  });
  return out.sort((a, b) => (a.sd + (a.allDay ? '0' : '1') + a.st).localeCompare(b.sd + (b.allDay ? '0' : '1') + b.st));
}

/* 캘린더 목록과 표시 설정 — 앱의 '캘린더 선택'에서 켜고 끕니다(기기와 상관없이 유지) */
const DEFAULT_HIDDEN = ['한국어교육능력검정', '공인노무사', '관세사', '세무사', '변리사', '감정평가사', 'Todoist', '루틴'];
function calName_(n) { return String(n).replace(/@.*$/, ''); }
function hiddenCals_() {
  const v = PropertiesService.getUserProperties().getProperty('hiddenCals');
  if (v) return JSON.parse(v);
  const ids = CalendarApp.getAllCalendars().filter(c => DEFAULT_HIDDEN.indexOf(c.getName()) >= 0).map(c => c.getId());
  PropertiesService.getUserProperties().setProperty('hiddenCals', JSON.stringify(ids));
  return ids;
}
function calendars() {
  const hidden = hiddenCals_();
  return CalendarApp.getAllCalendars().map(c => ({ id: c.getId(), name: calName_(c.getName()), color: c.getColor(), hidden: hidden.indexOf(c.getId()) >= 0 }))
    .sort((a, b) => (a.hidden - b.hidden) || a.name.localeCompare(b.name, 'ko'));
}
function setCalHidden(id, hide) {
  const h = hiddenCals_().filter(x => x !== id);
  if (hide) h.push(id);
  PropertiesService.getUserProperties().setProperty('hiddenCals', JSON.stringify(h));
  return h;
}

/* ---------- 슈가맨워크 현황 (계약관리 통합시트 읽기 전용) ----------
 * 아침 브리핑과 같은 기준으로 계산합니다.
 *  - 비상주계약: 4행 머리글, 상태 '재실' = 유효 계약, '퇴실예정' = 정리중(주소 증빙 대기)
 *  - 상주현황: 5~29행, 3행 N~Y열 = 월별 입금칸
 *  - 호실재고: D열 '공실' = 지금 팔 수 있는 호실
 */
function sugarman(force) {
  const cache = CacheService.getScriptCache();
  if (!force) { const c = cache.get('sugar'); if (c) return JSON.parse(c); }

  const ss = SpreadsheetApp.openById(SUGAR_ID);
  const tKey = today_();
  const ym = tKey.slice(0, 7);
  const day0 = Utilities.parseDate(tKey, TZ, 'yyyy-MM-dd');
  const dKey = d => Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
  const mKey = d => Utilities.formatDate(d, TZ, 'yyyy-MM');
  const dday = d => Math.round((Utilities.parseDate(dKey(d), TZ, 'yyyy-MM-dd') - day0) / 864e5);

  // 비상주
  const nv = ss.getSheetByName('비상주계약').getDataRange().getValues();
  const H = nv[3].map(v => String(v).trim());
  const c = k => H.indexOf(k);
  const iSt = c('상태'), iEnd = c('계약종료일'), iNo = c('계약번호'), iNm = c('상호명'), iCeo = c('대표자명');
  if (iSt < 0 || iEnd < 0) throw new Error('비상주계약 탭의 머리글(4행)을 찾지 못했습니다');
  let active = 0, cleanup = 0;
  const over = [], d7 = [], d30 = [], d60 = [];
  let noExp = 0;
  nv.slice(4).forEach(r => {
    if (!r[0]) return;
    const st = String(r[iSt]).trim();
    if (st === '퇴실예정') { cleanup++; return; }
    if (st !== '재실') return;
    active++;
    const exp = r[iEnd] instanceof Date ? r[iEnd] : null;
    if (!exp) { noExp++; return; }
    const item = {
      no: String(r[iNo] || ''), room: String(r[0]),
      name: String(r[iNm] || '').trim() || String(r[iCeo] || '').trim() || '(상호미상)',
      exp: dKey(exp), dday: dday(exp)
    };
    if (item.dday < 0) over.push(item);
    else if (item.dday <= 7) d7.push(item);
    else if (item.dday <= 30) d30.push(item);
    else if (item.dday <= 60) d60.push(item);
  });
  const byD = (a, b) => a.dday - b.dday;
  over.sort(byD); d7.sort(byD);

  // 상주
  const rs = ss.getSheetByName('상주현황');
  const rv = rs.getRange(1, 1, Math.min(Math.max(rs.getLastRow(), 5), 40), 26).getValues();
  let cur = -1;
  for (let k = 13; k <= 24; k++) {
    const v = rv[2][k];
    if (v instanceof Date && mKey(v) === ym) cur = k;
  }
  let occupied = 0;
  const unpaid = [], leaving = [];
  let pending = 0;
  for (let i = 4; i < Math.min(rv.length, 29); i++) {
    const r = rv[i];
    const rm = String(r[0] || '').trim();
    if (!rm || rm === '합계') continue;
    const st = String(r[2] || '').trim();
    const name = String(r[3] || '').trim() || String(r[4] || '').trim() || '(상호없음)';
    const movein = r[6] instanceof Date ? r[6] : null;
    const moveout = r[7] instanceof Date ? r[7] : null;
    const rent = Number(r[8]) || 0;
    if (st === '사용중') occupied++;
    if (st === '퇴실예정') leaving.push({ room: rm, name: name, out: moveout ? dKey(moveout) : '', dday: moveout ? dday(moveout) : null });
    if (st !== '사용중' && st !== '퇴실예정') continue;
    if (cur < 0) continue;
    if (movein && mKey(movein) > ym) continue;
    if (moveout && mKey(moveout) < ym) continue;
    const v = String(r[cur] == null ? '' : r[cur]).trim().toLowerCase();
    let due = rent;
    if (r[11] instanceof Date && mKey(r[11]) === ym && r[12]) due = Number(r[12]) || rent;
    if (!v) unpaid.push({ room: rm, name: name, amt: due });
    else if (v === 'o') pending++;
  }
  leaving.sort((a, b) => (a.dday == null ? 99999 : a.dday) - (b.dday == null ? 99999 : b.dday));

  // 공실
  const iv = ss.getSheetByName('호실재고').getDataRange().getValues();
  const vacant = [];
  iv.slice(4).forEach(r => {
    if (!r[0]) return;
    if (String(r[3]).trim() === '공실') vacant.push({ room: String(r[0]), kind: String(r[1] || ''), type: String(r[2] || '') });
  });

  const out = {
    at: Utilities.formatDate(new Date(), TZ, 'HH:mm'), day: Number(tKey.slice(8)),
    active: active, occupied: occupied, total: 24, cleanup: cleanup, vacant: vacant,
    over: over.length, overTop: over.slice(0, 3), d7: d7, d30: d30.length, d60: d60.length, noExp: noExp,
    monthFound: cur >= 0, unpaid: unpaid, unpaidSum: unpaid.reduce((s, x) => s + (Number(x.amt) || 0), 0),
    pending: pending, leaving: leaving, url: 'https://docs.google.com/spreadsheets/d/' + SUGAR_ID + '/edit'
  };
  cache.put('sugar', JSON.stringify(out), 600);
  return out;
}
