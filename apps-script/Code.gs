/* ============================================================
   SJ SCIENCE 허브 — 구글 시트 연동 코드 (Apps Script)

   구글 시트에서 [확장 프로그램] → [Apps Script]를 열고
   이 파일 내용을 통째로 붙여 넣은 뒤 README.md 순서대로 배포하세요.

   필요한 시트 탭 (처음 한 번 setup 을 실행하면 자동으로 만들어져요)
   - 학생명단 : 학번 | 이름 | 비밀번호 | 역할 | 과목   (역할 칸에 '교사'라고 쓰면 대시보드를 볼 수 있어요.
                 과목 칸은 비워 두면 모든 과목이 보이고, '통합과학'(통합과학1·2 모두), '통합과학1', '지구과학,통합과학2'처럼 쓰면 그 과목만 보여요)
   - 진도공개 : 단원코드 | 제목 | 과목 | 주차 | (반마다 칸)   ← 칸에 체크하면 그 반에 단원이 열려요.
                 시트 메뉴 [SJ 대시보드] → [진도 공개 표 만들기·새로고침]으로 만들어요.
   - 진도     : 학번 | 이름 | 활동 | 완료 시각
   - 기록     : 학번 | 이름 | 단원 | 항목 | 내용 | 마지막 저장 | 처음 저장
                 항목 summary = 개념 정리, ox = OX 점수, oxDetail = 문항별 정오(1 맞음, 0 틀림),
                      confuse = 헷갈린 개념, fbScore·fbComment = 선생님 점수·코멘트
   - 대시보드 : 시트 메뉴 [SJ 대시보드] → [요약 탭 새로 만들기]를 누르면 만들어져요

   사이트 내용 탭 (여기를 고치면 사이트에 바로 반영돼요)
   - 단원       : 과목 | 대단원 | 단원코드 | 제목 | 목표 | 시간 | 그림 | 그림 설명 | 개념 요약 | 제시어 | 모범 정리 | 공개
   - OX         : 단원코드 | 번호 | 문항 | 정답 | 해설
   - 뉴스       : 단원코드 | 제목 | 출처 | 날짜 | 링크 | 요약 | 질문 | 설명
   - 헷갈린개념 : 단원코드 | 보기
   ※ OX 정답과 모범 정리는 사이트 코드에 들어가지 않고, 서버에서만 확인해요.
   ============================================================ */

const SHEET_STUDENTS = '학생명단';
const SHEET_PROGRESS = '진도';
const SHEET_NOTES = '기록';
const SHEET_SUMMARY = '대시보드';
const TEACHER = '교사';
const MAX_NOTE = 2000;          // 글 하나의 최대 글자 수
const TOKEN_HOURS = 6;          // 로그인 유지 시간 (최대 6시간, 지나면 다시 로그인)
const MAX_FAILS = 5;            // 비밀번호를 이 횟수만큼 틀리면
const LOCK_MINUTES = 10;        // 이 시간 동안 로그인을 막아요
const LOCKED_FIELDS = ['ox', 'oxDetail', 'blankFirst'];   // 처음 기록한 뒤에는 바뀌지 않는 항목
const T_UNITS = '단원', T_OX = 'OX', T_NEWS = '뉴스', T_CONFUSE = '헷갈린개념', T_BLANK = '빈칸', T_ESSAY = '서술형';
const CONTENT_TABS = [T_UNITS, T_OX, T_NEWS, T_CONFUSE, T_BLANK, T_ESSAY];
const CONTENT_CACHE = 'content_v1';
const SHEET_OPEN = '진도공개';

/* ---------- 처음 한 번 실행: 시트 탭 만들기 ---------- */
function setup() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_STUDENTS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_STUDENTS);
    sh.appendRow(['학번', '이름', '비밀번호', '역할', '과목', '반']);
    sh.appendRow(['teacher', '이상준', '바꿔주세요', TEACHER]);
    sh.appendRow(['10101', '홍길동', '1234', '']);
    sh.getRange('A:A').setNumberFormat('@');   // 학번을 글자로 (앞자리 0 유지)
    sh.getRange('C:C').setNumberFormat('@');
    sh.setFrozenRows(1);
  } else if (sh.getRange(1, 4).getValue() === '') {
    sh.getRange(1, 4).setValue('역할');          // 기존 명단에는 '역할' 칸만 추가
  }
  if (sh.getRange(1, 5).getValue() === '') sh.getRange(1, 5).setValue('과목');   // 기존 명단에는 '과목' 칸 추가
  if (sh.getRange(1, 6).getValue() === '') sh.getRange(1, 6).setValue('반');     // 기존 명단에는 '반' 칸 추가 (비워 두면 학번에서 자동)
  if (!ss.getSheetByName(SHEET_PROGRESS)) {
    const p = ss.insertSheet(SHEET_PROGRESS);
    p.appendRow(['학번', '이름', '활동', '완료 시각']);
    p.getRange('A:A').setNumberFormat('@');
    p.setFrozenRows(1);
  }
  let n = ss.getSheetByName(SHEET_NOTES);
  if (!n) {
    n = ss.insertSheet(SHEET_NOTES);
    n.appendRow(['학번', '이름', '단원', '항목', '내용', '마지막 저장', '처음 저장']);
    n.getRange('A:A').setNumberFormat('@');
    n.setColumnWidth(5, 420);
    n.setFrozenRows(1);
  } else if (n.getRange(1, 7).getValue() === '') {
    n.getRange(1, 6, 1, 2).setValues([['마지막 저장', '처음 저장']]);
  }
  setupContent();
}

/* 사이트 내용 탭 만들기 + 예시 단원(대기 대순환과 표층 해류) 채우기 */
function setupContent() {
  const ss = SpreadsheetApp.getActive();
  const make = (name, head, rows, widths) => {
    if (ss.getSheetByName(name)) return;
    const sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#e8f5f3');
    if (rows.length) sh.getRange(2, 1, rows.length, head.length).setNumberFormat('@').setValues(rows);
    sh.getRange('A:Z').setWrap(true).setVerticalAlignment('top');
    (widths || []).forEach((w, i) => w && sh.setColumnWidth(i + 1, w));
    sh.setFrozenRows(1);
  };
  make(T_UNITS,
    ['과목', '대단원', '단원코드', '제목', '목표', '시간', '그림', '그림 설명', '개념 요약', '제시어', '모범 정리', '공개'],
    [['지구과학', '1. 대기와 해양의 상호작용', '1-2-1', '대기 대순환과 표층 해류',
      '대기 대순환으로 바람의 방향을 설명하고, 바람이 표층 해류를 만드는 과정을 정리한다.', '5분',
      'earth-1-2-1/winds-currents-map.jpg', '대기 대순환과 전 세계의 표층 해류 분포 (빨간색: 난류, 파란색: 한류)',
      ['저위도는 에너지가 남고 고위도는 모자라요. 대기와 해수가 남는 열을 고위도로 옮기는 흐름이 **대기 대순환**이에요.',
       '지구가 자전하기 때문에 위도에 따라 **무역풍(0°~30°)**, **편서풍(30°~60°)**, **극동풍(60°~90°)**이 불어요.',
       '바람이 바다 표면을 끌고 가서 **표층 해류**가 생겨요. 무역풍대에서는 동→서, 편서풍대에서는 서→동으로 흘러요.',
       '대양의 큰 순환은 북반구에서 **시계 방향**, 남반구에서 **시계 반대 방향**이에요. 대양 서쪽에는 난류, 동쪽에는 한류가 흘러요.'].join('\n'),
      '위도, 무역풍, 편서풍, 표층 해류',
      '저위도는 태양 에너지가 남고 고위도는 부족해서, 대기가 열을 고위도로 옮기는 대기 대순환이 일어난다. 지구가 자전하기 때문에 위도 0°~30°에는 무역풍, 30°~60°에는 편서풍이 분다. 이 바람이 바다 표면을 끌고 가서 표층 해류가 생기고, 무역풍대에서는 동→서, 편서풍대에서는 서→동으로 흐른다.',
      'O']],
    [80, 160, 80, 160, 220, 50, 200, 200, 420, 160, 360, 50]);
  make(T_OX, ['단원코드', '번호', '문항', '정답', '해설'], [
    ['1-2-1', '1', '위도 0°~30° 지역의 지표 부근에서는 편서풍이 분다.', 'X', '위도 0°~30°에는 무역풍이 불어요. 편서풍은 위도 30°~60°에서 불어요.'],
    ['1-2-1', '2', '위도 30° 부근은 공기가 내려오는 곳이라 건조한 지역이 많다.', 'O', '공기가 내려오면 구름이 잘 생기지 않아 맑고 건조해요.'],
    ['1-2-1', '3', '북태평양의 표층 해류는 시계 방향으로 순환한다.', 'O', '북반구 대양의 큰 순환은 시계 방향이에요.'],
    ['1-2-1', '4', '쿠로시오 해류는 저위도에서 고위도로 흐르는 한류이다.', 'X', '쿠로시오 해류는 따뜻한 물을 고위도로 옮기는 난류예요.'],
    ['1-2-1', '5', '남반구 대양의 큰 해류 순환은 시계 방향으로 돈다.', 'X', '남반구는 북반구와 반대로 시계 반대 방향으로 돌아요.'],
  ], [80, 50, 420, 50, 360]);
  make(T_NEWS, ['단원코드', '제목', '출처', '날짜', '링크', '요약', '질문', '설명'], [
    ['1-2-1', '태평양 한가운데 떠 있는 거대한 쓰레기 섬', 'YTN', '2022.11.08', 'https://www.ytn.co.kr/_ln/0105_202211080428284662',
     '미국 캘리포니아와 하와이 사이 바다에 플라스틱 쓰레기가 모여 있는 곳을 "태평양 거대 쓰레기 섬"이라고 불러요. 이곳 쓰레기의 80% 이상은 그물, 로프, 부표 같은 어업 폐기물이에요.',
     '쓰레기가 하필 북태평양 한가운데(북위 30° 부근)에 모이는 이유는 무엇일까요?',
     '이곳은 북태평양 해류, 캘리포니아 해류, 북적도 해류, 쿠로시오 해류가 둘러싸고 시계 방향으로 도는 순환의 한가운데예요. 둘레를 도는 해류가 떠다니는 물체를 안쪽에 가두기 때문에 쓰레기가 쌓여요.'],
  ], [80, 220, 80, 90, 200, 360, 260, 360]);
  make(T_CONFUSE, ['단원코드', '보기'], [
    ['1-2-1', '바람 띠의 위도와 이름'], ['1-2-1', '무역풍·편서풍의 방향'], ['1-2-1', '해류가 도는 방향 (북반구·남반구)'],
    ['1-2-1', '난류·한류가 흐르는 위치'], ['1-2-1', '헷갈린 것 없음'],
  ], [80, 300]);
}

/* ---------- 시트 메뉴 ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('SJ 대시보드')
    .addItem('요약 탭 새로 만들기', 'buildSummary')
    .addItem('진도 공개 표 만들기·새로고침', 'buildOpenSheet')
    .addItem('주차까지 열기 (반 선택)', 'openByWeek')
    .addItem('사이트 내용 바로 반영', 'clearContentCache')
    .addItem('필터·드롭다운 적용', 'applySheetTools')
    .addItem('내용 탭 만들기 (처음 한 번)', 'setupContent')
    .addToUi();
}

/* 내용 탭을 고치면 사이트에 반영되도록 저장해 둔 내용을 지워요 */
function onEdit(e) {
  if (e && e.range && isContentTab(e.range.getSheet().getName())) clearContentCache();
}
function clearContentCache() { CacheService.getScriptCache().remove(CONTENT_CACHE); }

/* 표 한 장을 [{제목: 값}] 으로 읽기 */
function isContentTab(n) { return CONTENT_TABS.some(b => n === b || n.endsWith(' ' + b)); }
function readTable(name) {   // '단원' 이라고 부르면 '통합과학1 단원', '지구과학 단원' 같은 탭도 함께 읽어요
  const out = [];
  SpreadsheetApp.getActive().getSheets().forEach(sh => {
    const n = sh.getName();
    if (n !== name && !n.endsWith(' ' + name)) return;
    const rows = sh.getDataRange().getDisplayValues();
    const head = rows[0].map(h => String(h).trim());
    rows.slice(1).filter(r => r.some(c => String(c).trim() !== ''))
      .forEach(r => out.push(Object.fromEntries(head.map((h, i) => [h, String(r[i] ?? '').trim()]))));
  });
  return out;
}
const normOX = v => /^(o|○|ㅇ|true|참|맞음)$/i.test(String(v).trim()) ? 'O' : 'X';

/* 사이트에 보낼 내용 — 정답·모범 정리는 빼고 보내요 */
function getContent() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(CONTENT_CACHE);
  if (hit) return JSON.parse(hit);
  const ox = readTable(T_OX), news = readTable(T_NEWS), conf = readTable(T_CONFUSE), blank = readTable(T_BLANK), essay = readTable(T_ESSAY);
  const byNum = (a, b) => Number(a['번호']) - Number(b['번호']);
  const units = readTable(T_UNITS).filter(u => u['단원코드'] && u['공개'] !== 'X').map(u => {
    const code = u['단원코드'];
    return {
      subject: u['과목'], chapter: u['대단원'], code, title: u['제목'], goal: u['목표'], time: u['시간'],
      img: u['그림'], imgCaption: u['그림 설명'],
      points: (u['개념 요약'] || '').split('\n').map(x => x.trim()).filter(Boolean),
      hint: u['제시어'], hasModel: !!u['모범 정리'],
      ox: ox.filter(o => o['단원코드'] === code && o['문항']).sort((a, b) => Number(a['번호']) - Number(b['번호'])).map(o => ({ q: o['문항'] })),
      news: news.filter(n => n['단원코드'] === code && n['제목']).map(n => ({
        title: n['제목'], source: n['출처'], date: n['날짜'], url: n['링크'], summary: n['요약'], ask: n['질문'], answer: n['설명'] })),
      confuse: conf.filter(c => c['단원코드'] === code && c['보기']).map(c => c['보기']),
      blanks: blank.filter(b => b['단원코드'] === code && b['문장']).sort(byNum).map(b => ({ q: String(b['문장']).replace(/\{[^}]*\}/g, '{{}}'), sec: b['소제목'] || '' })),   // 정답은 빼고 보내요
      essays: essay.filter(e => e['단원코드'] === code && e['질문']).sort(byNum).map(e => ({ n: e['번호'], q: e['질문'], hint: e['제시어'], outline: e['힌트'], sec: e['소제목'] || '' })),
    };
  });
  const data = { ok: true, units, at: new Date().toISOString() };
  try { cache.put(CONTENT_CACHE, JSON.stringify(data), 600); } catch (err) {}   // 10분 동안 저장 (빠르게)
  return data;
}

function findUnitRow(code) { return readTable(T_UNITS).find(u => u['단원코드'] === code); }

/* 빈칸 채우기 채점 — 문장 안의 {정답} 이 정답이에요. {산화/산화 반응} 처럼 / 로 여러 정답을 쓸 수 있어요 */
const normAns = s => String(s || '').replace(/\([^)]*\)/g, '').replace(/[\s·.,]/g, '').toLowerCase();
function blankKeys(text) { return (String(text).match(/\{[^}]*\}/g) || []).map(m => m.slice(1, -1).split('/').map(x => x.trim()).filter(Boolean)); }
function checkBlank(keys, ans) { const a = normAns(ans); return !!a && keys.some(k => normAns(k) === a); }
function gradeBlank(u, req) {
  const code = String(req.code || ''), unit = String(req.unit || '').slice(0, 100);
  if (code && !isOpenFor(code, u, readOpen())) return { ok: false, error: '아직 열리지 않은 단원이에요.' };
  const rows = readTable(T_BLANK).filter(b => b['단원코드'] === code && b['문장']).sort((a, b) => Number(a['번호']) - Number(b['번호']));
  const keys = []; rows.forEach(r => blankKeys(r['문장']).forEach(k => keys.push(k)));
  const answers = Array.isArray(req.answers) ? req.answers : [];
  if (!keys.length || answers.length !== keys.length) return { ok: false, error: '빈칸 수가 맞지 않아요. 새로고침해 주세요.' };
  const results = keys.map((k, i) => ({ ok: checkBlank(k, answers[i]), a: k[0] }));
  const n = results.filter(r => r.ok).length, score = `${n}/${results.length}`;
  // 점수는 풀 때마다 새로 기록하고, 첫 번째 점수는 따로 남겨 둬요
  writeNote(u.id, u.name, unit, 'blankDetail', results.map(r => r.ok ? '1' : '0').join(''));
  writeNote(u.id, u.name, unit, 'blank', score);
  const w = writeNote(u.id, u.name, unit, 'blankFirst', score);
  return { ok: true, results, score, recorded: true, first: w.kept || score };
}

/* OX 채점 — 정답은 서버에만 있어요. 처음 제출한 결과만 기록 */
function gradeOx(u, req) {
  const code = String(req.code || ''), unit = String(req.unit || '').slice(0, 100);
  if (code && !isOpenFor(code, u, readOpen())) return { ok: false, error: '아직 열리지 않은 단원이에요.' };
  const key = readTable(T_OX).filter(o => o['단원코드'] === code && o['문항']).sort((a, b) => Number(a['번호']) - Number(b['번호']));
  const picks = Array.isArray(req.picks) ? req.picks : [];
  if (!key.length || picks.length !== key.length) return { ok: false, error: '모든 문항에 답해 주세요.' };
  const results = key.map((o, i) => { const a = normOX(o['정답']); return { ok: normOX(picks[i]) === a, a, exp: o['해설'] }; });
  const n = results.filter(r => r.ok).length, score = `${n}/${results.length}`;
  const detail = results.map(r => r.ok ? '1' : '0').join('');
  const before = getNotes(u.id, unit);
  if (before.ox) return { ok: true, results, score, recorded: false, kept: before.ox };
  writeNote(u.id, u.name, unit, 'oxDetail', detail);
  writeNote(u.id, u.name, unit, 'ox', score);
  return { ok: true, results, score, recorded: true };
}

/* 학생 × 단원 표: OX 점수와 개념 정리 제출 여부 */
function buildSummary() {
  const ss = SpreadsheetApp.getActive();
  const students = readStudents().filter(s => s.role !== TEACHER);
  const notes = readNotes();
  const units = [...new Set(notes.map(n => n.unit))].sort();
  const head = ['학번', '이름'];
  units.forEach(u => head.push(u + ' OX', u + ' 개념 정리', u + ' 점수', u + ' 헷갈린 개념'));
  const rows = students.map(s => {
    const row = [s.id, s.name];
    units.forEach(u => {
      const mine = notes.filter(n => n.id === s.id && n.unit === u);
      const get = f => (mine.find(n => n.field === f) || {}).text || '';
      row.push(get('ox'), get('summary'), get('fbScore'), get('confuse'));
    });
    return row;
  });
  let sh = ss.getSheetByName(SHEET_SUMMARY);
  if (sh) sh.clear(); else sh = ss.insertSheet(SHEET_SUMMARY);
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#e8f5f3');
  if (rows.length) sh.getRange(2, 1, rows.length, head.length).setNumberFormat('@').setValues(rows).setWrap(true).setVerticalAlignment('top');
  sh.setFrozenRows(1); sh.setFrozenColumns(2);
  units.forEach((u, i) => sh.setColumnWidth(4 + i * 4, 360));
  sh.getRange(rows.length + 3, 1).setValue('만든 시각: ' + new Date().toLocaleString('ko-KR'));
}

/* ---------- 사이트에서 오는 요청 받기 ---------- */
function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out({ ok: false, error: '잘못된 요청이에요.' }); }
  try {
    switch (req.action) {
      case 'login':    return out(login(req));
      case 'content':  return out(contentFor(userOf(req.token)));
      case 'gradeOx':  return out(withUser(req, u => gradeOx(u, req)));
      case 'gradeBlank': return out(withUser(req, u => gradeBlank(u, req)));
      case 'progress': return out(withUser(req, u => ({ ok: true, done: getDone(u.id) })));
      case 'mark':     return out(withUser(req, u => mark(u, req)));
      case 'notes':    return out(withUser(req, u => {
        const notes = getNotes(u.id, String(req.unit || ''));
        if (notes.summary && req.code) { const r = findUnitRow(String(req.code)); if (r) notes._model = r['모범 정리']; }
        if (req.code) readTable(T_ESSAY).filter(e => e['단원코드'] === String(req.code) && notes['essay' + e['번호']]).forEach(e => { notes['_m_essay' + e['번호']] = e['모범 답안']; });   // 서술형은 저장한 뒤에 모범 답안 보여 주기
        return { ok: true, notes };
      }));
      case 'myNotes':  return out(withUser(req, u => ({ ok: true, notes: getAllNotes(u.id) })));
      case 'saveNote': return out(withUser(req, u => saveNote(u, req)));
      case 'changePw': return out(withUser(req, u => changePw(u, req)));
      case 'dash':     return out(withTeacher(req, () => dash()));
      case 'openGet':  return out(withTeacher(req, () => openGet()));
      case 'openSet':  return out(withTeacher(req, () => openSet(req)));
      case 'openWeek': return out(withTeacher(req, () => openWeek(req)));
      case 'feedback': return out(withTeacher(req, () => feedback(req)));
      case 'logout':   CacheService.getScriptCache().remove('tok_' + req.token); return out({ ok: true });
      default:         return out({ ok: false, error: '알 수 없는 요청이에요.' });
    }
  } catch (err) {
    return out({ ok: false, error: '서버 오류: ' + err.message });
  }
}
function doGet() { return out({ ok: true, message: 'SJ SCIENCE API 동작 중' }); }

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 로그인 ---------- */
function login(req) {
  const id = String(req.id || '').trim(), pw = String(req.pw || '').trim();
  if (!id || !pw) return { ok: false, error: '학번과 비밀번호를 모두 입력해 주세요.' };

  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('fail_' + id) || 0);
  if (fails >= MAX_FAILS) return { ok: false, error: `비밀번호를 여러 번 틀렸어요. ${LOCK_MINUTES}분 뒤에 다시 해 주세요.` };

  const st = findStudent(id);
  if (!st || st.pw !== pw) {
    cache.put('fail_' + id, String(fails + 1), LOCK_MINUTES * 60);
    return { ok: false, error: '학번 또는 비밀번호가 맞지 않아요.' };
  }
  cache.remove('fail_' + id);

  // 출입증(토큰) 발급 — 사이트는 비밀번호 대신 이 값을 들고 다녀요
  const token = Utilities.getUuid();
  cache.put('tok_' + token, JSON.stringify({ id: st.id, name: st.name, role: st.role, subj: st.subj, cls: st.cls || classOf(st.id) }), TOKEN_HOURS * 3600);
  return { ok: true, token, id: st.id, name: st.name, role: st.role, subj: st.subj, cls: st.cls || classOf(st.id), done: getDone(st.id) };
}

/* 출입증 확인 후 작업 실행 */
function withUser(req, fn) {
  const raw = req.token && CacheService.getScriptCache().get('tok_' + req.token);
  if (!raw) return { ok: false, error: 'EXPIRED' };
  return fn(JSON.parse(raw));
}
function withTeacher(req, fn) {
  return withUser(req, u => u.role === TEACHER ? fn(u) : { ok: false, error: '선생님만 볼 수 있어요.' });
}

function readStudents() {
  const rows = SpreadsheetApp.getActive().getSheetByName(SHEET_STUDENTS).getDataRange().getDisplayValues();
  return rows.slice(1).filter(r => String(r[0]).trim()).map((r, i) => ({
    row: i + 2, id: String(r[0]).trim(), name: String(r[1]).trim(), pw: String(r[2]).trim(), role: String(r[3] || '').trim(),
    subj: String(r[4] || '').trim(), cls: String(r[5] || '').trim(),   // 반 칸이 있으면 학번 대신 그 이름으로 반을 인식해요 (예: 방과후A)
  }));
}
function findStudent(id) { return readStudents().find(s => s.id === id) || null; }

/* ---------- 진도 기록 ---------- */
function getDone(id) {
  const rows = SpreadsheetApp.getActive().getSheetByName(SHEET_PROGRESS).getDataRange().getDisplayValues();
  return rows.slice(1).filter(r => String(r[0]).trim() === id).map(r => r[2]);
}

function mark(u, req) {
  const key = String(req.key || '').slice(0, 300);
  if (!key) return { ok: false, error: '활동 이름이 없어요.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);   // 여러 학생이 동시에 눌러도 꼬이지 않게
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_PROGRESS);
    const rows = sh.getDataRange().getDisplayValues();
    const found = [];
    for (let r = 1; r < rows.length; r++) if (rows[r][0] === u.id && rows[r][2] === key) found.push(r + 1);
    if (req.on && !found.length) {
      sh.appendRow([u.id, u.name, key, new Date()]);
    } else if (!req.on) {
      found.reverse().forEach(row => sh.deleteRow(row));
    }
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* ---------- 학생이 쓴 글·점수 ---------- */
function readNotes() {
  const rows = SpreadsheetApp.getActive().getSheetByName(SHEET_NOTES).getDataRange().getDisplayValues();
  return rows.slice(1).map(r => ({ id: r[0], name: r[1], unit: r[2], field: r[3], text: r[4], last: r[5], first: r[6] || r[5] }));
}

function getNotes(id, unit) {
  const notes = {};
  readNotes().forEach(n => { if (n.id === id && n.unit === unit) notes[n.field] = n.text; });
  return notes;
}

function getAllNotes(id) {   // 내 기록(포트폴리오): { 단원: { 항목: 내용 } }
  const all = {};
  readNotes().forEach(n => { if (n.id === id) (all[n.unit] = all[n.unit] || {})[n.field] = n.text; });
  return all;
}

function saveNote(u, req) {
  const unit = String(req.unit || '').slice(0, 100), field = String(req.field || '').slice(0, 50);
  let text = String(req.text || '').slice(0, MAX_NOTE);
  if (/^[=+\-@]/.test(text)) text = "'" + text;   // 수식으로 바뀌지 않게
  if (!unit || !field) return { ok: false, error: '저장할 위치가 없어요.' };
  if (req.code && !isOpenFor(String(req.code), u, readOpen())) return { ok: false, error: '아직 열리지 않은 단원이에요.' };
  if (field.startsWith('fb') || field.startsWith('blank') || LOCKED_FIELDS.includes(field)) return { ok: false, error: '직접 저장할 수 없는 항목이에요.' };
  const res = writeNote(u.id, u.name, unit, field, text);
  if (field === 'summary' && req.code) { const r = findUnitRow(String(req.code)); if (r) res.model = r['모범 정리']; }   // 저장한 뒤에 모범 정리 보여 주기
  const em = /^essay(\d+)$/.exec(field);
  if (em && req.code) { const r = readTable(T_ESSAY).find(e => e['단원코드'] === String(req.code) && String(e['번호']) === em[1]); if (r) res.model = r['모범 답안']; }
  return res;
}

/* 기록 탭에 한 줄 쓰기 (이미 있으면 고쳐 쓰기) */
function writeNote(id, name, unit, field, text) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NOTES);
    const rows = sh.getDataRange().getDisplayValues();
    const now = new Date();
    for (let r = 1; r < rows.length; r++) {
      if (rows[r][0] === id && rows[r][2] === unit && rows[r][3] === field) {
        if (LOCKED_FIELDS.includes(field)) return { ok: true, kept: rows[r][4] };   // OX 기록은 처음 것 유지
        sh.getRange(r + 1, 5, 1, 2).setValues([[text, now]]);                       // 고쳐 쓰기: 처음 저장 시각은 그대로
        return { ok: true };
      }
    }
    sh.appendRow([id, name, unit, field, text, now, now]);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* ---------- 교사 대시보드 자료 ---------- */
function dash() {
  const students = readStudents().filter(s => s.role !== TEACHER).map(s => ({ id: s.id, name: s.name, cls: s.cls || classOf(s.id) }));
  const oxKey = {};   // 문항별 정답 (선생님 화면에만)
  readTable(T_OX).sort((a, b) => Number(a['번호']) - Number(b['번호'])).forEach(o => (oxKey[o['단원코드']] = oxKey[o['단원코드']] || []).push(normOX(o['정답'])));
  return { ok: true, students, notes: readNotes(), oxKey };
}

/* 선생님 채점·코멘트 → 학생의 '내 기록'에 보여요 */
function feedback(req) {
  const st = findStudent(String(req.sid || '').trim());
  const unit = String(req.unit || '').slice(0, 100);
  if (!st || !unit) return { ok: false, error: '학생 또는 단원을 찾을 수 없어요.' };
  const score = String(req.score ?? '').trim();
  if (score !== '' && !/^[0-5]$/.test(score)) return { ok: false, error: '점수는 0~5 사이로 입력해 주세요.' };
  let comment = String(req.comment || '').slice(0, 300);
  if (/^[=+\-@]/.test(comment)) comment = "'" + comment;
  writeNote(st.id, st.name, unit, 'fbScore', score);
  writeNote(st.id, st.name, unit, 'fbComment', comment);
  return { ok: true };
}

/* ---------- 비밀번호 바꾸기 ---------- */
function changePw(u, req) {
  const oldPw = String(req.oldPw || '').trim(), newPw = String(req.newPw || '').trim();
  if (newPw.length < 4) return { ok: false, error: '새 비밀번호는 4자 이상이어야 해요.' };
  const st = findStudent(u.id);
  if (!st || st.pw !== oldPw) return { ok: false, error: '지금 비밀번호가 맞지 않아요.' };
  SpreadsheetApp.getActive().getSheetByName(SHEET_STUDENTS).getRange(st.row, 3).setNumberFormat('@').setValue(newPw);
  return { ok: true };
}


/* ============================================================
   반별 진도 공개 + 과목별 학생 구분
   ============================================================ */
/* 학번(5자리)에서 반 이름 만들기 — 사이트 대시보드와 같은 규칙: 10305 → '1학년 3반' */
function classOf(id) { return /^\d{5}$/.test(id) ? `${id[0]}학년 ${Number(id.slice(1, 3))}반` : '기타'; }

function userOf(token) {
  const raw = token && CacheService.getScriptCache().get('tok_' + token);
  return raw ? JSON.parse(raw) : null;
}

/* '진도공개' 탭 읽기 → { classes, map: {단원코드: {반: true/false}}, week: {단원코드: 주차} } */
function openSheets() {   // '통합과학1 진도공개', '통합과학2 진도공개', '지구과학 진도공개' … (옛 '진도공개' 한 장도 읽어요)
  return SpreadsheetApp.getActive().getSheets().filter(sh => { const n = sh.getName(); return n === SHEET_OPEN || n.endsWith(' ' + SHEET_OPEN); });
}
function readOpen() {
  const res = { classes: [], map: {}, week: {} };
  openSheets().forEach(sh => {
    if (sh.getLastRow() < 1) return;
    const rows = sh.getDataRange().getDisplayValues();
    const cl = rows[0].slice(4).map(x => String(x).trim()).filter(Boolean);
    cl.forEach(c => { if (!res.classes.includes(c)) res.classes.push(c); });
    rows.slice(1).forEach(r => {
      const code = String(r[0]).trim(); if (!code) return;
      res.week[code] = String(r[3]).trim();
      res.map[code] = {};
      cl.forEach((c, i) => { res.map[code][c] = String(r[4 + i]).trim().toUpperCase() === 'TRUE'; });
    });
  });
  res.classes.sort();
  return res;
}

/* 이 사용자에게 단원이 열려 있나요? 표에 없는 단원은 모두에게 열려 있어요 */
function isOpenFor(code, user, open) {
  const m = open.map[code];
  if (!m) return true;
  if (user && user.role === TEACHER) return true;
  if (!user) return Object.keys(m).some(c => m[c]);
  return m[user.cls || classOf(user.id)] === true;
}

/* 사이트에 보낼 내용: 학생의 과목만 남기고, 안 열린 단원은 제목만 보내요 */
function contentFor(user) {
  const data = getContent();
  if (!user) return { ok: true, units: [], at: data.at };   // 로그인 전에는 내용을 보내지 않아요
  const open = readOpen();
  const subj = user && user.role !== TEACHER && user.subj ? user.subj.split(/[,\s]+/).filter(Boolean) : null;
  const units = data.units.filter(u => !subj || subj.some(x => u.subject === x || u.subject.indexOf(x) === 0)).map(u => {
    if (isOpenFor(u.code, user, open)) return u;
    return { subject: u.subject, chapter: u.chapter, code: u.code, title: u.title, goal: u.goal, time: u.time, locked: true, week: open.week[u.code] || '' };
  });
  return { ok: true, units, at: data.at };
}

/* 시트 메뉴: 진도 공개 표 만들기·새로고침 (이미 체크한 칸은 그대로 두고 새 단원·새 반만 추가) */
function buildOpenSheet() {
  const ss = SpreadsheetApp.getActive();
  const old = readOpen();
  const classes = [...new Set([...old.classes, ...readStudents().filter(s => s.role !== TEACHER).map(s => s.cls || classOf(s.id))])].sort();
  const units = readTable(T_UNITS).filter(u => u['단원코드'] && u['공개'] !== 'X')
    .sort((a, b) => a['단원코드'].localeCompare(b['단원코드'], 'ko', { numeric: true }));
  const bySubj = {};
  units.forEach(u => { const s = u['과목'] || '기타'; (bySubj[s] = bySubj[s] || []).push(u); });
  const legacy = ss.getSheetByName(SHEET_OPEN);   // 옛 한 장짜리 표는 읽은 뒤 과목별 탭으로 바꿔요
  const head = ['단원코드', '제목', '과목', '주차', ...classes];
  const made = [];
  Object.keys(bySubj).sort().forEach(subj => {
    const rows = bySubj[subj].map(u => {
      const code = u['단원코드'];
      return [code, u['제목'], subj, old.week[code] || '', ...classes.map(c => !!(old.map[code] && old.map[code][c]))];
    });
    const name = subj + ' ' + SHEET_OPEN;
    let sh = ss.getSheetByName(name);
    if (sh) { const f = sh.getFilter(); if (f) f.remove(); sh.clear(); sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations(); } else sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#e8f5f3');
    sh.getRange(2, 1, rows.length, 3).setNumberFormat('@');
    sh.getRange(2, 1, rows.length, head.length).setValues(rows);
    if (classes.length) sh.getRange(2, 5, rows.length, classes.length).insertCheckboxes();
    sh.setFrozenRows(1); sh.setFrozenColumns(2);
    sh.setColumnWidth(2, 220);
    made.push(sh);
  });
  if (legacy && !made.some(s => s.getName() === SHEET_OPEN)) ss.deleteSheet(legacy);
  applySheetTools();
  ss.toast('과목별 [진도공개] 탭에서 체크한 반에 그 단원이 열려요. 주차 칸에 숫자를 쓰면 [주차까지 열기]를 쓸 수 있어요.');
}

/* 시트 메뉴: N주차까지 열기 (주차 칸에 숫자가 있는 단원만 — N주차 이하는 열고, 더 뒤 주차는 닫아요) */
function openByWeek() {
  const ui = SpreadsheetApp.getUi();
  const cl = readOpen().classes;
  if (!cl.length) { ui.alert('먼저 [진도 공개 표 만들기·새로고침]을 눌러 주세요.'); return; }
  const a = ui.prompt('어느 반?', '반 이름을 쓰세요 (비우면 모든 반).\n' + cl.join(', '), ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK) return;
  const b = ui.prompt('몇 주차까지 열까요?', '숫자만 쓰세요. 주차 칸이 비어 있는 단원은 그대로 둬요.', ui.ButtonSet.OK_CANCEL);
  if (b.getSelectedButton() !== ui.Button.OK) return;
  const r = openWeek({ cls: a.getResponseText().trim(), week: b.getResponseText().trim() });
  ui.alert(r.ok ? `열었어요. (${r.changed}칸 바뀜)` : r.error);
}

function openSetCells(fn) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const shs = openSheets();
    if (!shs.length) return { ok: false, error: '먼저 시트 메뉴에서 [진도 공개 표 만들기·새로고침]을 눌러 주세요.' };
    let changed = 0;
    shs.forEach(sh => { changed += fn(sh.getDataRange().getDisplayValues(), sh) || 0; });
    return { ok: true, changed };
  } finally { lock.releaseLock(); }
}

function openGet() {
  if (!openSheets().length) buildOpenSheet();
  const o = readOpen();
  const units = readTable(T_UNITS).filter(u => u['단원코드'] && u['공개'] !== 'X').map(u => ({ code: u['단원코드'], title: u['제목'], subject: u['과목'], week: o.week[u['단원코드']] || '' }));
  return { ok: true, classes: o.classes, units, map: o.map };
}

function openSet(req) {   // 대시보드의 스위치 한 개
  const code = String(req.code || ''), cls = String(req.cls || ''), on = !!req.on;
  return openSetCells((vals, sh) => {
    const ci = vals[0].indexOf(cls), ri = vals.findIndex((r, i) => i > 0 && String(r[0]).trim() === code);
    if (ci < 4 || ri < 1) return 0;
    sh.getRange(ri + 1, ci + 1).setValue(on); return 1;
  });
}

function openWeek(req) {
  const w = Number(req.week), cls = String(req.cls || '').trim();
  if (!isFinite(w) || String(req.week).trim() === '') return { ok: false, error: '주차는 숫자로 써 주세요.' };
  return openSetCells((vals, sh) => {
    let changed = 0;
    for (let r = 1; r < vals.length; r++) {
      const wk = Number(String(vals[r][3]).trim()); if (String(vals[r][3]).trim() === '' || !isFinite(wk)) continue;
      for (let c = 4; c < vals[0].length; c++) {
        if (cls && vals[0][c] !== cls) continue;
        const want = wk <= w, now = String(vals[r][c]).trim().toUpperCase() === 'TRUE';
        if (want !== now) { sh.getRange(r + 1, c + 1).setValue(want); changed++; }
      }
    }
    return changed;
  });
}

/* ---------- 시트 편의: 필터·드롭다운 ---------- */
// 메뉴 [필터·드롭다운 적용]을 누르면 모든 탭에 필터와 선택 목록이 붙어요.
function applySheetTools() {
  const ss = SpreadsheetApp.getActive();
  const list = (values, strict) => SpreadsheetApp.newDataValidation().requireValueInList(values, true).setAllowInvalid(!strict).build();
  const colOf = (sh, name) => sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getDisplayValues()[0].map(h => String(h).trim()).indexOf(name) + 1;
  const dropdown = (sh, name, values, strict) => {
    const c = colOf(sh, name);
    if (c > 0) sh.getRange(2, c, Math.max(sh.getMaxRows() - 1, 1), 1).setDataValidation(list(values, strict));
  };
  const SUBJ = ['통합과학1', '통합과학2', '지구과학'];
  ss.getSheets().forEach(sh => {
    const n = sh.getName();
    if (sh.getLastRow() < 1 || sh.getLastColumn() < 1) return;
    const base = CONTENT_TABS.find(b => n === b || n.endsWith(' ' + b)) || (n === SHEET_OPEN || n.endsWith(' ' + SHEET_OPEN) ? SHEET_OPEN : n);
    if (base === '학생명단') {
      dropdown(sh, '역할', ['학생', '교사'], false);
      dropdown(sh, '과목', ['통합과학', '통합과학1', '통합과학2', '지구과학'], false);   // 두 과목은 직접 '통합과학1,지구과학'처럼 써도 돼요
    } else if (base === T_UNITS) {
      dropdown(sh, '과목', SUBJ, false);
      dropdown(sh, '공개', ['O', 'X'], false);
    } else if (base === T_OX) {
      dropdown(sh, '정답', ['O', 'X'], true);
    } else if (base === SHEET_OPEN) {
      dropdown(sh, '주차', Array.from({ length: 40 }, (_, i) => String(i + 1)), false);
      dropdown(sh, '과목', SUBJ, false);
    } else if (!(base === T_NEWS || base === T_CONFUSE || n === SHEET_PROGRESS || n === SHEET_NOTES)) return;
    const old = sh.getFilter(); if (old) old.remove();
    sh.getDataRange().createFilter();
    sh.setFrozenRows(1);
  });
  ss.toast('필터(머리글 옆 ▼)와 선택 목록을 붙였어요.');
}
