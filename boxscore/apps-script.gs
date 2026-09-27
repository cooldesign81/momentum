/**
 * Momentum 박스스코어 → 리그 기록 시트 연동용 Apps Script
 *
 * 설치 (한 번만):
 * 1. 리그 기록 구글 시트를 열고 [확장 프로그램] > [Apps Script]
 * 2. 이 파일 내용을 통째로 붙여넣고 저장
 * 3. 왼쪽 톱니바퀴 [프로젝트 설정] > [스크립트 속성] > 속성 추가
 *      속성: BOX_KEY   값: 기록 입력용 비밀번호 (아무 문자열)
 * 4. 오른쪽 위 [배포] > [새 배포] > 유형: 웹 앱
 *      실행 계정: 나 / 액세스 권한: 모든 사용자  → [배포] 후 권한 허용
 * 5. 나오는 웹 앱 URL(https://script.google.com/macros/s/.../exec)을
 *    박스스코어 페이지 [박스스코어] 탭 > '구글 시트 연동'에 붙여넣기
 *
 * 스크립트를 고친 뒤에는 [배포] > [배포 관리] > 연필 > 버전: 새 버전 으로 다시 배포해야 반영돼요.
 */

const NAME_COL = { a: 1, b: 21 };   // A열, U열 (1부터 셈)
const WIDTH = 19;                   // PLAYER ~ PF

function doGet() {
  return json({ ok: true, msg: 'Momentum 박스스코어 연동이 켜져 있어요.' });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const key = PropertiesService.getScriptProperties().getProperty('BOX_KEY');
    if (!key || req.key !== key) return json({ ok: false, msg: '비밀번호가 맞지 않아요.' });

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      return json(writeGame(req));
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return json({ ok: false, msg: String(err && err.message || err) });
  }
}

// req: { week: 3, teams: ['팀A', '팀B'], rows: [[...19칸], ...] x2, dryRun }
function writeGame(req) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tab = 'Week ' + ('0' + req.week).slice(-2);
  const sh = ss.getSheetByName(tab);
  if (!sh) return { ok: false, msg: `'${tab}' 탭이 없어요.` };

  const last = Math.max(sh.getLastRow(), 1);
  const colA = sh.getRange(1, 1, last, 1).getDisplayValues().map(r => r[0].trim());
  const colU = sh.getRange(1, 21, last, 1).getDisplayValues().map(r => r[0].trim());
  const [ta, tb] = req.teams.map(s => String(s).trim());

  // 'AB MATCH' 같은 제목 줄을 찾고, 두 줄 아래의 팀 이름으로 경기 블록을 골라요.
  let start = -1, swap = false;
  for (let r = 0; r < last; r++) {
    if (!/MATCH\s*$/i.test(colA[r])) continue;
    const a = colA[r + 2], b = colU[r + 2];
    if (a === ta && b === tb) { start = r; break; }
    if (a === tb && b === ta) { start = r; swap = true; break; }
  }
  if (start < 0) return { ok: false, msg: `${tab}에서 '${ta} vs ${tb}' 경기를 못 찾았어요. 팀 이름이 시트와 같은지 확인해주세요.` };

  // PLAYER 머리글 다음 줄부터 TEAM 줄 전까지가 선수 칸
  let head = -1, team = -1;
  for (let r = start + 1; r < last; r++) {
    if (head < 0 && /^PLAYER$/i.test(colA[r])) head = r;
    else if (head >= 0 && /^(TEAM|TOTAL|합계)$/i.test(colA[r])) { team = r; break; }
    else if (/MATCH\s*$/i.test(colA[r])) break;
  }
  if (head < 0 || team < 0) return { ok: false, msg: '경기 블록의 PLAYER / TEAM 줄을 못 찾았어요.' };

  const first = head + 2;              // 시트 줄 번호 (1부터)
  const slots = team - head - 1;
  const sides = swap ? [['a', req.rows[1]], ['b', req.rows[0]]] : [['a', req.rows[0]], ['b', req.rows[1]]];
  for (const [, rows] of sides) if (rows.length > slots) return { ok: false, msg: `선수 칸이 ${slots}칸인데 ${rows.length}명이에요.` };
  if (req.dryRun) return { ok: true, msg: `${tab} ${colA[start]} 블록을 찾았어요.`, tab, swap };

  let written = 0;
  for (const [side, rows] of sides) {
    const range = sh.getRange(first, NAME_COL[side], slots, WIDTH);
    const formulas = range.getFormulas();
    const values = range.getValues();
    // 수식이 있는 칸(PTS, REB, % 등)은 그대로 두고, 나머지 칸만 새 기록으로 덮어써요.
    for (let r = 0; r < slots; r++) {
      for (let c = 0; c < WIDTH; c++) {
        if (formulas[r][c]) continue;
        const v = rows[r] ? rows[r][c] : '';
        values[r][c] = v === null || v === undefined ? '' : v;
      }
      if (rows[r]) written++;
    }
    // 수식 칸은 수식을 그대로 다시 써서 보존
    for (let r = 0; r < slots; r++) for (let c = 0; c < WIDTH; c++) if (formulas[r][c]) values[r][c] = formulas[r][c];
    range.setValues(values);
  }
  return { ok: true, msg: `${tab} ${colA[start]}에 ${written}명 기록을 넣었어요.`, tab, swap };
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
