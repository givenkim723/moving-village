/**
 * 움직이는교회 — Family Table 참여 신청 수신 (Google Apps Script)
 *
 * 하는 일
 *  1. /newcomer 페이지에서 보낸 신청을 검증해 시트에 한 줄 추가
 *  2. 담당 교역자에게 이메일 알림 (스크립트 속성 NOTIFY_EMAIL)
 *  3. 알림톡 발송 (솔라피) — 스크립트 속성이 모두 채워졌을 때만 동작
 *  4. 매일 1회, 성도 등록하지 않은 신청은 신청일로부터 1년이 지나면 삭제
 *
 * 설치 방법은 docs/newcomer-setup.md 를 참고하세요.
 *
 * ── 스크립트 속성 (프로젝트 설정 → 스크립트 속성) ──────────────────
 *  NOTIFY_EMAIL            새 신청 알림을 받을 이메일 (쉼표로 여러 개)
 *  -- 아래는 알림톡 준비 후 입력 (하나라도 비어 있으면 알림톡을 건너뜀) --
 *  SOLAPI_API_KEY          솔라피 API Key
 *  SOLAPI_API_SECRET       솔라피 API Secret
 *  SOLAPI_PF_ID            카카오 채널 발신 프로필 ID (pfId)
 *  SOLAPI_FROM             솔라피에 등록한 발신번호 (예: 01012345678)
 *  STAFF_PHONE             교역자 알림을 받을 휴대폰 번호
 *  TPL_STAFF               교역자용 알림톡 템플릿 ID
 *  TPL_WELCOME             신청자용 알림톡 템플릿 ID
 */

const SHEET_NAME = 'Family Table 신청';
const RETENTION_DAYS = 365;
const REGISTERED_STATUS = '성도 등록';

const STATUSES = ['신청', '연락함', '1주차', '2주차', '3주차', '성도 등록', 'SC 연결', '보류'];

const HEADERS = [
  '신청일시', '이름', '연락처', '선호 연락 방법', '연락 가능 시간', '참여 희망',
  '알게 된 경로', '출생 연도', '동반인', '기대·궁금한 점', '동의 일시', '유입',
  '진행 상태', '담당 교역자', '메모', '파기 예정일', '알림 결과',
];
const COL = Object.fromEntries(HEADERS.map((h, i) => [h, i + 1]));

const CHOICES = {
  contactMethod: ['카카오톡', '문자', '전화'],
  contactTime: ['평일 낮', '평일 저녁', '주말', '상관없음'],
  joinIntent: ['참여하고 싶어요', '조금 더 생각해 볼게요'],
  source: ['지인 소개', '인터넷·SNS', '지나가다', '기타'],
};

/* ════════════════════ 웹앱 진입점 ════════════════════ */

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    // 봇: 숨김 칸이 채워졌거나 3초 안에 제출 → 저장하지 않고 성공처럼 응답
    if (body.website || (Number(body.elapsed) > 0 && Number(body.elapsed) < 3000)) {
      return json_({ ok: true });
    }

    const data = validate_(body);

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    let row;
    try {
      row = appendRow_(data);
    } finally {
      lock.releaseLock();
    }

    // 알림은 실패해도 신청 자체는 성공으로 처리
    const results = [];
    results.push(notifyEmail_(data));
    results.push(...notifyAlimtalk_(data));
    sheet_().getRange(row, COL['알림 결과']).setValue(results.filter(Boolean).join(' / '));

    return json_({ ok: true });
  } catch (err) {
    console.error(err);
    const msg = err && err.userMessage ? err.userMessage : 'server_error';
    return json_({ ok: false, error: msg });
  }
}

function doGet() {
  return json_({ ok: true, service: 'newcomer' });
}

/* ════════════════════ 검증·저장 ════════════════════ */

function validate_(b) {
  const fail = (m) => { const e = new Error(m); e.userMessage = m; throw e; };
  const str = (v, max) => String(v == null ? '' : v).trim().slice(0, max);
  const pick = (v, list) => (list.indexOf(v) >= 0 ? v : '');

  const name = str(b.name, 30);
  const digits = String(b.phone || '').replace(/\D/g, '');
  if (!name) fail('name_required');
  // 페이지(newcomer/index.html RULES)와 같은 기준: 2자 이상, 글자로 시작, 글자·공백·가운뎃점·하이픈만
  if (name.length < 2 || !/^[\p{L}][\p{L} .·'-]*$/u.test(name)) fail('name_invalid');
  if (!/^01[016789]\d{7,8}$/.test(digits)) fail('phone_invalid');
  if (b.consent !== true) fail('consent_required');

  return {
    name,
    phone: formatPhone_(digits),
    phoneDigits: digits,
    contactMethod: pick(b.contactMethod, CHOICES.contactMethod),
    contactTime: pick(b.contactTime, CHOICES.contactTime),
    joinIntent: pick(b.joinIntent, CHOICES.joinIntent),
    source: pick(b.source, CHOICES.source),
    birthYear: birthYear_(b.birthYear),
    companions: /\p{L}/u.test(str(b.companions, 100)) ? str(b.companions, 100) : '',  // 글자 없는 값은 버림
    hope: str(b.hope, 1000),
    src: str(b.src, 30).replace(/[^\w-]/g, ''),
  };
}

// 출생 연도(선택): 1920년~올해의 네 자리 숫자만 받는다. 페이지와 같은 기준
function birthYear_(v) {
  const s = String(v == null ? '' : v).trim();
  if (!s) return '';
  const y = Number(s);
  if (!/^\d{4}$/.test(s) || y < 1920 || y > new Date().getFullYear()) {
    const e = new Error('birth_year_invalid'); e.userMessage = 'birth_year_invalid'; throw e;
  }
  return s;
}

function formatPhone_(d) {
  return d.length === 10
    ? d.slice(0, 3) + '-' + d.slice(3, 6) + '-' + d.slice(6)
    : d.slice(0, 3) + '-' + d.slice(3, 7) + '-' + d.slice(7);
}

// 수식 주입 방지: = + - @ 로 시작하는 값은 글자로 저장
function safe_(v) {
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function appendRow_(d) {
  const sh = sheet_();
  const now = new Date();
  const purge = new Date(now.getTime() + RETENTION_DAYS * 864e5);
  sh.appendRow([
    now, safe_(d.name), d.phone, d.contactMethod, d.contactTime, d.joinIntent,
    d.source, d.birthYear, safe_(d.companions), safe_(d.hope), now, d.src,
    '신청', '', '', purge, '',
  ]);
  return sh.getLastRow();
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || renameLegacy_(ss) || setupSheet_(ss);
  syncHeaders_(sh);
  return sh;
}

// 2026-11-28 '연령대' → '출생 연도'처럼 열 이름이 바뀌면 기존 시트 머리글을 맞춘다(열 위치는 그대로)
function syncHeaders_(sh) {
  const range = sh.getRange(1, 1, 1, HEADERS.length);
  const cur = range.getValues()[0];
  if (cur.some((v, i) => v !== HEADERS[i])) range.setValues([HEADERS]);
}

// 명칭 변경 이력(새이웃 → 뉴패밀리 → Family Table): 옛 탭이 있으면 이름만 바꿔 데이터를 그대로 쓴다
const LEGACY_SHEET_NAMES = ['뉴패밀리 신청', '새이웃 신청'];
function renameLegacy_(ss) {
  for (const n of LEGACY_SHEET_NAMES) {
    const sh = ss.getSheetByName(n);
    if (sh) return sh.setName(SHEET_NAME);
  }
  return null;
}

/* ════════════════════ 알림 ════════════════════ */

function notifyEmail_(d) {
  const to = prop_('NOTIFY_EMAIL');
  if (!to) return '';
  try {
    const lines = [
      'Family Table 참여 신청이 접수되었습니다.',
      '',
      '이름: ' + d.name,
      '연락처: ' + d.phone,
      d.contactMethod ? '선호 연락 방법: ' + d.contactMethod : null,
      d.contactTime ? '연락 가능 시간: ' + d.contactTime : null,
      d.joinIntent ? '참여 희망: ' + d.joinIntent : null,
      d.source ? '알게 된 경로: ' + d.source : null,
      d.birthYear ? '출생 연도: ' + d.birthYear : null,
      d.companions ? '동반인: ' + d.companions : null,
      d.hope ? '기대·궁금한 점: ' + d.hope : null,
      '',
      '시트 열기: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl(),
    ].filter((l) => l !== null);
    MailApp.sendEmail({
      to,
      subject: '[Family Table] ' + d.name + '님 참여 신청',
      body: lines.join('\n'),
    });
    return '메일 발송';
  } catch (err) {
    console.error(err);
    return '메일 실패';
  }
}

function notifyAlimtalk_(d) {
  const keys = ['SOLAPI_API_KEY', 'SOLAPI_API_SECRET', 'SOLAPI_PF_ID', 'STAFF_PHONE', 'TPL_STAFF', 'TPL_WELCOME'];
  if (keys.some((k) => !prop_(k))) return [];

  const pfId = prop_('SOLAPI_PF_ID');
  const from = prop_('SOLAPI_FROM');
  const base = (to, templateId, variables) => {
    const m = { to, type: 'ATA', kakaoOptions: { pfId, templateId, variables, disableSms: true } };
    if (from) m.from = from;
    return m;
  };

  // 템플릿 변수 이름은 심사받은 템플릿과 정확히 같아야 합니다 (docs/newcomer-setup.md)
  const messages = [
    base(prop_('STAFF_PHONE').replace(/\D/g, ''), prop_('TPL_STAFF'), {
      '#{이름}': d.name,
      '#{연락처}': d.phone,
      '#{참여희망}': d.joinIntent || '미선택',
    }),
    base(d.phoneDigits, prop_('TPL_WELCOME'), {
      '#{이름}': d.name,
    }),
  ];

  try {
    const res = solapi_('/messages/v4/send-many/detail', { messages });
    const failed = (res.failedMessageList || []).length;
    return [failed ? '알림톡 일부 실패(' + failed + ')' : '알림톡 발송'];
  } catch (err) {
    console.error(err);
    return ['알림톡 실패'];
  }
}

function solapi_(path, payload) {
  const apiKey = prop_('SOLAPI_API_KEY');
  const secret = prop_('SOLAPI_API_SECRET');
  const date = new Date().toISOString();
  const salt = Utilities.getUuid().replace(/-/g, '');
  const sig = Utilities.computeHmacSha256Signature(date + salt, secret)
    .map((b) => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');

  const res = UrlFetchApp.fetch('https://api.solapi.com' + path, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'HMAC-SHA256 apiKey=' + apiKey + ', date=' + date + ', salt=' + salt + ', signature=' + sig,
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  const text = res.getContentText();
  if (code >= 300) throw new Error('Solapi ' + code + ': ' + text);
  return JSON.parse(text);
}

/* ════════════════════ 1년 파기 ════════════════════ */

function purgeExpired() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return;
  const values = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  const cutoff = Date.now() - RETENTION_DAYS * 864e5;
  let removed = 0;
  // 아래에서 위로 지워야 행 번호가 밀리지 않습니다
  for (let i = values.length - 1; i >= 0; i--) {
    const applied = values[i][COL['신청일시'] - 1];
    const status = values[i][COL['진행 상태'] - 1];
    if (status === REGISTERED_STATUS) continue;
    if (applied instanceof Date && applied.getTime() < cutoff) {
      sh.deleteRow(i + 2);
      removed++;
    }
  }
  if (removed) console.log('파기: ' + removed + '건');
}

/* ════════════════════ 최초 설정 (한 번만 실행) ════════════════════ */

/** 편집기에서 이 함수를 선택해 한 번 실행하세요. 시트·드롭다운·파기 트리거를 만듭니다. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SHEET_NAME)) renameLegacy_(ss);
  setupSheet_(ss);
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'purgeExpired')
    .forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('purgeExpired').timeBased().everyDays(1).atHour(4).create();
  console.log('설정 완료');
}

function setupSheet_(ss) {
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#F4ECDF');
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);

  const rows = sh.getMaxRows() - 1;
  sh.getRange(2, COL['신청일시'], rows, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange(2, COL['동의 일시'], rows, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange(2, COL['파기 예정일'], rows, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, COL['연락처'], rows, 1).setNumberFormat('@');

  const rule = SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false).build();
  sh.getRange(2, COL['진행 상태'], rows, 1).setDataValidation(rule);

  // 성도 등록된 행은 파기 대상이 아니므로 연한 초록으로 표시
  const statusCol = columnLetter_(COL['진행 상태']);
  const cf = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + statusCol + '2="' + REGISTERED_STATUS + '"')
    .setBackground('#E3F3E8')
    .setRanges([sh.getRange(2, 1, rows, HEADERS.length)])
    .build();
  sh.setConditionalFormatRules([cf]);

  sh.setColumnWidth(COL['기대·궁금한 점'], 280);
  sh.setColumnWidth(COL['메모'], 220);
  return sh;
}

/* ════════════════════ 유틸 ════════════════════ */

function prop_(k) {
  return (PropertiesService.getScriptProperties().getProperty(k) || '').trim();
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function columnLetter_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

/** 편집기에서 실행해 이메일 알림을 시험합니다 (시트에는 저장하지 않음). */
function testNotify() {
  const d = validate_({ name: '테스트', phone: '010-0000-0000', consent: true, joinIntent: '참여하고 싶어요' });
  console.log(notifyEmail_(d));
}
