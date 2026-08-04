// =====================================================
// 호르메 몬테소리 부모교육 — Google Apps Script
// 신청 수신 → 구글 시트 기록 + 솔라피 문자 발송
// =====================================================

// ▼▼▼ 여기에 솔라피 정보 입력 ▼▼▼
const SOLAPI_API_KEY    = '';
const SOLAPI_API_SECRET = '';
const SENDER_PHONE      = '01031134221'; // 발신번호
// ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

// 열 번호 (1부터 시작) — "기수" 열 추가로 한 칸씩 밀림
const COL_TIME    = 1;  // 신청시간
const COL_ROUND   = 2;  // 기수 (예: "4기")
const COL_COURSE  = 3;  // 과정구분 (기본교육 / 감각교육)
const COL_NAME    = 4;  // 이름
const COL_PHONE   = 5;  // 연락처
const COL_AGE     = 6;  // 아이나이
const COL_BANK    = 7;  // 은행
const COL_ACCOUNT = 8;  // 계좌번호
const COL_HOLDER  = 9;  // 예금주
const COL_RECEIPT = 10; // 현금영수증
const COL_CONFIRM = 11; // 입금확인 ✅

const MAX_COUNT = 40; // 기수·과정별 마감 인원

// ── 입금 카운트 조회 (페이지에서 GET 호출) ────────────
// 반환 형태: { "4기_기본교육": { count, closed }, "4기_감각교육": { count, closed }, ... }
function doGet(_e) {
  const ss    = SpreadsheetApp.openById('1_LAxbBSHoT-M_akuqvwwBB2jP5chUD4Li8bFB4An3mk');
  const sheet = ss.getActiveSheet();
  const data  = sheet.getDataRange().getValues();

  const counts = {};

  for (let i = 1; i < data.length; i++) {
    const round   = data[i][COL_ROUND - 1]   ? data[i][COL_ROUND - 1].toString().trim()   : '';
    const course  = data[i][COL_COURSE - 1]  ? data[i][COL_COURSE - 1].toString().trim()  : '';
    const confirm = data[i][COL_CONFIRM - 1] ? data[i][COL_CONFIRM - 1].toString().trim() : '';
    const paid = confirm === '입금' || confirm === '입금확인' || confirm === 'O' || confirm === 'o' || confirm === '✅';
    if (!paid || !round || !course) continue;

    const key = round + '_' + course;
    counts[key] = (counts[key] || 0) + 1;
  }

  const result = {};
  Object.keys(counts).forEach(function (key) {
    result[key] = { count: counts[key], closed: counts[key] >= MAX_COUNT };
  });

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 신청 수신 (폼 제출 시) ──────────────────────────
function doPost(e) {
  const ss    = SpreadsheetApp.openById('1_LAxbBSHoT-M_akuqvwwBB2jP5chUD4Li8bFB4An3mk');
  const sheet = ss.getActiveSheet();

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['신청시간', '기수', '과정구분', '이름', '연락처', '아이나이', '은행', '계좌번호', '예금주', '현금영수증', '입금확인']);
    sheet.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#fce5cd');
  }

  const payload = JSON.parse(e.postData.contents);
  const data    = payload.data || payload;

  const 신청시간 = payload.created_at
    ? new Date(payload.created_at).toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' })
    : new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' });

  sheet.appendRow([
    신청시간,
    data['기수']     || '',
    data['과정구분'] || '',
    data['이름']     || '',
    '',
    data['아이나이'] || '',
    data['은행']     || '',
    data['계좌번호'] || '',
    data['예금주']   || '',
    data['현금영수증'] || '',
    '', // 입금확인 빈칸
  ]);
  // 연락처는 텍스트로 저장해야 앞 0이 안 날아감
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, COL_PHONE).setNumberFormat('@').setValue(data['연락처'] || '');

  // 신청 접수 문자 발송
  const 이름   = data['이름']   || '';
  const 연락처 = data['연락처'] || '';

  if (연락처) {
    const 수신번호 = 연락처.replace(/-/g, '');
    const 문자내용 = `[키키맘] ${이름}님, 몬테소리 부모교육 신청해주셔서 감사합니다!\n\n신청은 입금 순으로 마감되니 참고 부탁드립니다.\n입금 계좌: 농협 301-0316-8861-41 (주식회사 슈필하우스)\n입금 금액: 119,900원\n입금자명: 이름+핸드폰 뒷 4자리\n예) ${이름}${수신번호.slice(-4)}\n\n감사합니다.`;
    sendSms(수신번호, 문자내용);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ result: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 입금확인 체크 시 문자 발송 ──────────────────────
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;

  if (range.getColumn() !== COL_CONFIRM) return;
  const value = range.getValue().toString().trim();
  if (value !== 'O' && value !== 'o' && value !== '✅' && value !== '입금' && value !== '입금확인') return;

  const row = range.getRow();
  if (row <= 1) return; // 헤더 제외

  const 이름   = sheet.getRange(row, COL_NAME).getValue();
  const 연락처 = sheet.getRange(row, COL_PHONE).getValue();

  if (!연락처) return;

  const 수신번호 = 연락처.toString().replace(/-/g, '');
  const 문자내용 = `[키키맘] ${이름}님, 입금이 확인되었습니다!\n\n4주 교육이 끝나면 불안한 육아가 자신감과 확신으로 바뀔 겁니다.\n\n좋은 시간을 위해 저희도 최선을 다하겠습니다. 감사합니다.`;

  sendSms(수신번호, 문자내용);
}

// ── 솔라피 문자 발송 함수 ───────────────────────────
function sendSms(to, text) {
  const now       = new Date();
  const date      = Utilities.formatDate(now, 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
  const salt      = Utilities.getUuid();
  const signature = makeSignature(date, salt);

  const payload = {
    message: {
      to:   to,
      from: SENDER_PHONE,
      text: text,
    }
  };

  const options = {
    method:  'POST',
    headers: {
      'Content-Type':  'application/json',
      'Authorization': `HMAC-SHA256 apiKey=${SOLAPI_API_KEY}, date=${date}, salt=${salt}, signature=${signature}`,
    },
    payload: JSON.stringify(payload),
  };

  UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send', options);
}

function makeSignature(date, salt) {
  const message   = date + salt;
  const key       = Utilities.newBlob(SOLAPI_API_SECRET).getBytes();
  const msg       = Utilities.newBlob(message).getBytes();
  const signature = Utilities.computeHmacSha256Signature(msg, key);
  return signature.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}
