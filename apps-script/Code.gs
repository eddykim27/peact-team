/**
 * 피스팀 출퇴근 QR 체크 — 구글 시트에 붙이는 Apps Script
 *
 * 하는 일
 *  - 휴대폰 체크 페이지(fiaoffice.netlify.app/checkin)에서 출근/퇴근을 누르면
 *    누른 순간의 한국 시간을 '출퇴근기록' 탭에 적습니다.
 *  - 같은 날 줄이 있으면 그 줄에 채우고, 없으면 새 줄을 만듭니다.
 *  - 출근은 처음 누른 시간을 지키고, 퇴근은 마지막에 누른 시간으로 바뀝니다.
 *  - 새벽 5시 전에 퇴근을 누르면, 전날 퇴근 칸이 비어 있을 때 전날 퇴근으로 적습니다.
 *  - 모든 체크는 '체크인로그' 탭에 따로 남습니다(위치 포함).
 */

// ---------- 설정 ----------
const SECRET = 'fia-2026';          // QR 주소의 k= 값과 같아야 합니다
const OFFICE_LAT = 37.531712;       // 사무실 신발장 위치 (2026-10-18 실제 체크로 확인). null이면 위치 확인을 하지 않습니다
const OFFICE_LNG = 126.707942;      // 사무실 신발장 위치
const OFFICE_RADIUS_M = 200;        // 사무실에서 이 거리(미터) 안일 때만 체크
const OVERNIGHT_CUTOFF_HOUR = 5;    // 이 시각 전의 퇴근은 전날 퇴근으로 봅니다
const TZ = 'Asia/Seoul';

const TAB_LOG = '출퇴근기록';
const TAB_INFO = '기본정보';
const TAB_AUDIT = '체크인로그';

function doGet(e) {
  const p = (e && e.parameter) || {};
  let result;
  try {
    result = handle_(p);
  } catch (err) {
    result = { ok: false, message: '처리 중 문제가 생겼어요: ' + err.message };
  }
  const json = JSON.stringify(result);
  const cb = String(p.callback || '').replace(/[^\w.$]/g, '');
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function handle_(p) {
  if (p.ping) return { ok: true, message: 'pong' };
  if (String(p.k || '') !== SECRET) return { ok: false, code: 'key', message: '사무실 QR코드로 다시 열어주세요.' };

  const name = String(p.name || '').trim();
  const action = p.action === 'out' ? 'out' : p.action === 'in' ? 'in' : '';
  if (!name || !action) return { ok: false, message: '이름과 출근/퇴근을 골라주세요.' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!memberNames_(ss).includes(name)) return { ok: false, message: '기본정보에 없는 이름이에요: ' + name };

  // 위치 확인 (설정했을 때만)
  const lat = parseFloat(p.lat), lng = parseFloat(p.lng), acc = parseFloat(p.acc);
  let dist = null;
  if (OFFICE_LAT !== null && OFFICE_LNG !== null) {
    if (isNaN(lat) || isNaN(lng)) return { ok: false, code: 'loc', message: '위치 권한을 허용해야 체크할 수 있어요.' };
    dist = distanceM_(lat, lng, OFFICE_LAT, OFFICE_LNG);
    if (dist > OFFICE_RADIUS_M + (isNaN(acc) ? 0 : Math.min(acc, 200))) {
      audit_(ss, name, action, '거절(위치)', lat, lng, acc, dist);
      return { ok: false, code: 'far', message: '사무실 근처에서만 체크할 수 있어요.' };
    }
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const now = new Date();
    const hhmm = Utilities.formatDate(now, TZ, 'H:mm');
    const hour = Number(Utilities.formatDate(now, TZ, 'H'));
    let day = ymd_(now);

    const sh = ss.getSheetByName(TAB_LOG);
    let row = findRow_(sh, day, name);

    if (action === 'out' && hour < OVERNIGHT_CUTOFF_HOUR) {
      const prev = ymd_(new Date(now.getTime() - 24 * 3600 * 1000));
      const prevRow = findRow_(sh, prev, name);
      if (prevRow && !sh.getRange(prevRow, 6).getDisplayValue()) { day = prev; row = prevRow; }
    }

    let message, kept = false;
    if (action === 'in') {
      if (row) {
        const existing = sh.getRange(row, 5).getDisplayValue();
        if (existing) { kept = true; message = '오늘은 이미 ' + existing + '에 출근했어요.'; }
        else { setTime_(sh.getRange(row, 5), hhmm); message = hhmm + ' 출근 완료!'; }
      } else {
        appendRow_(sh, day, name, hhmm, '');
        message = hhmm + ' 출근 완료!';
      }
    } else {
      if (row) setTime_(sh.getRange(row, 6), hhmm);
      else appendRow_(sh, day, name, '', hhmm);
      message = hhmm + ' 퇴근 완료! 수고했어요.';
    }

    audit_(ss, name, action, kept ? '유지' : '기록', lat, lng, acc, dist);
    return { ok: true, name: name, action: action, time: hhmm, date: day.join('-'), kept: kept, message: message };
  } finally {
    lock.releaseLock();
  }
}

// ---------- 도우미 ----------
function ymd_(d) {
  return [Number(Utilities.formatDate(d, TZ, 'yyyy')), Number(Utilities.formatDate(d, TZ, 'M')), Number(Utilities.formatDate(d, TZ, 'd'))];
}

function findRow_(sh, ymd, name) {
  const last = sh.getLastRow();
  if (last < 2) return null;
  const vals = sh.getRange(2, 1, last - 1, 4).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    const r = vals[i];
    if (Number(r[0]) === ymd[0] && Number(r[1]) === ymd[1] && Number(r[2]) === ymd[2] && String(r[3]).trim() === name) return i + 2;
  }
  return null;
}

function setTime_(range, hhmm) {
  const [h, m] = hhmm.split(':').map(Number);
  range.setValue((h * 60 + m) / 1440).setNumberFormat('h:mm');
}

function appendRow_(sh, ymd, name, inT, outT) {
  const row = sh.getLastRow() + 1;
  sh.getRange(row, 1, 1, 4).setValues([[ymd[0], ymd[1], ymd[2], name]]);
  if (inT) setTime_(sh.getRange(row, 5), inT);
  if (outT) setTime_(sh.getRange(row, 6), outT);
}

function memberNames_(ss) {
  const sh = ss.getSheetByName(TAB_INFO);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0]).trim()).filter(Boolean);
}

function audit_(ss, name, action, result, lat, lng, acc, dist) {
  let sh = ss.getSheetByName(TAB_AUDIT);
  if (!sh) {
    sh = ss.insertSheet(TAB_AUDIT);
    sh.appendRow(['체크시각', '이름', '구분', '결과', '위도', '경도', '정확도(m)', '사무실까지(m)']);
    sh.setFrozenRows(1);
  }
  sh.appendRow([
    Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), name, action === 'in' ? '출근' : '퇴근', result,
    isNaN(lat) ? '' : lat, isNaN(lng) ? '' : lng, isNaN(acc) ? '' : Math.round(acc), dist === null ? '' : Math.round(dist)
  ]);
}

function distanceM_(lat1, lng1, lat2, lng2) {
  const R = 6371000, rad = x => x * Math.PI / 180;
  const dLat = rad(lat2 - lat1), dLng = rad(lng2 - lng1);
  const a = Math.sin(dLat / 2) ** 2 + Math.cos(rad(lat1)) * Math.cos(rad(lat2)) * Math.sin(dLng / 2) ** 2;
  return 2 * R * Math.asin(Math.sqrt(a));
}
