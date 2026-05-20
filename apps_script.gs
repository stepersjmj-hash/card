/**
 * 가계부 변환기 - 구글시트 동기화 Apps Script (MJ/HJ 분리 + append-only)
 *
 * 시트 구조 (5열):
 *   A=거래일자  B=이름  C=이름(중복, 사용자 임의 수정 가능)  D=금액  E=카테고리
 *
 * 시트명 규칙: MJ2026-05, HJ2026-05 (prefix + YYYY-MM)
 *
 * 진단:
 *   - 동기화 시마다 'SYNC_LOG' 시트에 키 비교 결과 기록 (어떤 키가 매칭됐는지 등)
 *   - 매칭 안 돼서 중복이 생기는 원인을 직접 눈으로 확인 가능
 */

function doPost(e) {
  try {
    let body = e.postData ? e.postData.contents : '';
    if (body && body.indexOf('json=') === 0) body = body.substring(5);
    if (!body && e.parameter && e.parameter.json) body = e.parameter.json;
    const data = JSON.parse(body);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetsData = data.sheets || {};
    const stats = {};
    const logRows = [['시각', '시트명', '동작', '기존행수', '신규MJ/HJ', '매칭샘플', '신규샘플', '추가된키(범인)']];

    const names = Object.keys(sheetsData).sort();
    for (const name of names) {
      const rows = sheetsData[name];
      if (!Array.isArray(rows) || rows.length === 0) continue;

      const isMJ = name.indexOf('MJ') === 0;
      let sh = ss.getSheetByName(name);
      if (!sh) {
        sh = ss.insertSheet(name);
        const result = writeNewSheet(sh, rows, isMJ);
        stats[name] = result;
        logRows.push([new Date(), name, '신규시트생성', 0, rows.length - 1, '', '']);
      } else {
        // 자동 dedup: append 전에 기존 시트의 중복 행 먼저 정리
        const dedupRemoved = dedupSheet(sh);
        const result = appendOnly(sh, rows);
        stats[name] = result;
        if (dedupRemoved > 0) stats[name].dedupedBefore = dedupRemoved;
        logRows.push([new Date(), name,
                      dedupRemoved > 0 ? ('dedup ' + dedupRemoved + ' + append') : 'append',
                      result.existingCount, result.added,
                      JSON.stringify(result.sampleExisting),
                      JSON.stringify(result.sampleNew),
                      JSON.stringify(result.addedKeys)]);
      }
      updateTotal(sh);
    }

    // SYNC_LOG 시트에 진단 정보 기록 (계속 누적)
    writeSyncLog(ss, logRows);

    return ContentService
      .createTextOutput(JSON.stringify({ok: true, stats: stats}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ok: false, error: err.toString(), stack: err.stack}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput("가계부 변환기 동기화 엔드포인트 - POST 요청만 처리합니다.")
    .setMimeType(ContentService.MimeType.TEXT);
}

function writeNewSheet(sh, rows, isMJ) {
  const padded = rows.map(function(r) {
    const out = r.slice();
    while (out.length < 5) out.push('');
    return out;
  });
  sh.getRange(1, 1, padded.length, 5).setValues(padded);
  sh.getRange(1, 1, 1, 5).setFontWeight('bold').setHorizontalAlignment('center');
  if (padded.length >= 2) {
    sh.getRange(2, 4, padded.length - 1, 1).setNumberFormat('#,##0');
  }
  sh.getRange('D1').setNumberFormat('#,##0');
  sh.getRange('A1:E1').setBackground(isMJ ? '#BDD7EE' : '#F8CBAD');
  sh.setColumnWidth(1, 60);
  sh.setColumnWidth(2, 240);
  sh.setColumnWidth(3, 240);
  sh.setColumnWidth(4, 100);
  sh.setColumnWidth(5, 120);
  return { created: true, totalRows: padded.length - 1 };
}

function appendOnly(sh, rows) {
  const lastRow = sh.getLastRow();
  const existing = {};
  const sampleExisting = [];
  if (lastRow >= 2) {
    const existingRange = sh.getRange(2, 1, lastRow - 1, 5).getValues();
    for (let i = 0; i < existingRange.length; i++) {
      const r = existingRange[i];
      const key = makeKey(r[0], r[1], r[3], r[2]);
      if (key) {
        existing[key] = true;
        if (sampleExisting.length < 5) sampleExisting.push(key);
      }
    }
  }

  const newRows = [];
  const sampleNew = [];
  const addedKeys = [];  // 진단: 새로 추가된 행들의 키
  let matchedCount = 0;
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const key = makeKey(r[0], r[1], r[3], r[2]);
    if (!key) continue;
    if (sampleNew.length < 5) sampleNew.push(key);
    if (existing[key]) {
      matchedCount++;
    } else {
      newRows.push([r[0] || '', r[1] || '', r[2] || '', r[3] || '', r[4] || '']);
      addedKeys.push(key);
      existing[key] = true;
    }
  }

  if (newRows.length > 0) {
    const startRow = Math.max(lastRow, 1) + 1;
    sh.getRange(startRow, 1, newRows.length, 5).setValues(newRows);
    sh.getRange(startRow, 4, newRows.length, 1).setNumberFormat('#,##0');
  }
  return {
    added: newRows.length,
    matched: matchedCount,
    existingCount: Object.keys(existing).length - newRows.length,
    sampleExisting: sampleExisting,
    sampleNew: sampleNew,
    addedKeys: addedKeys,
  };
}

// 매칭 키 생성 (보강: B/C 중 값 있는 쪽 사용, Number 정규화, 공백 정리)
function makeKey(date, nameB, amount, nameC) {
  if (date === '' || date == null) return null;
  if (amount === '' || amount == null) return null;

  // 이름: B 우선, B가 비어있으면 C 사용. 둘 다 비어있으면 null.
  let name = nameB;
  if (name === '' || name == null) name = nameC;
  if (name === '' || name == null) return null;

  // 날짜: Date 객체 → "M/D" 형식. 문자열이면 trim만.
  let dateStr;
  if (date instanceof Date) {
    dateStr = (date.getMonth() + 1) + '/' + date.getDate();
  } else {
    dateStr = String(date).trim();
  }

  // 이름: 공백 정리 (앞뒤 trim + 중복 공백 1개로)
  let nameStr = String(name).trim().replace(/\s+/g, ' ');
  // 가맹점명이 순수 숫자라면 Number 정규화 (시트의 자동 number 변환으로 leading-zero가
  // 떨어진 케이스와 동일한 키를 만들도록. 예: "07011604744834" → "7011604744834")
  if (/^\d+$/.test(nameStr)) {
    nameStr = String(Number(nameStr));
  }

  // 금액: 항상 Number로 정규화 후 정수 문자열로
  const amountNum = Number(String(amount).replace(/,/g, ''));
  if (isNaN(amountNum)) return null;
  const amountStr = String(amountNum);

  return dateStr + '|' + nameStr + '|' + amountStr;
}

function updateTotal(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  let sum = 0;
  const data = sh.getRange(2, 4, lastRow - 1, 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (typeof data[i][0] === 'number') sum += data[i][0];
  }
  sh.getRange('D1').setValue(sum).setNumberFormat('#,##0');
}

/**
 * 진단 함수: 모든 MJ/HJ 시트를 훑어서 13자리 이상 숫자 가맹점(예: 7011604744834)을 찾고
 * 그 셀의 raw value, type, 매칭 키를 'DEBUG' 시트에 정리.
 * Apps Script 편집기 함수 목록에서 'debugLongNumber' 선택 → ▶ 실행.
 */
function debugLongNumber() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let debug = ss.getSheetByName('DEBUG');
  if (debug) debug.clear();
  else debug = ss.insertSheet('DEBUG');

  debug.appendRow([
    '시트명', '행', 'A 값', 'A 타입', 'B 값', 'B 타입',
    'C 값', 'C 타입', 'D 값', 'D 타입', '키(makeKey 결과)'
  ]);

  const sheets = ss.getSheets();
  let foundCount = 0;
  for (let s = 0; s < sheets.length; s++) {
    const sh = sheets[s];
    const name = sh.getName();
    if (name === 'DEBUG' || name === 'SYNC_LOG') continue;
    if (name.indexOf('MJ') !== 0 && name.indexOf('HJ') !== 0) continue;

    const lastRow = sh.getLastRow();
    if (lastRow < 2) continue;
    const data = sh.getRange(2, 1, lastRow - 1, 5).getValues();
    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      const bStr = String(r[1] == null ? '' : r[1]);
      // 가맹점명이 10자리 이상 숫자만으로 구성된 행 (예: 7011604744834, 230361096228)
      if (/^\d{10,}$/.test(bStr.replace(/\.\d+$/, ''))) {
        const key = makeKey(r[0], r[1], r[3], r[2]);
        const a = r[0];
        let aDesc = String(a);
        if (a instanceof Date) aDesc += ' [Date: ' + a.getFullYear() + '-' + (a.getMonth()+1) + '-' + a.getDate() + ']';
        debug.appendRow([
          name, i + 2,
          aDesc, typeof a + (a instanceof Date ? '/Date' : ''),
          String(r[1]), typeof r[1],
          String(r[2]), typeof r[2],
          String(r[3]), typeof r[3],
          key
        ]);
        foundCount++;
      }
    }
  }
  debug.appendRow([]);
  debug.appendRow(['총 ' + foundCount + '개 행 발견']);
  debug.appendRow(['같은 키가 여러 번 나오면 그 행들은 사실 같은 거래로 매칭됐어야 하는 케이스입니다.']);
  debug.appendRow(['키 형태/타입에 차이가 있다면 거기가 매칭 실패 원인.']);
  return 'DEBUG 시트에 ' + foundCount + '개 행 정리됨';
}

/**
 * 한 시트의 중복 행 제거. 같은 키(날짜+이름+금액)면 첫 번째만 남기고 나머지 제거.
 * 반환값: 제거된 행 수.
 *
 * 이 함수는 doPost에서 동기화 시 자동 호출되고, 수동으로 dedupAll을 통해서도 호출 가능.
 */
function dedupSheet(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;

  const data = sh.getRange(2, 1, lastRow - 1, 5).getValues();
  const seen = {};
  const keep = [];
  let removed = 0;
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    const key = makeKey(r[0], r[1], r[3], r[2]);
    if (!key) {
      keep.push(r);
      continue;
    }
    if (seen[key]) {
      removed++;
      continue;
    }
    seen[key] = true;
    keep.push(r);
  }

  if (removed > 0) {
    sh.getRange(2, 1, data.length, 5).clearContent();
    if (keep.length > 0) {
      sh.getRange(2, 1, keep.length, 5).setValues(keep);
      sh.getRange(2, 4, keep.length, 1).setNumberFormat('#,##0');
    }
  }
  return removed;
}

/**
 * 수동 dedup: 모든 MJ/HJ 시트를 한 번에 정리.
 * Apps Script 편집기에서 'dedupAll' 선택 → ▶ 실행. (보통 doPost가 자동으로 처리하니
 * 따로 안 눌러도 됩니다. 그래도 한 번 손으로 청소하고 싶을 때 쓰면 됨.)
 */
function dedupAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let totalRemoved = 0;
  const summary = [];

  for (let s = 0; s < sheets.length; s++) {
    const sh = sheets[s];
    const name = sh.getName();
    if (name === 'DEBUG' || name === 'SYNC_LOG') continue;
    if (name.indexOf('MJ') !== 0 && name.indexOf('HJ') !== 0) continue;

    const removed = dedupSheet(sh);
    if (removed > 0) {
      summary.push(name + ': -' + removed + '행');
      totalRemoved += removed;
    }
    updateTotal(sh);
  }
  return '✅ 중복 제거 완료: ' + totalRemoved + '행. ' + (summary.length ? summary.join(', ') : '제거할 게 없었음.');
}

// 진단 로그: 'SYNC_LOG' 시트에 매번 추가
function writeSyncLog(ss, logRows) {
  let log = ss.getSheetByName('SYNC_LOG');
  if (!log) {
    log = ss.insertSheet('SYNC_LOG');
    log.getRange(1, 1, 1, logRows[0].length).setValues([logRows[0]]).setFontWeight('bold');
    log.setColumnWidth(1, 160);
    log.setColumnWidth(2, 120);
    log.setColumnWidth(3, 120);
    log.setColumnWidth(4, 80);
    log.setColumnWidth(5, 100);
    log.setColumnWidth(6, 400);
    log.setColumnWidth(7, 400);
  }
  // 헤더 제외한 데이터 부분만 append
  const dataRows = logRows.slice(1);
  if (dataRows.length > 0) {
    const startRow = log.getLastRow() + 1;
    log.getRange(startRow, 1, dataRows.length, dataRows[0].length).setValues(dataRows);
  }
}
