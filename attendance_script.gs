// ═══════════════════════════════════════════════════════════
// 출퇴근 기록 자동화 - Google Apps Script 최종본
// 기능: 출퇴근 기록 | 점심 1h 차감 | 공휴일/주말 빨간폰트
//       월별 법정근무일/목표시간 요약 | 8h 미달 체크
// ═══════════════════════════════════════════════════════════


// ───────────────────────────────────────────────────────────
// 1. 한국 공휴일 목록 (대체공휴일 포함)
// ───────────────────────────────────────────────────────────
function getKoreanHolidays() {
  return [
    // 2026년
    '2026-01-01', // 신정
    '2026-01-28', // 설날 연휴
    '2026-01-29', // 설날
    '2026-01-30', // 설날 연휴
    '2026-02-16', // 설날 대체공휴일 (1/28 일요일)
    '2026-03-01', // 삼일절 (일요일)
    '2026-03-02', // 삼일절 대체공휴일
    '2026-05-01', // 근로자의날
    '2026-05-05', // 어린이날
    '2026-05-24', // 부처님오신날 (일요일)
    '2026-05-25', // 부처님오신날 대체공휴일
    '2026-06-03', // 제9회 전국동시지방선거 (임시공휴일)
    '2026-06-06', // 현충일 (토요일, 대체공휴일 없음)
    '2026-07-17', // 제헌절 (공휴일 재지정)
    '2026-08-15', // 광복절 (토요일)
    '2026-08-17', // 광복절 대체공휴일
    '2026-09-24', // 추석 연휴
    '2026-09-25', // 추석
    '2026-09-26', // 추석 연휴
    '2026-09-27', // 추석 연휴
    '2026-10-03', // 개천절 (토요일)
    '2026-10-05', // 개천절 대체공휴일
    '2026-10-09', // 한글날
    '2026-12-25', // 크리스마스
  ];
}


// ───────────────────────────────────────────────────────────
// 2. 유틸 함수
// ───────────────────────────────────────────────────────────

// 휴일(주말+공휴일) 여부
function isHoliday(dateStr) {
  const date = new Date(dateStr + 'T00:00:00');
  const day  = date.getDay(); // 0=일, 6=토
  return day === 0 || day === 6 || getKoreanHolidays().includes(dateStr);
}

// 해당 월의 법정 근무일수 계산
function getWorkingDaysInMonth(year, month) {
  const holidays    = getKoreanHolidays();
  const daysInMonth = new Date(year, month, 0).getDate();
  let   workingDays = 0;

  for (let d = 1; d <= daysInMonth; d++) {
    const mm      = String(month).padStart(2, '0');
    const dd      = String(d).padStart(2, '0');
    const dateStr = `${year}-${mm}-${dd}`;
    const date    = new Date(dateStr + 'T00:00:00');
    const day     = date.getDay();
    if (day !== 0 && day !== 6 && !holidays.includes(dateStr)) {
      workingDays++;
    }
  }
  return workingDays;
}

// "HH:MM" → 분
function timeToMins(timeStr) {
  if (!timeStr || !String(timeStr).includes(':')) return 0;
  const parts = String(timeStr).split(':');
  return parseInt(parts[0]) * 60 + parseInt(parts[1]);
}

// 분 → "Xh Ym"
function minsToStr(mins) {
  if (mins <= 0) return '0h 0m';
  return `${Math.floor(mins / 60)}h ${mins % 60}m`;
}

// 실근무시간 계산 (점심 1시간 차감)
function calcReal(inT, outT) {
  if (!inT || !outT) return '';
  const mins = timeToMins(outT) - timeToMins(inT) - 60;
  return minsToStr(Math.max(0, mins));
}

// 총근무시간 계산 (차감 없음)
function calcTotal(inT, outT) {
  if (!inT || !outT) return '';
  const mins = timeToMins(outT) - timeToMins(inT);
  return minsToStr(Math.max(0, mins));
}


// ───────────────────────────────────────────────────────────
// 3. 시트 상단 요약 업데이트 (1~4행)
// ───────────────────────────────────────────────────────────
function updateSummary(sheet) {
  const now      = new Date();
  const year     = now.getFullYear();
  const month    = now.getMonth() + 1;
  const monthStr = `${year}-${String(month).padStart(2, '0')}`;

  const workingDays  = getWorkingDaysInMonth(year, month);
  const targetHours  = workingDays * 8;

  // 이번 달 실근무시간 합산 (5행~, 점심 1h 차감)
  const lastRow   = sheet.getLastRow();
  let actualMins  = 0;
  let workedDays  = 0;

  for (let i = 5; i <= lastRow; i++) {
    const dateVal = sheet.getRange(i, 1).getValue();
    if (!dateVal) continue;
    const dateStr = String(dateVal);
    if (!dateStr.startsWith(monthStr)) continue;

    const inT  = String(sheet.getRange(i, 2).getValue() || '');
    const outT = String(sheet.getRange(i, 3).getValue() || '');
    if (inT.includes(':') && outT.includes(':')) {
      const realMins = timeToMins(outT) - timeToMins(inT) - 60; // 점심 1h 차감
      actualMins += Math.max(0, realMins);
      workedDays++;
    }
  }

  const actualHours  = (actualMins / 60).toFixed(1);
  const remainMins   = Math.max(0, targetHours * 60 - actualMins);
  const remainHours  = (remainMins / 60).toFixed(1);
  const isComplete   = remainMins === 0;

  // ── 1행: 제목 (A1:F1 병합) ──
  try { sheet.getRange('A1:F1').merge(); } catch(e) {}
  sheet.getRange('A1')
    .setValue(`📅 ${year}년 ${month}월 근무 현황`)
    .setFontSize(13)
    .setFontWeight('bold')
    .setFontColor('#000000')
    .setBackground('#e8f0fe');

  // ── 2행: 법정근무일 / 목표근무시간 ──
  sheet.getRange('A2').setValue('법정 근무일').setFontWeight('bold');
  sheet.getRange('B2').setValue(`${workingDays}일`).setFontColor('#1a73e8').setFontWeight('bold');
  sheet.getRange('C2').setValue('목표 근무시간').setFontWeight('bold');
  sheet.getRange('D2').setValue(`${targetHours}h`).setFontColor('#1a73e8').setFontWeight('bold');
  sheet.getRange('E2').setValue('');
  sheet.getRange('F2').setValue('');

  // ── 3행: 출근일수 / 실근무시간 / 잔여 ──
  sheet.getRange('A3').setValue('출근 일수').setFontWeight('bold');
  sheet.getRange('B3').setValue(`${workedDays}일`).setFontColor('#0f9d58').setFontWeight('bold');
  sheet.getRange('C3').setValue('실근무시간').setFontWeight('bold');
  sheet.getRange('D3').setValue(`${actualHours}h`).setFontColor('#0f9d58').setFontWeight('bold');
  sheet.getRange('E3').setValue('잔여').setFontWeight('bold');
  sheet.getRange('F3')
    .setValue(`${remainHours}h`)
    .setFontColor(isComplete ? '#0f9d58' : '#ea4335')
    .setFontWeight('bold');

  // ── 2~3행 배경 ──
  sheet.getRange('A2:F2').setBackground('#f8f9fa').setFontSize(11);
  sheet.getRange('A3:F3').setBackground('#f8f9fa').setFontSize(11);

  // ── 4행: 헤더 ──
  const headers = ['날짜', '출근', '퇴근', '총근무', '실근무(점심제외)', '8h 달성'];
  headers.forEach((h, i) => {
    sheet.getRange(4, i + 1)
      .setValue(h)
      .setFontWeight('bold')
      .setBackground('#dadce0')
      .setFontColor('#000000');
  });
}


// ───────────────────────────────────────────────────────────
// 4. 행 스타일 적용
// ───────────────────────────────────────────────────────────
function applyRowStyle(sheet, row, dateStr, inT, outT) {
  // 전체 폰트 초기화
  sheet.getRange(row, 1, 1, 6).setFontColor('#000000').setBackground('#ffffff');

  if (isHoliday(dateStr)) {
    // 주말/공휴일 → 빨간 폰트
    sheet.getRange(row, 1, 1, 6).setFontColor('#ea4335');
    sheet.getRange(row, 6).setValue('휴일');

  } else if (inT && outT && String(inT).includes(':') && String(outT).includes(':')) {
    // 근무일 → 8h 달성 여부
    const realMins = timeToMins(outT) - timeToMins(inT) - 60;
    const achieved = realMins >= 480; // 8시간 = 480분

    sheet.getRange(row, 6)
      .setValue(achieved ? '✅' : '⚠️ 미달')
      .setFontColor(achieved ? '#0f9d58' : '#ea4335');

    if (!achieved) {
      // 미달인 경우 해당 행 연한 빨간 배경
      sheet.getRange(row, 1, 1, 6).setBackground('#fce8e6');
      sheet.getRange(row, 6).setBackground('#fce8e6');
    }
  }
}


// ───────────────────────────────────────────────────────────
// 5. 메인: 출퇴근 기록 수신 (웹앱 POST)
// ───────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const data  = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    const date  = data.date || '';
    const inT   = data.in   || '';
    const outT  = data.out  || '';
    const total = calcTotal(inT, outT);
    const real  = calcReal(inT, outT);

    // 데이터는 5행부터 (1~3: 요약, 4: 헤더)
    const lastRow = sheet.getLastRow();
    let targetRow = -1;
    let found     = false;

    for (let i = 5; i <= lastRow; i++) {
      if (String(sheet.getRange(i, 1).getValue()) === date) {
        targetRow = i;
        found     = true;
        break;
      }
    }

    if (!found) {
      targetRow = Math.max(lastRow + 1, 5);
      sheet.getRange(targetRow, 1).setValue(date);
    }

    if (inT)   sheet.getRange(targetRow, 2).setValue(inT);
    if (outT)  sheet.getRange(targetRow, 3).setValue(outT);
    if (total) sheet.getRange(targetRow, 4).setValue(total);
    if (real)  sheet.getRange(targetRow, 5).setValue(real);

    applyRowStyle(sheet, targetRow, date, inT, outT);
    updateSummary(sheet);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


// ───────────────────────────────────────────────────────────
// 6. 테스트용 함수 (Apps Script에서 직접 실행 가능)
// ───────────────────────────────────────────────────────────
function testRecord() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const testData = [
    { date: '2026-05-04', inT: '09:00', outT: '18:00' }, // 월 8h ✅
    { date: '2026-05-05', inT: '',      outT: ''      }, // 화 어린이날 🔴
    { date: '2026-05-06', inT: '09:10', outT: '18:30' }, // 수 8.3h ✅
    { date: '2026-05-07', inT: '09:00', outT: '16:30' }, // 목 6.5h ⚠️
    { date: '2026-05-08', inT: '08:50', outT: '18:10' }, // 금 8.3h ✅
    { date: '2026-05-09', inT: '',      outT: ''      }, // 토 주말 🔴
    { date: '2026-05-25', inT: '',      outT: ''      }, // 월 대체공휴일 🔴
  ];

  testData.forEach((d, i) => {
    const row   = 5 + i;
    const total = calcTotal(d.inT, d.outT);
    const real  = calcReal(d.inT, d.outT);
    sheet.getRange(row, 1).setValue(d.date);
    sheet.getRange(row, 2).setValue(d.inT);
    sheet.getRange(row, 3).setValue(d.outT);
    sheet.getRange(row, 4).setValue(total);
    sheet.getRange(row, 5).setValue(real);
    applyRowStyle(sheet, row, d.date, d.inT, d.outT);
  });

  updateSummary(sheet);
  Logger.log('테스트 완료!');
}
