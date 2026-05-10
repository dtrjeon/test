// ╔══════════════════════════════════════════════════════════════╗
// ║     전도폭발훈련 · 전문 구두점검 — 로그 수신 웹앱              ║
// ║     Google Apps Script  /  index-CP1-log.html 전용           ║
// ╚══════════════════════════════════════════════════════════════╝
//
// ─── 설치 순서 ────────────────────────────────────────────────
//  1. Google 스프레드시트 새 파일 생성
//  2. 메뉴 → 확장 프로그램 → Apps Script
//  3. 이 코드 전체를 붙여넣기 (기존 코드 전체 삭제 후)
//  4. 상단 CONFIG 값 필요시 수정
//  5. 저장 (Ctrl+S)
//  6. 배포 → 새 배포
//       유형            : 웹앱
//       다음 사용자로 실행: 나 (본인 Google 계정)
//       액세스 권한     : 모든 사람
//  7. 승인 → 고급 → "안전하지 않은 앱으로 이동" 클릭 후 허용
//  8. 웹앱 URL 복사 → HTML 파일 관리자 설정에 붙여넣기
// ──────────────────────────────────────────────────────────────

// ━━━━━━━━━━━━━━━━━━━━ CONFIG ━━━━━━━━━━━━━━━━━━━━
const CONFIG = {
  SHEET_NAME  : '점검기록',    // 메인 로그 시트 이름
  STATS_SHEET : '통계',        // 통계 시트 이름 (자동 생성)
  ADMIN_EMAIL : '',            // 알림 받을 이메일 (빈칸이면 알림 없음)
  ALERT_ON_LOW: 60,            // 이 점수 미만이면 이메일 알림 (0 = 알림없음)
};
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

// ─── 수신 데이터 구조 (HTML → 서버) ───
// {
//   time     : "2026.05.10 14:30"   점검 시각 (KST)
//   name     : "홍길동"              점검자 이름
//   chapter  : "2장 은혜(아니요)"    점검 챕터
//   accuracy : 87                   일치율 (숫자)
//   grade    : "우수"               등급 (완벽/우수/보통/미흡)
// }

// ══════════════════════════════════════════════════
//  POST 수신 — HTML fetch() 호출 시 실행
// ══════════════════════════════════════════════════
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // 필수 필드 검증
    if (!data.name || !data.chapter) {
      return jsonResponse({ result: 'error', message: '필수 필드 누락 (name, chapter)' });
    }

    const rowNum = appendLog(data);
    updateStats(data);

    // 저점수 알림
    if (CONFIG.ADMIN_EMAIL && CONFIG.ALERT_ON_LOW > 0
        && Number(data.accuracy) < CONFIG.ALERT_ON_LOW) {
      sendLowScoreAlert(data);
    }

    return jsonResponse({ result: 'ok', row: rowNum });

  } catch (err) {
    console.error('doPost error:', err);
    return jsonResponse({ result: 'error', message: err.message });
  }
}

// ══════════════════════════════════════════════════
//  GET — 브라우저 직접 접속 시 연결 확인용
// ══════════════════════════════════════════════════
function doGet(e) {
  // ?action=stats 로 접근하면 JSON 통계 반환
  if (e && e.parameter && e.parameter.action === 'stats') {
    return jsonResponse(getSummaryStats());
  }
  return ContentService
    .createTextOutput('✅ 구두점검 로그 수신 서버가 정상 동작 중입니다.')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ══════════════════════════════════════════════════
//  메인 로그 시트에 행 추가
// ══════════════════════════════════════════════════
function appendLog(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  // 시트 없으면 자동 생성
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    buildLogHeader(sheet);
  }

  const lastRow = sheet.getLastRow();
  const rowNum  = Math.max(lastRow, 1); // 헤더 제외 순번

  // 서버 수신 시각 (KST)
  const serverTime = getKST();

  sheet.appendRow([
    rowNum,                                    // A: 번호
    data.time        || '',                    // B: 점검 시각 (클라이언트)
    data.name        || '',                    // C: 점검자
    data.chapter     || '',                    // D: 챕터
    Number(data.accuracy) || 0,               // E: 일치율(%)
    data.grade       || '',                    // F: 등급
    serverTime,                                // G: 서버 수신 시각
  ]);

  // 등급별 행 색상
  const newRow = sheet.getLastRow();
  const color  = gradeColor(data.grade);
  sheet.getRange(newRow, 1, 1, 7).setBackground(color);

  // 일치율 셀 볼드
  sheet.getRange(newRow, 5).setFontWeight('bold');

  return newRow;
}

// ══════════════════════════════════════════════════
//  통계 시트 업데이트 (사람별 × 챕터별 평균)
// ══════════════════════════════════════════════════
function updateStats(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(CONFIG.STATS_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.STATS_SHEET);
    buildStatsHeader(sheet);
  }

  const name    = data.name    || '';
  const chapter = data.chapter || '';
  const acc     = Number(data.accuracy) || 0;

  // 기존 행 탐색
  const lastRow = sheet.getLastRow();
  let   found   = -1;

  if (lastRow > 1) {
    const nameCol  = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const chapCol  = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    for (let i = 0; i < nameCol.length; i++) {
      if (nameCol[i][0] === name && chapCol[i][0] === chapter) {
        found = i + 2; // 실제 행 번호 (헤더 제외)
        break;
      }
    }
  }

  if (found === -1) {
    // 새 항목 추가
    sheet.appendRow([name, chapter, acc, acc, acc, 1, getKST()]);
  } else {
    // 기존 항목 업데이트
    const row      = sheet.getRange(found, 1, 1, 7).getValues()[0];
    const prevMin  = Number(row[2]);
    const prevMax  = Number(row[3]);
    const prevAvg  = Number(row[4]);
    const prevCnt  = Number(row[5]);
    const newCnt   = prevCnt + 1;
    const newAvg   = Math.round((prevAvg * prevCnt + acc) / newCnt);

    sheet.getRange(found, 3).setValue(Math.min(prevMin, acc));  // C: 최저
    sheet.getRange(found, 4).setValue(Math.max(prevMax, acc));  // D: 최고
    sheet.getRange(found, 5).setValue(newAvg);                   // E: 평균
    sheet.getRange(found, 6).setValue(newCnt);                   // F: 횟수
    sheet.getRange(found, 7).setValue(getKST());                 // G: 최근 점검
  }
}

// ══════════════════════════════════════════════════
//  저점수 이메일 알림
// ══════════════════════════════════════════════════
function sendLowScoreAlert(data) {
  try {
    const subject = `[구두점검 알림] ${data.name} — ${data.accuracy}% (${data.grade})`;
    const body = [
      '점검 결과 알림',
      '',
      `점검자  : ${data.name}`,
      `챕터    : ${data.chapter}`,
      `일치율  : ${data.accuracy}%`,
      `등급    : ${data.grade}`,
      `점검시각: ${data.time}`,
      '',
      `기준(${CONFIG.ALERT_ON_LOW}%) 미만으로 알림이 발송되었습니다.`,
    ].join('\n');
    MailApp.sendEmail(CONFIG.ADMIN_EMAIL, subject, body);
  } catch (e) {
    console.error('메일 발송 실패:', e);
  }
}

// ══════════════════════════════════════════════════
//  요약 통계 반환 (?action=stats)
// ══════════════════════════════════════════════════
function getSummaryStats() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return { total: 0, data: [] };

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  const total = rows.length;
  const avgAcc = rows.reduce((s, r) => s + Number(r[4]), 0) / total;

  const gradeCount = { 완벽: 0, 우수: 0, 보통: 0, 미흡: 0 };
  rows.forEach(r => { if (gradeCount[r[5]] !== undefined) gradeCount[r[5]]++; });

  return {
    total,
    avgAccuracy : Math.round(avgAcc),
    gradeCount,
    lastUpdated : getKST(),
  };
}

// ══════════════════════════════════════════════════
//  헬퍼 함수들
// ══════════════════════════════════════════════════

function buildLogHeader(sheet) {
  const headers = ['번호', '점검 시각', '점검자', '챕터', '일치율(%)', '등급', '서버 수신'];
  sheet.appendRow(headers);

  const hRange = sheet.getRange(1, 1, 1, headers.length);
  hRange.setFontWeight('bold')
        .setBackground('#1F3864')
        .setFontColor('#FFFFFF')
        .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  // 열 너비
  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 220);
  sheet.setColumnWidth(5, 80);
  sheet.setColumnWidth(6, 70);
  sheet.setColumnWidth(7, 160);
}

function buildStatsHeader(sheet) {
  const headers = ['점검자', '챕터', '최저(%)', '최고(%)', '평균(%)', '횟수', '최근 점검'];
  sheet.appendRow(headers);

  const hRange = sheet.getRange(1, 1, 1, headers.length);
  hRange.setFontWeight('bold')
        .setBackground('#2E4057')
        .setFontColor('#FFFFFF')
        .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 80);
  sheet.setColumnWidth(6, 60);
  sheet.setColumnWidth(7, 160);
}

function gradeColor(grade) {
  const map = {
    '완벽': '#D9F7E8',  // 연초록
    '우수': '#DBEAFE',  // 연파랑
    '보통': '#FEF9C3',  // 연노랑
    '미흡': '#FEE2E2',  // 연빨강
  };
  return map[grade] || '#FFFFFF';
}

function getKST() {
  const now = new Date();
  const kst = new Date(now.getTime() + 9 * 60 * 60 * 1000);
  return Utilities.formatDate(kst, 'UTC', 'yyyy.MM.dd HH:mm:ss');
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
