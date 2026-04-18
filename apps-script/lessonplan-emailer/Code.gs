/**
 * 太平新光分校｜Lesson Plan 老師入口 Gmail 自動寄送
 *
 * 使用方式：
 * 1. Google Sheet → 擴充功能 → Apps Script
 * 2. 貼上本檔 Code.gs
 * 3. 執行 setupLessonPlanAutomation()
 * 4. 授權 Gmail / Spreadsheet 權限
 * 5. 回到「老師入口」分頁，選班級與週次後即可自動寄出備課包
 *
 * 注意：
 * - 請先到「教師Email」分頁填入每班老師 email
 * - 請先到「表單設定」填入 lesson_log_form_link
 * - 為避免重複寄送，同一班級同一週只會自動寄一次；若要重寄，可用上方選單「Lesson Plan 寄送工具」→「手動寄出目前選擇」
 */

const CONFIG = {
  ENTRY_SHEET: '老師入口',
  INDEX_SHEET: 'LessonPlan_Index',
  SETTINGS_SHEET: '表單設定',
  TEACHER_EMAIL_SHEET: '教師Email',
  SEND_LOG_SHEET: '寄送紀錄',
  CLASS_CELL: 'B4',
  WEEK_CELL: 'B5',
  COURSE_CELL: 'E4',
  UNIT_CELL: 'E5',
  LESSON_LINK_CELL: 'B7',
};

// 測試模式設定：啟用時所有信件都會改寄到 test_recipient_email，
// 原本教師Email 分頁中的收件人僅會保留在信件內容與寄送紀錄做為稽核用。
// 正式上線前請將 test_mode_enabled 改為 false，並先確認「教師Email」分頁資料正確。
const TEST_MODE_CONFIG = {
  test_mode_enabled: true,
  test_recipient_email: 'miyutang1980@gmail.com',
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Lesson Plan 寄送工具')
    .addItem('安裝自動寄送觸發器', 'setupLessonPlanAutomation')
    .addItem('手動寄出目前選擇', 'sendCurrentSelectionManual')
    .addItem('重設本班本週寄送紀錄', 'resetCurrentSelectionSendLog')
    .addToUi();
}

function setupLessonPlanAutomation() {
  const ss = SpreadsheetApp.getActive();
  ensureSendLogSheet_();
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'handleTeacherEntryEdit') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('handleTeacherEntryEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  SpreadsheetApp.getUi().alert('已安裝自動寄送觸發器。之後老師在「老師入口」選班級或週次後，系統會自動寄出本週備課包。');
}

function handleTeacherEntryEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getName() !== CONFIG.ENTRY_SHEET) return;
    const a1 = e.range.getA1Notation();
    if (![CONFIG.CLASS_CELL, CONFIG.WEEK_CELL].includes(a1)) return;
    Utilities.sleep(300);
    sendCurrentSelection_({ manual: false });
  } catch (err) {
    console.error(err);
  }
}

function sendCurrentSelectionManual() {
  sendCurrentSelection_({ manual: true });
}

function sendCurrentSelection_(options) {
  const ss = SpreadsheetApp.getActive();
  const entry = ss.getSheetByName(CONFIG.ENTRY_SHEET);
  const classId = String(entry.getRange(CONFIG.CLASS_CELL).getValue()).trim();
  const weekNo = Number(entry.getRange(CONFIG.WEEK_CELL).getValue());
  if (!classId || !weekNo) {
    if (options.manual) SpreadsheetApp.getUi().alert('請先選擇班級與週次。');
    return;
  }

  const data = getLessonPlanData_(classId, weekNo);
  if (!data) {
    if (options.manual) SpreadsheetApp.getUi().alert(`找不到 ${classId} Week ${weekNo} 的 Lesson Plan。`);
    return;
  }

  const teacher = getTeacherEmail_(classId);
  if (!teacher || !teacher.email) {
    if (options.manual) SpreadsheetApp.getUi().alert(`請先到「教師Email」分頁填入 ${classId} 對應老師 email。`);
    return;
  }

  const settings = getSettings_();
  const autoSendEnabled = String(settings.enable_auto_send || 'TRUE').toUpperCase() !== 'FALSE';
  if (!options.manual && !autoSendEnabled) return;

  const sendKey = `${classId}|${String(weekNo).padStart(2, '0')}`;
  if (!options.manual && hasAlreadySent_(sendKey)) return;

  const lessonLogLink = settings.lesson_log_form_link || '';
  const testScoreLink = settings.test_score_form_link || '';
  const observationLink = settings.observation_form_link || '';

  const intendedTeacherName = teacher.teacher_name || teacher.teacher || '老師';
  const intendedEmail = teacher.email;

  const subject = `太平新光分校｜${classId} Week ${String(weekNo).padStart(2, '0')} ${data.course} 備課包`;
  const htmlBody = buildEmailHtml_({
    teacherName: intendedTeacherName,
    classId,
    weekNo,
    course: data.course,
    unitRange: data.unit_range,
    lessonPlanLink: data.google_doc_link,
    lessonLogLink,
    testScoreLink,
    observationLink,
    intendedTeacherName,
    intendedEmail,
  });

  const sendResult = sendLessonPlanEmail_({
    intendedEmail,
    intendedTeacherName,
    subject,
    htmlBody,
  });

  appendSendLog_({
    sendKey,
    classId,
    weekNo,
    course: data.course,
    unitRange: data.unit_range,
    teacherName: intendedTeacherName,
    intendedEmail,
    actualEmail: sendResult.actualRecipient,
    testMode: sendResult.testMode,
    lessonPlanLink: data.google_doc_link,
    mode: options.manual ? 'Manual' : 'Auto',
  });

  if (options.manual) {
    const msg = sendResult.testMode
      ? `【測試模式】${classId} Week ${weekNo} 備課包原訂寄給 ${intendedEmail}，已改寄至 ${sendResult.actualRecipient}。`
      : `已寄出 ${classId} Week ${weekNo} 備課包給 ${sendResult.actualRecipient}`;
    SpreadsheetApp.getUi().alert(msg);
  }
}

/**
 * 統一寄送函式：自動寄送與手動寄送皆走此路徑。
 * 啟用測試模式時，實際 Gmail 收件人一律改為 TEST_MODE_CONFIG.test_recipient_email，
 * 不論「教師Email」分頁中設定的 email 為何。
 */
function sendLessonPlanEmail_(params) {
  const testMode = !!TEST_MODE_CONFIG.test_mode_enabled;
  const actualRecipient = testMode
    ? TEST_MODE_CONFIG.test_recipient_email
    : params.intendedEmail;

  const subject = testMode
    ? `[測試模式｜原收件人：${params.intendedEmail}] ${params.subject}`
    : params.subject;

  GmailApp.sendEmail(
    actualRecipient,
    subject,
    stripHtml_(params.htmlBody),
    {
      htmlBody: params.htmlBody,
      name: '太平新光分校 Lesson Plan 系統',
    }
  );

  console.log(
    `[LessonPlanEmail] testMode=${testMode} intended=${params.intendedTeacherName} <${params.intendedEmail}> actual=${actualRecipient}`
  );

  return { actualRecipient, testMode };
}

function getLessonPlanData_(classId, weekNo) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CONFIG.INDEX_SHEET);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const col = name => headers.indexOf(name);
  const key = `${classId}|${String(weekNo).padStart(2, '0')}`;
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowKey = row[col('lookup_key')] || `${row[col('class_id')]}|${String(row[col('week_no')]).padStart(2, '0')}`;
    if (rowKey === key) {
      return {
        class_id: row[col('class_id')],
        week_no: row[col('week_no')],
        course: row[col('course')],
        unit_range: row[col('unit_range')],
        file_name: row[col('file_name')],
        google_doc_link: row[col('google_doc_link')],
      };
    }
  }
  return null;
}

function getTeacherEmail_(classId) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CONFIG.TEACHER_EMAIL_SHEET);
  if (!sh) return null;
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const col = name => headers.indexOf(name);
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (String(row[col('class_id')]).trim() === classId) {
      return {
        class_id: row[col('class_id')],
        teacher_name: row[col('teacher_name')],
        email: row[col('email')],
      };
    }
  }
  return null;
}

function getSettings_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CONFIG.SETTINGS_SHEET);
  const values = sh.getDataRange().getValues();
  const out = {};
  for (let i = 1; i < values.length; i++) {
    out[String(values[i][0]).trim()] = values[i][1];
  }
  return out;
}

function buildEmailHtml_(data) {
  const lessonLog = data.lessonLogLink
    ? `<p><a href="${data.lessonLogLink}">填寫 Weekly Lesson Log</a></p>`
    : `<p style="color:#964219;">尚未設定 Lesson Log 表單連結，請到「表單設定」填入 lesson_log_form_link。</p>`;
  const testScore = data.testScoreLink
    ? `<p><a href="${data.testScoreLink}">測驗週填寫 Test Score Entry</a></p>`
    : '';
  const observation = data.observationLink
    ? `<p><a href="${data.observationLink}">主任看課填寫 Class Observation</a></p>`
    : '';

  const testModeBanner = TEST_MODE_CONFIG.test_mode_enabled
    ? `<div style="background:#FFF6D6; border:1px solid #E0B94A; padding:10px 14px; margin-bottom:14px; color:#7A5A00;">
         <strong>【測試模式】</strong>此信原訂收件人為 ${data.intendedTeacherName || ''} &lt;${data.intendedEmail || ''}&gt;，測試期間統一改寄至 ${TEST_MODE_CONFIG.test_recipient_email}。
       </div>`
    : '';

  return `
  <div style="font-family:Arial, sans-serif; color:#28251D; line-height:1.55;">
    ${testModeBanner}
    <h2 style="color:#01696F;">太平新光分校｜本週備課包</h2>
    <p>${data.teacherName}您好：</p>
    <p>以下是本週課程備課資訊，請上課前打開 Lesson Plan 檢查流程、mini check 與補救策略。</p>
    <table style="border-collapse:collapse; margin:12px 0;">
      <tr><td style="padding:6px 10px; background:#F7F6F2; font-weight:bold;">班級</td><td style="padding:6px 10px;">${data.classId}</td></tr>
      <tr><td style="padding:6px 10px; background:#F7F6F2; font-weight:bold;">週次</td><td style="padding:6px 10px;">Week ${String(data.weekNo).padStart(2, '0')}</td></tr>
      <tr><td style="padding:6px 10px; background:#F7F6F2; font-weight:bold;">課程</td><td style="padding:6px 10px;">${data.course}</td></tr>
      <tr><td style="padding:6px 10px; background:#F7F6F2; font-weight:bold;">單元</td><td style="padding:6px 10px;">${data.unitRange}</td></tr>
    </table>
    <p><a href="${data.lessonPlanLink}" style="font-size:16px; font-weight:bold; color:#01696F;">開啟本週 Lesson Plan</a></p>
    ${lessonLog}
    ${testScore}
    ${observation}
    <p style="margin-top:18px;">課後請務必填寫 Lesson Log，系統才能追蹤完成率、弱項學生與下週補救。</p>
    <p style="color:#7A7974; font-size:12px;">此信由太平新光分校 Lesson Plan 系統自動寄出。</p>
  </div>`;
}

function stripHtml_(html) {
  return html.replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

const SEND_LOG_HEADERS = [
  'timestamp', 'send_key', 'class_id', 'week_no', 'course', 'unit_range',
  'teacher_name', 'email', 'lesson_plan_link', 'mode',
  'intended_email', 'actual_email', 'test_mode',
];

function ensureSendLogSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CONFIG.SEND_LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SEND_LOG_SHEET);
    sh.appendRow(SEND_LOG_HEADERS);
    return sh;
  }
  // 若為舊版 schema（少了 intended/actual/test_mode 欄位），自動補齊欄位標題
  const firstRow = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0] || [];
  if (firstRow.length < SEND_LOG_HEADERS.length) {
    for (let i = firstRow.length; i < SEND_LOG_HEADERS.length; i++) {
      sh.getRange(1, i + 1).setValue(SEND_LOG_HEADERS[i]);
    }
  }
  return sh;
}

function appendSendLog_(data) {
  const sh = ensureSendLogSheet_();
  sh.appendRow([
    new Date(),
    data.sendKey,
    data.classId,
    data.weekNo,
    data.course,
    data.unitRange,
    data.teacherName,
    data.intendedEmail,         // 保留舊 email 欄位語意：老師原訂 email
    data.lessonPlanLink,
    data.mode,
    data.intendedEmail,
    data.actualEmail,
    data.testMode ? 'TRUE' : 'FALSE',
  ]);
}

function hasAlreadySent_(sendKey) {
  const sh = ensureSendLogSheet_();
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][1] === sendKey) return true;
  }
  return false;
}

function resetCurrentSelectionSendLog() {
  const ss = SpreadsheetApp.getActive();
  const entry = ss.getSheetByName(CONFIG.ENTRY_SHEET);
  const classId = String(entry.getRange(CONFIG.CLASS_CELL).getValue()).trim();
  const weekNo = Number(entry.getRange(CONFIG.WEEK_CELL).getValue());
  const sendKey = `${classId}|${String(weekNo).padStart(2, '0')}`;
  const sh = ensureSendLogSheet_();
  const values = sh.getDataRange().getValues();
  for (let i = values.length - 1; i >= 1; i--) {
    if (values[i][1] === sendKey) {
      sh.deleteRow(i + 1);
    }
  }
  SpreadsheetApp.getUi().alert(`已重設 ${sendKey} 的寄送紀錄，可重新寄出。`);
}
