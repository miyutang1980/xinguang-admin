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
 * - 為避免重複寄送，同一班級同一週只會自動寄一次；若要重寄，可用上方選單「老師備課包」→「寄出目前班級週次」
 * - 手動寄送（「寄出目前班級週次」）時，若老師入口只選了班級、週次留空，系統會自動從 Week_Plan
 *   的「目前週次」(數值等於 week_no 或 TRUE/Yes/是 等標記) /「time_status」/「date」欄
 *   推斷當週週次；若仍找不到才會跳警告。
 * - 若試算表開啟後看不到「老師備課包」選單，可在 Apps Script 編輯器中手動執行 forceCreateMenu() 後重新整理試算表。
 */

const CONFIG = {
  ENTRY_SHEET: '老師入口',
  INDEX_SHEET: 'LessonPlan_Index',
  SETTINGS_SHEET: '表單設定',
  TEACHER_EMAIL_SHEET: '教師Email',
  SEND_LOG_SHEET: '寄送紀錄',
  WEEK_PLAN_SHEET: 'Week_Plan',
  CLASS_CELL: 'B4',
  WEEK_CELL: 'B5',
  COURSE_CELL: 'E4',
  UNIT_CELL: 'E5',
  LESSON_LINK_CELL: 'B7',
};

const CURRENT_WEEK_TRUE_VALUES = ['true', 'yes', '是', 'current', '當週', '本週', '1', 'y', '✓', 'v'];
const THIS_WEEK_TIME_STATUS_VALUES = ['this week', '本週', '當週'];

// 測試模式設定：啟用時所有信件都會改寄到 test_recipient_email，
// 原本教師Email 分頁中的收件人僅會保留在信件內容與寄送紀錄做為稽核用。
// 正式上線前請將 test_mode_enabled 改為 false，並先確認「教師Email」分頁資料正確。
const TEST_MODE_CONFIG = {
  test_mode_enabled: true,
  test_recipient_email: 'miyutang1980@gmail.com',
};

const LESSON_PLAN_MENU_NAME = '老師備課包';

function onOpen(e) {
  buildLessonPlanMenu_();
}

/**
 * 手動在 Apps Script 編輯器中執行此函式，可強制重新建立「老師備課包」自訂選單。
 * 適用情境：試算表重新整理後選單未出現、onOpen 授權未完成、或剛貼上新版 Code.gs
 * 想立即看到選單時使用。執行前請先以目標試算表身份開啟 Apps Script，確保
 * SpreadsheetApp.getUi() 可取得對應試算表 UI。
 */
function forceCreateMenu() {
  buildLessonPlanMenu_();
  try {
    SpreadsheetApp.getUi().alert(`已重新建立「${LESSON_PLAN_MENU_NAME}」選單，請回到 Google Sheet 查看。`);
  } catch (err) {
    console.log('forceCreateMenu: 選單已建立，但無法顯示 alert（可能是從非試算表情境執行）。', err);
  }
}

function buildLessonPlanMenu_() {
  SpreadsheetApp.getUi()
    .createMenu(LESSON_PLAN_MENU_NAME)
    .addItem('寄出目前班級週次（週次空白時自動由 Week_Plan 目前週次推斷）', 'sendCurrentSelectionManual')
    .addItem('設定自動寄送觸發器', 'setupLessonPlanAutomation')
    .addItem('重設目前選擇寄送紀錄', 'resetCurrentSelectionSendLog')
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
  let classId = String(entry.getRange(CONFIG.CLASS_CELL).getValue()).trim();
  let weekNo = Number(entry.getRange(CONFIG.WEEK_CELL).getValue());

  // 自動偵測：只有手動寄送時才嘗試從 Week_Plan 推斷當週週次。
  // 自動（onEdit）路徑保持原本行為，避免觸發器意外亂寄。
  if (options.manual) {
    if (classId && !weekNo) {
      const detected = getCurrentWeekForClass_(classId);
      if (detected) {
        weekNo = detected.weekNo;
        console.log(`[LessonPlanEmail] auto-detected week ${weekNo} for class ${classId} via ${detected.source}`);
      }
    } else if (!classId && !weekNo) {
      const unique = findUniqueCurrentWeekAcrossClasses_();
      if (unique) {
        classId = unique.classId;
        weekNo = unique.weekNo;
        console.log(`[LessonPlanEmail] auto-detected class ${classId} week ${weekNo} via ${unique.source}`);
      } else {
        SpreadsheetApp.getUi().alert('請先在「老師入口」選擇班級（Week_Plan 內無法判斷唯一當週）。');
        return;
      }
    }
  }

  if (!classId || !weekNo) {
    if (options.manual) {
      const missing = !classId ? '班級' : '週次';
      SpreadsheetApp.getUi().alert(
        `請先選擇${missing}。（已嘗試從 Week_Plan「目前週次」/time_status/date 推斷但找不到符合的列）`
      );
    }
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

/**
 * 讀取 Week_Plan 分頁。若不存在則回傳 null，讓呼叫端自行處理（v3 以前的試算表沒有此分頁）。
 */
function getWeekPlanValues_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CONFIG.WEEK_PLAN_SHEET);
  if (!sh) return null;
  const values = sh.getDataRange().getValues();
  if (!values || values.length < 2) return null;
  const headers = values[0].map(h => String(h).trim());
  return { headers, rows: values.slice(1) };
}

/**
 * 在 Week_Plan 中為指定 classId 找出「目前週次」。
 * 判斷順序：
 *   1. 欄位「目前週次」為數值且等於該列的 week_no（v4 新公式：每列顯示同一個當週數字）
 *   2. 欄位「目前週次」為 TRUE / Yes / 是 / Current / 當週 / 1 等真值標記
 *   3. 欄位「time_status」為「This Week」/「本週」/「當週」
 *   4. 欄位「date」<= 今天 < date + 7 天
 * 若全部方式都找不到，回傳 null。
 */
function getCurrentWeekForClass_(classId) {
  const row = findCurrentWeekPlanRow_(classId);
  if (!row) return null;
  const weekNo = Number(row.weekNo);
  if (!weekNo) return null;
  return { weekNo, source: row.source };
}

function findCurrentWeekPlanRow_(classId) {
  const wp = getWeekPlanValues_();
  if (!wp) return null;
  const { headers, rows } = wp;
  const idx = name => headers.indexOf(name);
  const classCol = idx('class_id');
  const weekCol = idx('week_no');
  const currentWeekCol = idx('目前週次');
  const timeStatusCol = idx('time_status');
  const dateCol = idx('date');
  if (classCol < 0 || weekCol < 0) return null;

  const classRows = rows
    .map((r, i) => ({ r, i }))
    .filter(({ r }) => String(r[classCol]).trim() === String(classId).trim());
  if (classRows.length === 0) return null;

  // 1) 目前週次 欄位（v4 新增）
  //    優先判斷數值等於 week_no（新公式：整欄都填同一個當週數字），
  //    若無數值匹配則退回 TRUE/Yes/是 等標記值匹配。
  if (currentWeekCol >= 0) {
    for (const { r } of classRows) {
      if (currentWeekNumericMatches_(r[currentWeekCol], r[weekCol])) {
        return { weekNo: r[weekCol], source: '目前週次(numeric)' };
      }
    }
    for (const { r } of classRows) {
      if (isTruthyCurrentWeek_(r[currentWeekCol])) {
        return { weekNo: r[weekCol], source: '目前週次' };
      }
    }
  }

  // 2) time_status = This Week / 本週 / 當週
  if (timeStatusCol >= 0) {
    for (const { r } of classRows) {
      const ts = String(r[timeStatusCol] || '').trim().toLowerCase();
      if (THIS_WEEK_TIME_STATUS_VALUES.includes(ts)) {
        return { weekNo: r[weekCol], source: 'time_status' };
      }
    }
  }

  // 3) date 區間 (date <= today < date + 7)
  if (dateCol >= 0) {
    const today = stripTime_(new Date());
    for (const { r } of classRows) {
      const d = toDateOrNull_(r[dateCol]);
      if (!d) continue;
      const start = stripTime_(d);
      const end = new Date(start.getTime() + 7 * 24 * 60 * 60 * 1000);
      if (today.getTime() >= start.getTime() && today.getTime() < end.getTime()) {
        return { weekNo: r[weekCol], source: 'date_range' };
      }
    }
  }

  return null;
}

/**
 * 在所有 class 中尋找唯一的「當週」列。若多班同時有 current 列（常態），回傳 null。
 * 僅在 class 與 week 都空白時才會用到，用來處理使用者懶得選的情境。
 */
function findUniqueCurrentWeekAcrossClasses_() {
  const wp = getWeekPlanValues_();
  if (!wp) return null;
  const { headers, rows } = wp;
  const idx = name => headers.indexOf(name);
  const classCol = idx('class_id');
  const weekCol = idx('week_no');
  const currentWeekCol = idx('目前週次');
  const timeStatusCol = idx('time_status');
  const dateCol = idx('date');
  if (classCol < 0 || weekCol < 0) return null;

  const hits = [];
  const today = stripTime_(new Date());
  for (const r of rows) {
    const classId = String(r[classCol]).trim();
    const weekNo = r[weekCol];
    if (!classId || !weekNo) continue;
    let source = null;
    if (currentWeekCol >= 0 && currentWeekNumericMatches_(r[currentWeekCol], weekNo)) {
      source = '目前週次(numeric)';
    } else if (currentWeekCol >= 0 && isTruthyCurrentWeek_(r[currentWeekCol])) {
      source = '目前週次';
    } else if (timeStatusCol >= 0 && THIS_WEEK_TIME_STATUS_VALUES.includes(
      String(r[timeStatusCol] || '').trim().toLowerCase()
    )) {
      source = 'time_status';
    } else if (dateCol >= 0) {
      const d = toDateOrNull_(r[dateCol]);
      if (d) {
        const start = stripTime_(d);
        const end = new Date(start.getTime() + 7 * 24 * 60 * 60 * 1000);
        if (today.getTime() >= start.getTime() && today.getTime() < end.getTime()) {
          source = 'date_range';
        }
      }
    }
    if (source) hits.push({ classId, weekNo: Number(weekNo), source });
  }
  if (hits.length === 1) return hits[0];
  return null;
}

function isTruthyCurrentWeek_(v) {
  if (v === true) return true;
  if (v === 1) return true;
  if (v == null || v === '') return false;
  return CURRENT_WEEK_TRUE_VALUES.includes(String(v).trim().toLowerCase());
}

/**
 * v4 新公式：「目前週次」欄每列都填入同一個當週數字（例如全欄 = 3），
 * 僅有 week_no 等於該數字的列視為當週。支援數值與數字字串。
 * 布林 TRUE / 空值 / 非數字字串（如 Yes、當週）一律不算數值匹配，交給 isTruthyCurrentWeek_ 處理。
 */
function currentWeekNumericMatches_(currentWeekValue, weekNoValue) {
  if (currentWeekValue === true || currentWeekValue === false) return false;
  if (currentWeekValue == null || currentWeekValue === '') return false;
  const weekNo = Number(weekNoValue);
  if (!weekNo) return false;
  if (typeof currentWeekValue === 'number') {
    return currentWeekValue === weekNo;
  }
  const s = String(currentWeekValue).trim();
  if (s === '' || !/^-?\d+(\.\d+)?$/.test(s)) return false;
  return Number(s) === weekNo;
}

function toDateOrNull_(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

function stripTime_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
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
