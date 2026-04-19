/**
 * 太平新光分校｜全校版 Google Forms 表單回應自動分流系統
 *
 * 用途：
 * - Google Forms 回應自動分流到 Lesson_Log / Quiz_Result / Parent_Feedback_Biweekly
 * - 讀取 Class_Master / Teacher_Master / Student_Master 做主檔校驗與 ID 補齊
 * - 保存 Raw Response、分流紀錄、錯誤紀錄
 *
 * 安裝：
 * 1. 將本檔完整貼到全校版 Google 試算表：擴充功能 → Apps Script → Code.gs
 * 2. 執行 setupFormRouter()
 * 3. 授權 Spreadsheet 權限
 * 4. Google Forms 回應表請連到同一本 Google Sheet
 */

const SCHOOL_ROUTER_CONFIG = {
  school_name: '太平新光分校',
  timezone: 'Asia/Taipei',

  master_sheets: {
    class: 'Class_Master',
    teacher: 'Teacher_Master',
    student: 'Student_Master',
  },

  target_sheets: {
    lesson_log: 'Lesson_Log',
    quiz_result: 'Quiz_Result',
    parent_feedback: 'Parent_Feedback_Biweekly',
  },

  log_sheets: {
    raw: 'Form_Response_Raw',
    routing: 'Form_Routing_Log',
    error: 'Form_Error_Log',
  },

  // 表單可用這些值指定分流；若沒有 form_type，系統會用欄位特徵判斷。
  route_type_values: {
    lesson_log: ['lesson_log', 'lesson log', '上課紀錄', '課堂紀錄', 'lesson'],
    quiz_result: ['quiz_result', 'quiz result', '小考成績', '測驗成績', 'quiz'],
    parent_feedback: ['parent_feedback', 'parent feedback', '家長回饋', '聯絡簿', 'contactbook'],
  },
};

const TARGET_HEADERS = {
  Lesson_Log: [
    'log_id',
    'date',
    'teacher_id',
    'class_id',
    'week_no',
    'plan_id',
    'units_taught',
    'completion_rate',
    'student_engagement',
    'homework_assigned',
    'issues',
    'next_action',
    'submitted_at',
  ],
  Quiz_Result: [
    'result_id',
    'quiz_id',
    'student_id',
    'class_id',
    'test_date',
    'listening_score',
    'speaking_score',
    'reading_score',
    'writing_score',
    'grammar_score',
    'vocabulary_score',
    'total_score',
    'mastery_level',
    'error_tags',
    'remediation_needed',
    'notes',
  ],
  Parent_Feedback_Biweekly: [
    'feedback_id',
    'student_id',
    'class_id',
    'period_start',
    'period_end',
    'week_range',
    'units_covered',
    'learning_highlights',
    'skill_scores_summary',
    'vocabulary_status',
    'grammar_status',
    'homework_status',
    'notebook_status',
    'risk_area',
    'next_2week_focus',
    'parent_action',
    'generated_message',
    'review_status',
    'sent_status',
    'sent_at',
  ],
  Form_Response_Raw: [
    'timestamp',
    'source_sheet',
    'detected_route',
    'raw_json',
  ],
  Form_Routing_Log: [
    'timestamp',
    'route_type',
    'source_sheet',
    'target_sheet',
    'target_row',
    'record_id',
    'class_id',
    'student_id',
    'status',
    'warnings',
    'raw_json',
  ],
  Form_Error_Log: [
    'timestamp',
    'source_sheet',
    'route_type',
    'error_message',
    'stack',
    'raw_json',
  ],
};

const FIELD_ALIASES = {
  form_type: ['form_type', 'form type', '表單類型', '表單種類', '分流類型'],

  date: ['date', '上課日期', '日期', '授課日期', 'timestamp', '時間戳記'],
  submitted_at: ['submitted_at', '提交時間', '送出時間', 'timestamp', '時間戳記'],

  class_id: ['class_id', 'class id', '班級ID', '班級代碼', '班級', 'class', '班別'],
  class_name: ['class_name', '班級名稱'],

  teacher_id: ['teacher_id', 'teacher id', '老師ID', '教師ID', '老師代碼'],
  teacher_name: ['teacher_name', 'teacher', '老師姓名', '教師姓名', '老師'],
  teacher_email: ['teacher_email', '老師email', '老師信箱', '教師email', '教師信箱', 'email'],

  student_id: ['student_id', 'student id', '學生ID', '學生代碼', '學號'],
  student_name: ['student_name', 'student', '學生姓名', '中文姓名', '姓名'],
  english_name: ['english_name', '英文名', '英文名字'],

  week_no: ['week_no', 'week no', '週次', '第幾週', 'week', '週數'],
  plan_id: ['plan_id', 'lesson_plan_id', '課程計畫ID', '備課ID'],

  units_taught: ['units_taught', '完成單元', '授課單元', '本週單元', 'unit', 'units'],
  completion_rate: ['completion_rate', '完成率', '課程完成率', '進度完成率'],
  student_engagement: ['student_engagement', '學生參與度', '課堂參與', '參與度'],
  homework_assigned: ['homework_assigned', '作業', '指派作業', '本週作業'],
  issues: ['issues', '問題', '課堂問題', '狀況', '困難'],
  next_action: ['next_action', '下次行動', '下週重點', '後續處理'],

  quiz_id: ['quiz_id', 'quiz id', '小考ID', '測驗ID', 'assessment_id', '考試代碼'],
  test_date: ['test_date', '測驗日期', '考試日期', '小考日期', 'date', '日期'],
  listening_score: ['listening_score', 'listening', '聽力', '聽力分數'],
  speaking_score: ['speaking_score', 'speaking', '口說', '口說分數'],
  reading_score: ['reading_score', 'reading', '閱讀', '閱讀分數'],
  writing_score: ['writing_score', 'writing', '寫作', '寫作分數'],
  grammar_score: ['grammar_score', 'grammar', '文法', '文法分數'],
  vocabulary_score: ['vocabulary_score', 'vocabulary', '單字', '單字分數'],
  total_score: ['total_score', 'total', '總分', '合計', '分數'],
  mastery_level: ['mastery_level', '掌握程度', '學習狀態'],
  error_tags: ['error_tags', '錯誤類型', '錯題標籤', '錯誤標籤'],
  remediation_needed: ['remediation_needed', '需要補救', '是否補救', '補救教學'],

  period_start: ['period_start', '區間開始', '回饋開始日', '開始日期'],
  period_end: ['period_end', '區間結束', '回饋結束日', '結束日期'],
  week_range: ['week_range', '週次範圍', '回饋週次'],
  units_covered: ['units_covered', '本期單元', '完成單元', '學習單元'],
  learning_highlights: ['learning_highlights', '學習亮點', '優點', '表現亮點'],
  skill_scores_summary: ['skill_scores_summary', '能力摘要', '聽說讀寫摘要', '成績摘要'],
  vocabulary_status: ['vocabulary_status', '單字狀況', '單字'],
  grammar_status: ['grammar_status', '文法狀況', '文法'],
  homework_status: ['homework_status', '作業狀況', '作業'],
  notebook_status: ['notebook_status', '簿本狀況', '簿本', '簿本檢查'],
  risk_area: ['risk_area', '需加強', '風險項目', '弱點'],
  next_2week_focus: ['next_2week_focus', '下兩週重點', '下階段重點'],
  parent_action: ['parent_action', '家長配合', '家長可協助', '家庭任務'],
  generated_message: ['generated_message', '聯絡簿內容', '家長回饋內容', '回饋訊息'],
  review_status: ['review_status', '審核狀態'],
  sent_status: ['sent_status', '寄送狀態'],

  notes: ['notes', '備註', '其他說明'],
};

function onOpen(e) {
  createSchoolRouterMenu_();
}

function createSchoolRouterMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('全校表單分流')
    .addItem('安裝表單分流觸發器', 'setupFormRouter')
    .addItem('測試上課紀錄分流', 'testRouteLessonLog')
    .addItem('測試小考成績分流', 'testRouteQuizResult')
    .addItem('測試家長回饋分流', 'testRouteParentFeedback')
    .addSeparator()
    .addItem('建立缺少的系統分頁', 'ensureRouterSheets')
    .addToUi();
}

function forceCreateSchoolRouterMenu() {
  createSchoolRouterMenu_();
  try {
    SpreadsheetApp.getUi().alert('已建立「全校表單分流」選單。請重新整理 Google Sheet 後確認。');
  } catch (err) {
    console.log('Menu created; alert skipped: ' + err.message);
  }
}

function setupFormRouter() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureRouterSheets();

  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'handleFormSubmit') {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  ScriptApp.newTrigger('handleFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  SpreadsheetApp.getUi().alert(
    '表單分流觸發器已安裝完成。\n\n請確認 Google Forms 回應都連到這一本 Google Sheet。'
  );
}

function ensureRouterSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allHeaders = Object.assign({}, TARGET_HEADERS);
  for (const sheetName in allHeaders) {
    ensureSheetWithHeaders_(ss, sheetName, allHeaders[sheetName]);
  }
}

/**
 * Google Forms installable onFormSubmit trigger entry.
 */
function handleFormSubmit(e) {
  routeFormSubmit_(e);
}

/**
 * Optional direct trigger entry if user manually installs this function.
 */
function onFormSubmit(e) {
  routeFormSubmit_(e);
}

function routeFormSubmit_(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const raw = extractRawResponse_(e);
  let routeType = '';
  let sourceSheet = raw.source_sheet || '';

  try {
    ensureRouterSheets();

    const data = normalizeNamedValues_(raw.named_values);
    routeType = detectRouteType_(data, sourceSheet);
    appendRawLog_(raw, routeType);

    let result;
    if (routeType === 'lesson_log') {
      result = processLessonLog_(ss, data, raw);
    } else if (routeType === 'quiz_result') {
      result = processQuizResult_(ss, data, raw);
    } else if (routeType === 'parent_feedback') {
      result = processParentFeedback_(ss, data, raw);
    } else {
      throw new Error(
        '無法判斷表單類型。請在表單加入題目 form_type，值填 lesson_log / quiz_result / parent_feedback。'
      );
    }

    appendRoutingLog_({
      route_type: routeType,
      source_sheet: sourceSheet,
      target_sheet: result.target_sheet,
      target_row: result.target_row,
      record_id: result.record_id,
      class_id: result.class_id || '',
      student_id: result.student_id || '',
      status: 'Success',
      warnings: (result.warnings || []).join('；'),
      raw_json: JSON.stringify(raw.named_values),
    });
  } catch (err) {
    appendErrorLog_({
      source_sheet: sourceSheet,
      route_type: routeType,
      error_message: err.message,
      stack: err.stack || '',
      raw_json: JSON.stringify(raw.named_values || {}),
    });
    throw err;
  }
}

function processLessonLog_(ss, data, raw) {
  const warnings = [];
  const classId = resolveClassId_(ss, getAlias_(data, 'class_id'), getAlias_(data, 'class_name'), warnings);
  const teacherId = resolveTeacherId_(
    ss,
    getAlias_(data, 'teacher_id'),
    getAlias_(data, 'teacher_name'),
    getAlias_(data, 'teacher_email'),
    warnings
  );

  requireValue_(classId, 'Lesson_Log 缺少 class_id / 班級');

  const record = {
    log_id: getAlias_(data, 'log_id') || makeId_('LOG'),
    date: getAlias_(data, 'date') || raw.timestamp,
    teacher_id: teacherId,
    class_id: classId,
    week_no: normalizeWeekNo_(getAlias_(data, 'week_no')),
    plan_id: getAlias_(data, 'plan_id'),
    units_taught: getAlias_(data, 'units_taught'),
    completion_rate: normalizePercent_(getAlias_(data, 'completion_rate')),
    student_engagement: getAlias_(data, 'student_engagement'),
    homework_assigned: getAlias_(data, 'homework_assigned'),
    issues: getAlias_(data, 'issues'),
    next_action: getAlias_(data, 'next_action'),
    submitted_at: raw.timestamp,
  };

  const row = appendObjectRow_(ss, SCHOOL_ROUTER_CONFIG.target_sheets.lesson_log, record);
  return {
    target_sheet: SCHOOL_ROUTER_CONFIG.target_sheets.lesson_log,
    target_row: row,
    record_id: record.log_id,
    class_id: record.class_id,
    student_id: '',
    warnings,
  };
}

function processQuizResult_(ss, data, raw) {
  const warnings = [];
  const classId = resolveClassId_(ss, getAlias_(data, 'class_id'), getAlias_(data, 'class_name'), warnings);
  const studentId = resolveStudentId_(
    ss,
    getAlias_(data, 'student_id'),
    getAlias_(data, 'student_name'),
    getAlias_(data, 'english_name'),
    classId,
    warnings
  );

  requireValue_(classId, 'Quiz_Result 缺少 class_id / 班級');
  requireValue_(studentId, 'Quiz_Result 缺少 student_id / 學生');

  const scores = {
    listening_score: normalizeNumber_(getAlias_(data, 'listening_score')),
    speaking_score: normalizeNumber_(getAlias_(data, 'speaking_score')),
    reading_score: normalizeNumber_(getAlias_(data, 'reading_score')),
    writing_score: normalizeNumber_(getAlias_(data, 'writing_score')),
    grammar_score: normalizeNumber_(getAlias_(data, 'grammar_score')),
    vocabulary_score: normalizeNumber_(getAlias_(data, 'vocabulary_score')),
  };
  const providedTotal = normalizeNumber_(getAlias_(data, 'total_score'));
  const totalScore = providedTotal !== '' ? providedTotal : sumScores_(scores);
  const mastery = getAlias_(data, 'mastery_level') || inferMasteryLevel_(totalScore);
  const remediation = getAlias_(data, 'remediation_needed') || (mastery === 'At Risk' ? 'Yes' : 'No');

  const record = {
    result_id: getAlias_(data, 'result_id') || makeId_('QR'),
    quiz_id: getAlias_(data, 'quiz_id') || makeId_('QUIZ_UNKNOWN'),
    student_id: studentId,
    class_id: classId,
    test_date: getAlias_(data, 'test_date') || raw.timestamp,
    listening_score: scores.listening_score,
    speaking_score: scores.speaking_score,
    reading_score: scores.reading_score,
    writing_score: scores.writing_score,
    grammar_score: scores.grammar_score,
    vocabulary_score: scores.vocabulary_score,
    total_score: totalScore,
    mastery_level: mastery,
    error_tags: getAlias_(data, 'error_tags'),
    remediation_needed: remediation,
    notes: getAlias_(data, 'notes'),
  };

  const row = appendObjectRow_(ss, SCHOOL_ROUTER_CONFIG.target_sheets.quiz_result, record);
  return {
    target_sheet: SCHOOL_ROUTER_CONFIG.target_sheets.quiz_result,
    target_row: row,
    record_id: record.result_id,
    class_id: record.class_id,
    student_id: record.student_id,
    warnings,
  };
}

function processParentFeedback_(ss, data, raw) {
  const warnings = [];
  const classId = resolveClassId_(ss, getAlias_(data, 'class_id'), getAlias_(data, 'class_name'), warnings);
  const studentId = resolveStudentId_(
    ss,
    getAlias_(data, 'student_id'),
    getAlias_(data, 'student_name'),
    getAlias_(data, 'english_name'),
    classId,
    warnings
  );

  requireValue_(classId, 'Parent_Feedback_Biweekly 缺少 class_id / 班級');
  requireValue_(studentId, 'Parent_Feedback_Biweekly 缺少 student_id / 學生');

  const record = {
    feedback_id: getAlias_(data, 'feedback_id') || makeId_('PF'),
    student_id: studentId,
    class_id: classId,
    period_start: getAlias_(data, 'period_start'),
    period_end: getAlias_(data, 'period_end'),
    week_range: getAlias_(data, 'week_range'),
    units_covered: getAlias_(data, 'units_covered'),
    learning_highlights: getAlias_(data, 'learning_highlights'),
    skill_scores_summary: getAlias_(data, 'skill_scores_summary'),
    vocabulary_status: getAlias_(data, 'vocabulary_status'),
    grammar_status: getAlias_(data, 'grammar_status'),
    homework_status: getAlias_(data, 'homework_status'),
    notebook_status: getAlias_(data, 'notebook_status'),
    risk_area: getAlias_(data, 'risk_area'),
    next_2week_focus: getAlias_(data, 'next_2week_focus'),
    parent_action: getAlias_(data, 'parent_action'),
    generated_message: getAlias_(data, 'generated_message'),
    review_status: getAlias_(data, 'review_status') || 'Draft',
    sent_status: getAlias_(data, 'sent_status') || 'No',
    sent_at: '',
  };

  const row = appendObjectRow_(ss, SCHOOL_ROUTER_CONFIG.target_sheets.parent_feedback, record);
  return {
    target_sheet: SCHOOL_ROUTER_CONFIG.target_sheets.parent_feedback,
    target_row: row,
    record_id: record.feedback_id,
    class_id: record.class_id,
    student_id: record.student_id,
    warnings,
  };
}

function detectRouteType_(data, sourceSheet) {
  const explicit = String(getAlias_(data, 'form_type') || '').trim().toLowerCase();
  for (const route in SCHOOL_ROUTER_CONFIG.route_type_values) {
    if (SCHOOL_ROUTER_CONFIG.route_type_values[route].indexOf(explicit) >= 0) return route;
  }

  const source = String(sourceSheet || '').toLowerCase();
  if (source.indexOf('lesson') >= 0 || source.indexOf('上課') >= 0) return 'lesson_log';
  if (source.indexOf('quiz') >= 0 || source.indexOf('小考') >= 0 || source.indexOf('測驗') >= 0) return 'quiz_result';
  if (source.indexOf('parent') >= 0 || source.indexOf('家長') >= 0 || source.indexOf('聯絡簿') >= 0) return 'parent_feedback';

  if (hasAnyAlias_(data, ['units_taught', 'completion_rate', 'student_engagement', 'homework_assigned'])) {
    return 'lesson_log';
  }
  if (hasAnyAlias_(data, ['quiz_id', 'listening_score', 'speaking_score', 'reading_score', 'writing_score', 'total_score'])) {
    return 'quiz_result';
  }
  if (hasAnyAlias_(data, ['period_start', 'period_end', 'learning_highlights', 'generated_message', 'parent_action'])) {
    return 'parent_feedback';
  }
  return '';
}

function extractRawResponse_(e) {
  const timestamp = Utilities.formatDate(new Date(), SCHOOL_ROUTER_CONFIG.timezone, 'yyyy-MM-dd HH:mm:ss');
  const sourceSheet = e && e.range && e.range.getSheet ? e.range.getSheet().getName() : '';
  const namedValues = {};

  if (e && e.namedValues) {
    for (const key in e.namedValues) {
      const value = e.namedValues[key];
      namedValues[key] = Array.isArray(value) ? value.join(', ') : value;
    }
  }

  return {
    timestamp,
    source_sheet: sourceSheet,
    named_values: namedValues,
  };
}

function normalizeNamedValues_(namedValues) {
  const data = {};
  for (const key in namedValues || {}) {
    data[normalizeKey_(key)] = {
      original_key: key,
      value: namedValues[key],
    };
  }
  return data;
}

function getAlias_(data, canonicalKey) {
  const aliases = FIELD_ALIASES[canonicalKey] || [canonicalKey];
  for (const alias of aliases) {
    const item = data[normalizeKey_(alias)];
    if (item && item.value !== undefined && item.value !== null) {
      return String(item.value).trim();
    }
  }
  return '';
}

function hasAnyAlias_(data, canonicalKeys) {
  for (const key of canonicalKeys) {
    const value = getAlias_(data, key);
    if (value !== '') return true;
  }
  return false;
}

function normalizeKey_(key) {
  return String(key || '')
    .toLowerCase()
    .replace(/[\s_：:()（）\-\[\]【】]/g, '');
}

function resolveClassId_(ss, classInput, classNameInput, warnings) {
  const sheet = ss.getSheetByName(SCHOOL_ROUTER_CONFIG.master_sheets.class);
  if (!sheet) {
    warnings.push('找不到 Class_Master，略過班級主檔校驗');
    return classInput || classNameInput || '';
  }
  const input = classInput || classNameInput || '';
  if (!input) return '';
  const match = findMasterRow_(sheet, ['class_id', 'class_name'], input);
  if (!match) {
    warnings.push('Class_Master 找不到班級：' + input);
    return input;
  }
  return match.class_id || input;
}

function resolveTeacherId_(ss, teacherIdInput, teacherNameInput, teacherEmailInput, warnings) {
  const sheet = ss.getSheetByName(SCHOOL_ROUTER_CONFIG.master_sheets.teacher);
  if (!sheet) {
    warnings.push('找不到 Teacher_Master，略過教師主檔校驗');
    return teacherIdInput || '';
  }
  const input = teacherIdInput || teacherEmailInput || teacherNameInput || '';
  if (!input) return '';
  const match = findMasterRow_(sheet, ['teacher_id', 'email', 'teacher_name'], input);
  if (!match) {
    warnings.push('Teacher_Master 找不到老師：' + input);
    return teacherIdInput || teacherNameInput || teacherEmailInput || '';
  }
  return match.teacher_id || input;
}

function resolveStudentId_(ss, studentIdInput, studentNameInput, englishNameInput, classId, warnings) {
  const sheet = ss.getSheetByName(SCHOOL_ROUTER_CONFIG.master_sheets.student);
  if (!sheet) {
    warnings.push('找不到 Student_Master，略過學生主檔校驗');
    return studentIdInput || '';
  }
  const input = studentIdInput || studentNameInput || englishNameInput || '';
  if (!input) return '';
  const matches = findMasterRows_(sheet, ['student_id', 'student_name', 'english_name'], input);
  if (matches.length === 0) {
    warnings.push('Student_Master 找不到學生：' + input);
    return input;
  }
  if (classId) {
    const sameClass = matches.filter((row) => String(row.class_id || '') === String(classId));
    if (sameClass.length === 1) return sameClass[0].student_id || input;
    if (sameClass.length > 1) {
      warnings.push('Student_Master 同班級找到多位同名學生：' + input);
      return sameClass[0].student_id || input;
    }
  }
  if (matches.length > 1) warnings.push('Student_Master 找到多位同名學生，使用第一筆：' + input);
  return matches[0].student_id || input;
}

function findMasterRow_(sheet, candidateHeaders, input) {
  const rows = findMasterRows_(sheet, candidateHeaders, input);
  return rows.length ? rows[0] : null;
}

function findMasterRows_(sheet, candidateHeaders, input) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map((h) => String(h || '').trim());
  const normalizedHeaders = headers.map(normalizeKey_);
  const candidateIndexes = candidateHeaders
    .map((h) => normalizedHeaders.indexOf(normalizeKey_(h)))
    .filter((idx) => idx >= 0);
  const normalizedInput = normalizeComparable_(input);
  const results = [];
  for (let r = 1; r < values.length; r++) {
    for (const idx of candidateIndexes) {
      if (normalizeComparable_(values[r][idx]) === normalizedInput) {
        const obj = {};
        headers.forEach((h, i) => {
          obj[h] = values[r][i];
        });
        results.push(obj);
        break;
      }
    }
  }
  return results;
}

function normalizeComparable_(value) {
  return String(value || '').trim().toLowerCase();
}

function appendObjectRow_(ss, sheetName, record) {
  const sheet = ensureSheetWithHeaders_(ss, sheetName, TARGET_HEADERS[sheetName]);
  const headers = getHeaders_(sheet);
  const row = headers.map((header) => record[header] !== undefined ? record[header] : '');
  sheet.appendRow(row);
  return sheet.getLastRow();
}

function appendRawLog_(raw, routeType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  appendObjectRow_(ss, SCHOOL_ROUTER_CONFIG.log_sheets.raw, {
    timestamp: raw.timestamp,
    source_sheet: raw.source_sheet,
    detected_route: routeType,
    raw_json: JSON.stringify(raw.named_values),
  });
}

function appendRoutingLog_(record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  record.timestamp = Utilities.formatDate(new Date(), SCHOOL_ROUTER_CONFIG.timezone, 'yyyy-MM-dd HH:mm:ss');
  appendObjectRow_(ss, SCHOOL_ROUTER_CONFIG.log_sheets.routing, record);
}

function appendErrorLog_(record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  record.timestamp = Utilities.formatDate(new Date(), SCHOOL_ROUTER_CONFIG.timezone, 'yyyy-MM-dd HH:mm:ss');
  appendObjectRow_(ss, SCHOOL_ROUTER_CONFIG.log_sheets.error, record);
}

function ensureSheetWithHeaders_(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  const currentHeaders = getHeaders_(sheet);
  if (currentHeaders.length === 0 || currentHeaders.every((h) => h === '')) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    formatHeaderRow_(sheet, headers.length);
    return sheet;
  }

  const missing = headers.filter((h) => currentHeaders.indexOf(h) < 0);
  if (missing.length) {
    const startCol = currentHeaders.length + 1;
    sheet.getRange(1, startCol, 1, missing.length).setValues([missing]);
    formatHeaderRow_(sheet, startCol + missing.length - 1);
  }
  return sheet;
}

function getHeaders_(sheet) {
  if (sheet.getLastColumn() === 0) return [];
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map((h) => String(h || '').trim());
}

function formatHeaderRow_(sheet, lastCol) {
  sheet.getRange(1, 1, 1, lastCol)
    .setBackground('#1B474D')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.autoResizeColumns(1, Math.min(lastCol, 12));
}

function requireValue_(value, message) {
  if (value === undefined || value === null || String(value).trim() === '') {
    throw new Error(message);
  }
}

function makeId_(prefix) {
  const stamp = Utilities.formatDate(new Date(), SCHOOL_ROUTER_CONFIG.timezone, 'yyyyMMddHHmmss');
  const rand = Math.floor(Math.random() * 9000 + 1000);
  return prefix + '_' + stamp + '_' + rand;
}

function normalizeWeekNo_(value) {
  if (value === '') return '';
  const match = String(value).match(/\d+/);
  return match ? Number(match[0]) : value;
}

function normalizeNumber_(value) {
  if (value === '' || value === null || value === undefined) return '';
  const s = String(value).replace('%', '').trim();
  const n = Number(s);
  return isNaN(n) ? value : n;
}

function normalizePercent_(value) {
  if (value === '' || value === null || value === undefined) return '';
  const s = String(value).trim();
  if (s.indexOf('%') >= 0) {
    const n = Number(s.replace('%', '').trim());
    return isNaN(n) ? value : n / 100;
  }
  const n = Number(s);
  if (isNaN(n)) return value;
  return n > 1 ? n / 100 : n;
}

function sumScores_(scores) {
  let total = 0;
  let hasValue = false;
  for (const key in scores) {
    if (scores[key] !== '' && !isNaN(Number(scores[key]))) {
      total += Number(scores[key]);
      hasValue = true;
    }
  }
  return hasValue ? total : '';
}

function inferMasteryLevel_(totalScore) {
  if (totalScore === '' || isNaN(Number(totalScore))) return '';
  const score = Number(totalScore);
  if (score >= 85) return 'Mastered';
  if (score >= 70) return 'Developing';
  return 'At Risk';
}

function testRouteLessonLog() {
  const fake = {
    namedValues: {
      form_type: ['lesson_log'],
      class_id: ['Elite_1B_2026S1'],
      teacher_id: ['T001'],
      week_no: ['1'],
      units_taught: ['Unit 1'],
      completion_rate: ['90%'],
      student_engagement: ['Good'],
      homework_assigned: ['Workbook p.1-2'],
    },
  };
  routeFormSubmit_(fake);
  SpreadsheetApp.getUi().alert('測試 Lesson_Log 分流完成，請查看 Lesson_Log 與 Form_Routing_Log。');
}

function testRouteQuizResult() {
  const fake = {
    namedValues: {
      form_type: ['quiz_result'],
      quiz_id: ['QUIZ_Elite_1B_W05'],
      class_id: ['Elite_1B_2026S1'],
      student_id: ['S0001'],
      listening_score: ['20'],
      speaking_score: ['18'],
      reading_score: ['22'],
      writing_score: ['20'],
      grammar_score: ['10'],
      vocabulary_score: ['8'],
      error_tags: ['Listening keyword'],
    },
  };
  routeFormSubmit_(fake);
  SpreadsheetApp.getUi().alert('測試 Quiz_Result 分流完成，請查看 Quiz_Result 與 Form_Routing_Log。');
}

function testRouteParentFeedback() {
  const fake = {
    namedValues: {
      form_type: ['parent_feedback'],
      class_id: ['Elite_1B_2026S1'],
      student_id: ['S0001'],
      week_range: ['Week 1-2'],
      units_covered: ['Unit 1'],
      learning_highlights: ['課堂參與穩定，能主動回答問題。'],
      skill_scores_summary: ['聽力穩定，閱讀需加強細節理解。'],
      parent_action: ['每天朗讀 5 分鐘。'],
      generated_message: ['親愛的家長您好，本兩週孩子完成 Unit 1...'],
    },
  };
  routeFormSubmit_(fake);
  SpreadsheetApp.getUi().alert('測試 Parent_Feedback_Biweekly 分流完成，請查看 Parent_Feedback_Biweekly 與 Form_Routing_Log。');
}
