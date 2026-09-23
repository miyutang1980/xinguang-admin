function _semesterAssignmentsList(adminUser, adminPass, semester) {
  const started = Date.now();
  const v = _verifyStaffOrTeacher(adminUser, adminPass);
  if (!v.ok) return { success:false, error:v.error };
  try {
    const current = _currentStudentSemester();
    const target = _studentText(semester) || current;
    if (!/^\d{3}-[12]$/.test(target)) throw new Error('學期格式須為 115-1');
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(SEMESTER_ASSIGN_SHEET);
    if (!sh) throw new Error('缺少學期班級指派工作表');
    // Read schema once, and current payload + derived W in a single operation.
    // Never format/write cells as a side effect of listing.
    const header = sh.getRange(1, 1, 1, 23).getValues()[0];
    if (SEMESTER_ASSIGN_HEADERS.some(function(h, i) { return header[i] !== h; })) {
      throw new Error('學期班級指派欄位結構不符');
    }
    const hasDisplay = header[22] === '學生中英文姓名';
    if (target === current && !hasDisplay) throw new Error('缺少 W 欄學生中英文姓名，請先核對欄位');
    let start = 2;
    if (target === current) {
      const name = 'XG_STUDENTS_' + current.replace('-', '_');
      const anchor = ss.getRangeByName(name);
      if (!anchor || anchor.getSheet().getSheetId() !== sh.getSheetId() || anchor.getRow() < 2 ||
          anchor.getColumn() !== 1 || anchor.getNumColumns() !== 23) {
        throw new Error('缺少或錯誤的當期名冊範圍 ' + name + '；已停止，不回退全表搜尋');
      }
      start = anchor.getRow();
    }
    const last = sh.getLastRow();
    const values = last >= start ? sh.getRange(start, 1, last-start+1, 23).getValues() : [];
    // Historical choices are metadata. Do not scan historical student rows merely
    // to build a dropdown when serving the current semester.
    const configured = PropertiesService.getScriptProperties().getProperty('STUDENT_SEMESTERS');
    const semesters = Array.from(new Set(String(configured || '114-2').split(',')
      .map(function(s) { return s.trim(); }).filter(function(s) { return /^\d{3}-[12]$/.test(s); })
      .concat([current,target]))).sort().reverse();
    const list = [];
    const seen = {};
    values.forEach(function(row, i) {
      if (!row[2]) return;
      const term = String(row[1] || '').trim();
      if (target === current && term !== current) {
        throw new Error('當期命名範圍混入其他學期，請核對範圍；未更動資料');
      }
      if (term !== target) return;
      const r = {_row:start+i};
      SEMESTER_ASSIGN_HEADERS.forEach(function(h,j) { r[h] = String(row[j] == null ? '' : row[j]); });
      if (!r['學生姓名'] || r['指派編號'] !== target+'|'+r['學生編號'] || seen[r['學生編號']]) {
        throw new Error('學期指派編號、姓名或重複學生資料異常');
      }
      seen[r['學生編號']] = true;
      if (target === current) {
        const display = String(row[22] || '').trim();
        if (!display || display.charAt(0) === '#') throw new Error('學期班級指派 W 欄姓名公式未完成');
        r['學生中英文姓名'] = display;
      }
      list.push(r);
    });
    return {
      success:true,
      active:_studentModelActive(),
      currentSemester:current,
      semester:target,
      semesters:semesters,
      displayNameSource:target === current ? '學期班級指派!W' : '',
      scoped:target === current,
      readRows:values.length,
      readStartRow:start,
      elapsedMs:Date.now()-started,
      readerVersion:'semester-scoped-v2',
      list:list
    };
  } catch(e) { return { success:false, error:e.message }; }
}
