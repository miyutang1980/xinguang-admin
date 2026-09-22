function _semesterAssignmentsList(adminUser, adminPass, semester) {
  const v = _verifyStaffOrTeacher(adminUser, adminPass);
  if (!v.ok) return { success:false, error:v.error };
  try {
    const all = _studentDataRows(SEMESTER_ASSIGN_SHEET, SEMESTER_ASSIGN_HEADERS).filter(function(r){ return r['學生編號']; });
    const current = _currentStudentSemester();
    const semesters = Array.from(new Set(all.map(function(r){ return r['學期']; }).filter(Boolean).concat([current]))).sort().reverse();
    const target = _studentText(semester) || current;
    if (!/^\d{3}-[12]$/.test(target)) throw new Error('學期格式須為 115-1');
    // W is a derived formula column, deliberately NOT added to the 22 writable
    // SEMESTER_ASSIGN_HEADERS. Existing updates must never overwrite its formula.
    const sh = _studentDataSheet(SEMESTER_ASSIGN_SHEET, SEMESTER_ASSIGN_HEADERS);
    let displayValues = null;
    if (sh && sh.getLastColumn() >= 23) {
      const column = sh.getRange(1, 23, Math.max(1, sh.getLastRow()), 1).getValues();
      if (String(column[0][0]) === '學生中英文姓名') displayValues = column;
    }
    const list = all.filter(function(r) { return r['學期'] === target; }).map(function(r) {
      if (target !== current || !displayValues) return r;
      const display = String((displayValues[r._row - 1] || [])[0] || '').trim();
      if (!display || /^#(REF!|ERROR!|N\/A|VALUE!|NAME\?)/.test(display)) {
        throw new Error('學期班級指派 W 欄姓名公式未完成，請核對後重新讀取');
      }
      return Object.assign({}, r, {'學生中英文姓名':display});
    });
    return {
      success:true,
      active:_studentModelActive(),
      currentSemester:current,
      semester:target,
      semesters:semesters,
      displayNameSource:target === current && displayValues ? '學期班級指派!W' : '',
      list:list
    };
  } catch(e) { return { success:false, error:e.message }; }
}

