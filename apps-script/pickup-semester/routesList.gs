// Replace the existing _routesList function; do not leave two definitions.
// Named range XG_PICKUP_115_1 is a stable row anchor on the original sheet.
// Historical rows are retained. No row is moved, deleted, or renumbered here.
function _routesList(adminUser, adminPass) {
  const auth = _verifyPickupUser(adminUser, adminPass);
  if (!auth.ok) return {success:false,error:auth.error};
  try {
    const props = PropertiesService.getScriptProperties();
    const term = JSON.parse(props.getProperty('PICKUP_ACTIVE_TERM') ||
      '{"id":"115-1","start":"2026-08-31","end":"2027-01-20"}');
    if (!/^\d{3}-[12]$/.test(term.id) ||
        !/^\d{4}-\d{2}-\d{2}$/.test(term.start) ||
        !/^\d{4}-\d{2}-\d{2}$/.test(term.end) || term.start > term.end) {
      throw new Error('接送學期設定不完整');
    }
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(ROUTES_SHEET);
    const anchor = ss.getRangeByName('XG_PICKUP_' + term.id.replace('-', '_'));
    if (!sh || !anchor || anchor.getSheet().getSheetId() !== sh.getSheetId() || anchor.getRow() < 2) {
      throw new Error('缺少本學期接送範圍，請先完成後台學期設定；不回退全表');
    }
    const first = anchor.getRow(), last = sh.getLastRow();
    // Read only the active block. No _ensureRoutesSheet() full-column formatting.
    // No server list cache: each refresh sees current sheet values.
    const values = last >= first ? sh.getRange(first,1,last-first+1,ROUTES_COLS).getValues() : [];
    const rows = values.map(function(row,index) {
      const obj = _routeRow2Obj(row,first+index);
      if (row[0] instanceof Date) obj.date = Utilities.formatDate(row[0],'Asia/Taipei','yyyy-MM-dd');
      return obj;
    }).filter(function(row){return row.date >= term.start && row.date <= term.end;});
    rows.sort(function(a,b){return a.date.localeCompare(b.date);});
    return {success:true,rows:rows,semester:term.id,start:term.start,end:term.end,
      scoped:true,source:'接送排程',readRows:values.length};
  } catch(e) {
    return {success:false,error:e.message};
  }
}
