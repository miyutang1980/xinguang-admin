const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict'),path=require('node:path');
const source=fs.readFileSync(path.join(__dirname,'../../apps-script/semester-display/semesterAssignmentsList.gs'),'utf8');
const headers=['指派編號','學期','學生編號','學生帳號','學生姓名','弋果班級','弋果課程','學生類別',
  '學校','年級','小學班級','外籍教師','中籍教師','TXClass','學年度起始日','課堂時間','課程類別','教室','學期狀態','課後輔導','交通車','更新時間'];
const makeRow=(term,id)=>[term+'|'+id,term,id,'','合成姓名',...Array(17).fill(''),'合成姓名（Alex）'];
let authorized=true,anchorStart=53,anchorPresent=true,header=[...headers,'學生中英文姓名'];
let rows=Array.from({length:54},()=>Array(23).fill(''));
rows[0]=header; rows[1]=makeRow('114-2','S1'); rows[52]=makeRow('115-1','S1'); rows[53]=makeRow('115-1','S2');
let reads=[],opens=0;
const sheet={
  getSheetId:()=>921002,getLastRow:()=>rows.length,
  getRange:(row,col,count,width)=>{
    reads.push({row,col,count,width});
    assert.equal(col,1);assert.equal(width,23);
    return {getValues:()=>row===1?[header]:rows.slice(row-1,row-1+count)};
  }
};
const book={getSheetByName:()=>sheet,getRangeByName:name=>{
  assert.equal(name,'XG_STUDENTS_115_1');
  return anchorPresent?{getSheet:()=>sheet,getRow:()=>anchorStart,getColumn:()=>1,getNumColumns:()=>23}:null;
}};
const ctx=vm.createContext({
  SHEET_ID:'fake',SEMESTER_ASSIGN_SHEET:'學期班級指派',SEMESTER_ASSIGN_HEADERS:headers,
  _verifyStaffOrTeacher:()=>({ok:authorized,error:'denied'}),
  _currentStudentSemester:()=>'115-1',_studentModelActive:()=>true,_studentText:x=>String(x||'').trim(),
  SpreadsheetApp:{openById:()=>{opens++;return book;}},
  PropertiesService:{getScriptProperties:()=>({getProperty:()=>null})}
});
vm.runInContext(source,ctx);
const call=(term='115-1')=>{reads=[];opens=0;return ctx._semesterAssignmentsList('synthetic','synthetic',term);};
let r=call();
assert.equal(r.success,true);assert.equal(r.scoped,true);assert.equal(r.readRows,2);
assert.equal(r.list[0]._row,53);assert.equal(r.list[0]['學生中英文姓名'],'合成姓名（Alex）');
assert.deepEqual(reads,[{row:1,col:1,count:1,width:23},{row:53,col:1,count:2,width:23}]);
assert.equal(opens,1);assert.equal(r.readerVersion,'semester-scoped-v2');
r=call('114-2');assert.equal(r.success,true);assert.equal(r.list.length,1);
assert.equal(r.list[0]._row,2);assert.equal(r.list[0]['學生中英文姓名'],undefined);
for (const value of ['','#REF!','#ERROR!']) {
  rows[52][22]=value;assert.equal(call().success,false);
}
rows[52][22]='合成姓名（Alex）';
anchorPresent=false;assert.equal(call().success,false);assert.equal(reads.length,1);
anchorPresent=true;anchorStart=2;assert.equal(call().success,false);anchorStart=53;
rows[53][2]='S1';assert.equal(call().success,false);rows[53]=makeRow('115-1','S2');
header[22]='錯誤欄位';assert.equal(call().success,false);header[22]='學生中英文姓名';
rows.push(makeRow('115-1','S3'));r=call();assert.equal(r.readRows,3);assert.equal(r.list[2]._row,55);
authorized=false;assert.equal(call().success,false);assert.equal(reads.length,0);assert.equal(opens,0);
console.log('PASS: current range only, one payload read incl W, original row IDs, append coverage, explicit history, missing/mixed range + formula + duplicates fail closed, auth before reads. No writes/network.');
