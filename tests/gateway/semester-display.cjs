const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict'),path=require('node:path');
const source=fs.readFileSync(path.join(__dirname,'../../apps-script/semester-display/semesterAssignmentsList.gs'),'utf8');
let values=[['學生中英文姓名'],[''],['測試學生（Alex）']],authorized=true;
const rows=[{_row:2,'學生編號':'S1','學期':'114-2','學生姓名':'歷史姓名'},{_row:3,'學生編號':'S1','學期':'115-1','學生姓名':'測試學生'}];
const ctx=vm.createContext({
  SEMESTER_ASSIGN_SHEET:'學期班級指派',SEMESTER_ASSIGN_HEADERS:Array(22).fill(''),
  _verifyStaffOrTeacher:()=>({ok:authorized,error:'denied'}),
  _studentDataRows:()=>rows,_currentStudentSemester:()=>'115-1',_studentModelActive:()=>true,
  _studentText:x=>String(x||'').trim(),
  _studentDataSheet:()=>({getLastColumn:()=>values?23:22,getLastRow:()=>3,getRange:(row,col,count,width)=>{
    assert.equal(row,1);assert.equal(col,23);assert.equal(count,3);assert.equal(width,1);
    return {getValues:()=>values};
  }})
});
vm.runInContext(source,ctx);
let r=ctx._semesterAssignmentsList('synthetic','synthetic','115-1');
assert.equal(r.success,true);assert.equal(r.displayNameSource,'學期班級指派!W');
assert.equal(r.list[0]['學生中英文姓名'],'測試學生（Alex）');
r=ctx._semesterAssignmentsList('synthetic','synthetic','114-2');
assert.equal(r.list[0]['學生中英文姓名'],undefined);
for(const invalid of ['','#REF!','#ERROR!']) {
  values[2][0]=invalid;
  assert.equal(ctx._semesterAssignmentsList('synthetic','synthetic','115-1').success,false);
}
values=null;r=ctx._semesterAssignmentsList('synthetic','synthetic','115-1');
assert.equal(r.success,true);assert.equal(r.displayNameSource,'');
authorized=false;assert.equal(ctx._semesterAssignmentsList('synthetic','synthetic','115-1').success,false);
console.log('PASS: W-column read, historical isolation, formula-error fail-closed, legacy compatibility, authorization. No writes or network.');
