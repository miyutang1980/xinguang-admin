const fs=require('node:fs'), vm=require('node:vm'), assert=require('node:assert/strict');
const code=fs.readFileSync(require('node:path').join(__dirname,'../../apps-script/pickup-semester/routesList.gs'),'utf8');
let allowed=true,anchorRow=201,rangeMissing=false,config=null,reads=[];
let data=[['2026-08-31','走路'],['2026-09-21','A'],['2027-02-01','future']];
const sheet={getSheetId:()=>1515463418,getLastRow:()=>203,
 getRange:(start,col,count,width)=>{reads.push({start,col,count,width});return {getValues:()=>data};}};
const ctx={_verifyPickupUser:()=>({ok:allowed,error:'denied'}),
 PropertiesService:{getScriptProperties:()=>({getProperty:()=>config})},
 SpreadsheetApp:{openById:()=>({getSheetByName:()=>sheet,getRangeByName:()=>rangeMissing?null:{getRow:()=>anchorRow,getSheet:()=>sheet}})},
 SHEET_ID:'mock',ROUTES_SHEET:'接送排程',ROUTES_COLS:22,
 _routeRow2Obj:(r,i)=>({row_index:i,date:r[0],route:r[1]}),Date};
vm.createContext(ctx);vm.runInContext(code,ctx);
let result=ctx._routesList('test','test');
assert.equal(result.scoped,true);assert.equal(result.semester,'115-1');
assert.deepEqual(Array.from(result.rows,r=>r.row_index),[201,202]);
assert.deepEqual(reads,[{start:201,col:1,count:3,width:22}]);
reads=[];rangeMissing=true;assert.equal(ctx._routesList().success,false);assert.equal(reads.length,0);
rangeMissing=false;allowed=false;assert.equal(ctx._routesList().success,false);assert.equal(reads.length,0);
allowed=true;anchorRow=204;assert.equal(ctx._routesList().rows.length,0);
anchorRow=201;data=[['2026-09-21','changed']];assert.equal(ctx._routesList().rows[0].route,'changed');
config='{"id":"115-2","start":"2027-02-01","end":"2027-06-30"}';
assert.equal(ctx._routesList().rows.length,0);
console.log('PASS: active range only, original row IDs, fail-closed missing config/auth, fresh reads, term switch; no production writes');
