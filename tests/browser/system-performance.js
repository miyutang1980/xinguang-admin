// Synthetic API-only tests. Never issue production reads or writes.
var perfAssert=(await import('node:assert/strict')).default;
var perfCounts={}, perfHeld=[], perfMode='ok';
await ssQAContext.route('**/*', async route=>{
  const req=route.request(),u=new URL(req.url());
  let action=u.searchParams.get('_action')||u.searchParams.get('action');
  if(req.method()==='POST') {try{const p=req.postDataJSON(); action=p._action||p.action||action;}catch(_){}}
  if(!['bookingList','trialBookings','calendarSettings','leads_list','classes_save','announce_list'].includes(action))return route.fallback();
  perfCounts[action]=(perfCounts[action]||0)+1;
  if(perfMode==='hold'){perfHeld.push(route);return;}
  if(perfMode==='fail')return route.abort('failed');
  if(perfMode==='error')return route.fulfill({json:{success:false,error:'Synthetic denied'}});
  let data={success:true,ok:true};
  if(action==='bookingList'||action==='trialBookings')data.bookings=[
    {_row:2,parent:'Alpha家長',phone:'0001',child:'合成學生甲',type:'體驗課',date:'2026-09-23',slot:'13:00',status:'待確認'},
    {_row:3,parent:'Beta家長',phone:'0002',child:'合成學生乙',type:'體驗課',date:'2026-09-24',slot:'13:00',status:'已確認'}];
  if(action==='calendarSettings')data.settings={blockedDates:[],blockedSlots:[]};
  if(action==='leads_list')data.leads=[];
  return route.fulfill({json:data});
});
var perfMetrics=[];
var measureUI=async(label,act,ready)=>{
  const start=performance.now();await act();await ssQAPage.locator(ready).waitFor({timeout:3000});
  const ms=Math.round(performance.now()-start);perfAssert.ok(ms<3000,label);
  perfMetrics.push({label,milliseconds:ms});
};
await measureUI('已載入主檔切回',async()=>{
  await ssQAPage.locator('#nav-dashboard').click();await ssQAPage.locator('#nav-students').click();
},'#ssMasterSearch');
var masterReadsBefore=ssQACalls.filter(c=>c.action==='student_master_list').length;
await measureUI('主檔重訪',async()=>{
  await ssQAPage.locator('#nav-dashboard').click();await ssQAPage.locator('#nav-students').click();
},'#ssMasterSearch');
perfAssert.equal(ssQACalls.filter(c=>c.action==='student_master_list').length,masterReadsBefore);
await measureUI('主檔編輯視窗',()=>ssQAPage.locator('.ss-row-edit').first().click(),'#ssEditor');
await ssQAPage.locator('#ssEditor [data-field="英文名"]').fill('EditorTest');
ssQAMasters[0]['學生姓名']='其他裝置已修改';
await ssQAPage.locator('#ssSave').click();
await ssQAPage.waitForFunction(()=>document.querySelector('.ss-form-error')?.textContent.includes('資料已被其他人修改'));
perfAssert.equal(ssQAMutations.length,0,'warm view never skips prewrite conflict check');
await ssQAPage.locator('#ssCancel').click();
await ssQAPage.locator('#ssRefresh').click();
await ssQAPage.waitForFunction(()=>document.querySelector('#ssList')?.textContent.includes('其他裝置已修改'));
await ssQAPage.locator('.ss-row-edit').first().click();
await ssQAPage.locator('#ssEditor [data-field="英文名"]').fill('EditorTest');
var beforeSave=ssQACalls.length;
await ssQAPage.locator('#ssSave').click();
await ssQAPage.waitForFunction(()=>document.querySelector('#ssNotice')?.textContent.includes('已儲存並讀回確認'));
perfAssert.deepEqual(ssQACalls.slice(beforeSave).filter(c=>c.action.startsWith('student_master')).map(c=>c.action),
  ['student_master_list','student_master_update','student_master_list'],'fresh check + one write + fresh verification');

await measureUI('預約列表（合成快速後端）',()=>ssQAPage.locator('#nav-booking').click(),'#searchBooking');
var bookingReads=perfCounts.bookingList;
await ssQAPage.locator('#searchBooking').fill('Alpha');
perfAssert.equal(await ssQAPage.locator('#bookingTbody tr:visible').count(),1);
perfAssert.equal(await ssQAPage.locator('#searchBooking').inputValue(),'Alpha');
await ssQAPage.locator('#searchBooking').fill('');
await ssQAPage.locator('#filterBookingStatus').selectOption('已確認');
perfAssert.equal(await ssQAPage.locator('#bookingTbody tr:visible').count(),1);
perfAssert.match(await ssQAPage.locator('#bookingTbody tr:visible').innerText(),/Beta/);
perfAssert.equal(perfCounts.bookingList,bookingReads,'typing/status filtering sends no API requests');
await ssQAPage.screenshot({path:'/home/user/workspace/frontend_qa/system-search-desktop.png'});
await ssQAPage.setViewportSize({width:375,height:812});
await ssQAPage.locator('#searchBooking').scrollIntoViewIfNeeded();
await ssQAPage.screenshot({path:'/home/user/workspace/frontend_qa/system-search-mobile.png'});
await ssQAPage.setViewportSize({width:1440,height:960});
await measureUI('體驗預約列表（合成快速後端）',()=>ssQAPage.locator('#nav-trialBookings').click(),'#searchTrialBkg');
var trialReads=perfCounts.trialBookings;
await ssQAPage.locator('#searchTrialBkg').fill('Alpha');
perfAssert.equal(await ssQAPage.locator('#trialBkgTbody tr:visible').count(),1);
perfAssert.equal(perfCounts.trialBookings,trialReads);

// Every remaining legacy search retains its input and uses existing rows only.
await ssQAPage.evaluate(()=>{
  const rows=['Alpha','Beta'].map((parent,i)=>({parent,pname:parent,sname:parent,studentName:parent,child:parent,student:parent,phone:'000'+i,
    ts:'2026-09-23T10:00:0'+i,children:[],status:'待處理'}));
  _DB.inquiries=rows;_DB.camp=rows;_DB.assessments=rows;_DB.registrations=rows;
  _assessLastRefresh=Date.now();
});
for (const [nav,search,body] of [
  ['inquiry','searchInquiry','inquiryTbody'],['camp','searchCamp','campTbody'],
  ['assess','searchAssess','assessTbody'],['register','searchReg','regTbody']
]) {
  await ssQAPage.locator('#nav-'+nav).click();
  await ssQAPage.locator('#'+search).waitFor();
  const before=ssQACalls.length;
  await ssQAPage.locator('#'+search).fill('Alpha');
  perfAssert.equal(await ssQAPage.locator('#'+search).inputValue(),'Alpha');
  perfAssert.equal(await ssQAPage.locator('#'+body+' tr:visible').count(),1,nav);
  perfAssert.equal(ssQACalls.length,before,nav+' search is local');
}

// Health calls launch in parallel and late completion cannot overwrite another page.
const healthResult=await ssQAPage.evaluate(async()=>{
  const originalURL=checkUrlAlive, originalHeaders=fetchSheetHeaders;
  const held=[], calls=[];
  checkUrlAlive=url=>new Promise(resolve=>{calls.push('service');held.push(()=>resolve({ok:true,status:'test'}));});
  fetchSheetHeaders=name=>new Promise(resolve=>{calls.push('sheet');held.push(()=>resolve(SHEET_EXPECTED_HEADERS[name]));});
  try {
    const first=renderHealth(document.getElementById('mainContent'));
    const second=renderHealth(document.getElementById('mainContent'));
    const started=calls.length;
    switchSection('dashboard');
    const before=document.getElementById('mainContent').innerHTML;
    held.forEach(release=>release());
    await Promise.all([first,second]);
    return {started,unchanged:before===document.getElementById('mainContent').innerHTML};
  } finally {checkUrlAlive=originalURL;fetchSheetHeaders=originalHeaders;}
});
perfAssert.deepEqual(healthResult,{started:8,unchanged:true});
// Repeated sync clicks return the same operation rather than downloading twice.
const syncSame=await ssQAPage.evaluate(async()=>{
  const a=syncFromSheets(),b=syncFromSheets();const same=a===b;
  await Promise.all([a,b]);return same;
});
perfAssert.equal(syncSame,true);

// All read consumers get independent Response bodies; concurrent calls share work.
await ssQAPage.evaluate(()=>xgClearGatewayCache());
var beforeShared=perfCounts.calendarSettings||0;
var shared=await ssQAPage.evaluate(async()=>Promise.all([
  fetch(GW_URL+'?action=calendarSettings').then(r=>r.json()),
  fetch(GW_URL+'?action=calendarSettings').then(r=>r.json())
]));
perfAssert.equal(shared.length,2);
perfAssert.equal(perfCounts.calendarSettings,beforeShared+1);
await ssQAPage.evaluate(()=>fetch(GW_URL+'?action=calendarSettings').then(r=>r.json()));
perfAssert.equal(perfCounts.calendarSettings,beforeShared+1,'warm read cache');
await ssQAPage.evaluate(()=>fetch(GW_URL+'?action=calendarSettings',{cache:'no-store'}).then(r=>r.json()));
perfAssert.equal(perfCounts.calendarSettings,beforeShared+2,'explicit fresh bypass');
await ssQAPage.evaluate(()=>{_authEpoch++;return fetch(GW_URL+'?action=calendarSettings').then(r=>r.json());});
perfAssert.equal(perfCounts.calendarSettings,beforeShared+3,'login epoch isolation');

// Error responses not cached; transport failures not retried, including mutations.
perfMode='error';
await ssQAPage.evaluate(()=>xgClearGatewayCache());
var beforeErrors=perfCounts.leads_list||0;
await ssQAPage.evaluate(async()=>{await gwCallJson('leads_list');await gwCallJson('leads_list');});
perfAssert.equal(perfCounts.leads_list,beforeErrors+2);
perfMode='fail';
var beforeFailures=perfCounts.leads_list;
await ssQAPage.evaluate(()=>gwCallJson('leads_list').catch(()=>null));
perfAssert.equal(perfCounts.leads_list,beforeFailures+1,'no hidden retry');
var beforeMutations=perfCounts.classes_save||0;
await ssQAPage.evaluate(()=>gwCallJson('classes_save',{data_json:'{}'}).catch(()=>null));
perfAssert.equal(perfCounts.classes_save,beforeMutations+1,'GET write sent once on failure');
perfMode='ok';
await ssQAPage.evaluate(async()=>{
  await gwCallJson('classes_save',{data_json:'{}'});
  await gwCallJson('classes_save',{data_json:'{}'});
});
perfAssert.equal(perfCounts.classes_save,beforeMutations+3,'GET writes never cached');

// Stale reads cannot populate caches after a mutation.
perfMode='hold';
await ssQAPage.evaluate(()=>{
  xgClearGatewayCache();
  window.perfOld=gwCallJson('leads_list').then(()=> 'wrong').catch(e=>e.message);
});
await ssQAPage.waitForFunction(()=>window._gatewayTimings.some(x=>x.action==='leads_list'));
while(!perfHeld.length)await ssQAPage.waitForTimeout(10);
perfMode='ok';
await ssQAPage.evaluate(()=>gwCallJson('classes_save',{data_json:'{}'}));
await perfHeld.shift().fulfill({json:{success:true,leads:[]}});
perfAssert.match(await ssQAPage.evaluate(()=>window.perfOld),/狀態已變更/);
// Legacy read timeout ends once, not 15s + 15s. Use controlled browser time.
perfMode='hold';
await ssQAPage.clock.install();
await ssQAPage.evaluate(()=>{
  xgClearGatewayCache();
  window.perfTimeout=gwCallJson('announce_list',{status:'test'}).then(()=> 'wrong').catch(e=>e.message);
});
while(!perfHeld.length)await ssQAPage.waitForTimeout(10);
await ssQAPage.clock.fastForward(15001);
perfAssert.match(await ssQAPage.evaluate(()=>window.perfTimeout),/15 秒|abort/i);
perfAssert.equal(perfCounts.announce_list,1);
perfAssert.equal(await ssQAPage.evaluate(()=>JSON.stringify(window._gatewayTimings).includes('offline-only')),false);
perfAssert.deepEqual(ssQAErrors,[]);
console.log({systemPerformance:'PASS',metrics:perfMetrics,
 checks:['warm navigation','editor under 3s','live prewrite conflict','read-write-read preserved','six local filters',
 'parallel single-flight health with navigation guard','single-flight sheet sync',
 'shared reads','fresh bypass','auth isolation','errors not cached','no read/write retry','mutation invalidation','bounded 15s once'],
 productionWrites:0});
