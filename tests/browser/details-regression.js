// Dedicated modal regression: HTTP failure is not an empty editable roster.
var detailsAssert = (await import('node:assert/strict')).default;
var detailsMode = '404', detailsWrites = 0, detailsPending = [];
ssQAAssignments.forEach(s => { s['學校'] = '新光國小'; });
ssQAMasters[0]['英文名'] = 'Alice <QA>';
ssQAMasters[1]['英文名'] = 'Bob';
ssQAMasters.reverse(); // Joining by row order or Chinese name would pick the wrong name.
await ssQAContext.route('**/*', async route => {
  const action = new URL(route.request().url()).searchParams.get('_action');
  if (action === 'routes_get_details') {
    if (detailsMode === 'hold') { detailsPending.push(route); return; }
    if (detailsMode === '404') return route.fulfill({status:404, body:'not found'});
    if (detailsMode === 'auth') return route.fulfill({json:{success:false,error:'權限驗證失敗'}});
    if (detailsMode === 'schema') return route.fulfill({json:{success:true,details:{}}});
    if (detailsMode === 'legacy') return route.fulfill({json:{success:true,details:[{school_name:'新光國小',student_names:['舊明細匿名學生 (301)']}] }});
    if (detailsMode === 'missing-id') return route.fulfill({json:{success:true,details:[{school_name:'新光國小',student_ids:['QA-missing'],student_names:['未在學學生']}] }});
    return route.fulfill({json:{success:true,details:[]}});
  }
  if (action === 'routes_save_details') {
    detailsWrites++;
    const payload = JSON.parse(new URL(route.request().url()).searchParams.get('details_json'));
    detailsAssert.equal(payload[0].student_names[0],'測試用匿名記錄1','bilingual labels do not alter saved Chinese names');
    detailsAssert.deepEqual(payload[0].student_ids,['QA1'],'student IDs remain stable');
    return detailsMode === 'save-fail'
      ? route.fulfill({status:503,body:'uncertain'})
      : route.fulfill({json:{success:true,total_count:1}});
  }
  return route.fallback();
});
var openDetails = () => ssQAPage.evaluate(() => {
  _allStudentsCache = null;
  return openRouteDetailsModal({
    row_index:206,date:'2026-09-21',weekday:'週一',route:'走路接送',
    noon_school:'新光國小',capacity:'4'
  },'noon',{}, count => {window.detailsSavedCount=count;});
});
for (var mode of ['404','auth','schema']) {
  detailsMode=mode;
  await openDetails();
  detailsAssert.equal(await ssQAPage.locator('#rd-save').isDisabled(),true);
  detailsAssert.equal(await ssQAPage.locator('#rd-clear-all').isDisabled(),true);
  detailsAssert.equal(await ssQAPage.locator('#rd-summary').textContent(),'人數尚未確認');
  detailsAssert.equal(await ssQAPage.locator('#rd-retry').count(),1);
  if(mode==='404') await ssQAPage.screenshot({path:'/home/user/workspace/frontend_qa/details-error-desktop.png'});
  await ssQAPage.locator('#rd-cancel').click();
  detailsAssert.equal(await ssQAPage.locator('#rd-modal-overlay').count(),0);
}
// Retry retains the onSaved closure, fetches data, and enables only after a valid roster.
detailsMode='404'; await openDetails();
detailsMode='ok'; await ssQAPage.locator('#rd-retry').click();
await ssQAPage.waitForFunction(()=>!document.querySelector('#rd-save').disabled);
detailsAssert.equal(await ssQAPage.locator('.rd-stu').count(),1,'inactive student excluded');
detailsAssert.equal(await ssQAPage.locator('.rd-stu').locator('..').textContent().then(s=>s.trim()),'測試用匿名記錄1（Alice <QA>）','English name joins by student ID and is safely escaped');
detailsAssert.equal(await ssQAPage.locator('#rd-body qa').count(),0,'name is text, not markup');
var masterCallsBefore = ssQACalls.filter(c=>c.action==='student_master_list').length;
await ssQAPage.evaluate(()=>_getSchoolStudents('新高國小',{}));
detailsAssert.equal(ssQACalls.filter(c=>c.action==='student_master_list').length,masterCallsBefore,'one master fetch serves all schools');
await ssQAPage.locator('.rd-stu').check();
await ssQAPage.locator('#rd-save').click();
await ssQAPage.waitForFunction(()=>window.detailsSavedCount===1);
detailsAssert.equal(detailsWrites,1);
// Successful legacy response remains visible, not silently converted to zero checked students.
detailsMode='legacy'; await openDetails();
detailsAssert.equal(await ssQAPage.locator('#rd-save').isDisabled(),true);
detailsAssert.match(await ssQAPage.locator('#rd-body').textContent(),/舊明細匿名學生/);
await ssQAPage.locator('#rd-close').click();
// Missing English name is not guessed from another student or assignment.
ssQAMasters.find(m=>m['學生編號']==='QA1')['英文名']='';
detailsMode='ok'; await openDetails();
detailsAssert.equal(await ssQAPage.locator('.rd-stu').locator('..').textContent().then(s=>s.trim()),'測試用匿名記錄1');
await ssQAPage.locator('#rd-close').click();
ssQAMasters.find(m=>m['學生編號']==='QA1')['英文名']='Alice <QA>';
detailsMode='missing-id'; await openDetails();
detailsAssert.equal(await ssQAPage.locator('#rd-save').isDisabled(),true);
detailsAssert.match(await ssQAPage.locator('#rd-body').textContent(),/不在目前學期/);
await ssQAPage.locator('#rd-close').click();
// Roster errors are not cached as empty; retry works after server recovery.
detailsMode='ok'; ssQAMockMode='error'; await openDetails();
detailsAssert.equal(await ssQAPage.locator('#rd-save').isDisabled(),true);
ssQAMockMode='active'; await ssQAPage.locator('#rd-retry').click();
await ssQAPage.waitForFunction(()=>!document.querySelector('#rd-save').disabled);
detailsAssert.equal(await ssQAPage.locator('.rd-stu').count(),1);
await ssQAPage.setViewportSize({width:375,height:812});
await ssQAPage.screenshot({path:'/home/user/workspace/frontend_qa/details-mobile.png'});
detailsAssert.equal(await ssQAPage.evaluate(()=>document.querySelector('#rd-modal-overlay').scrollWidth<=innerWidth),true);
await ssQAPage.locator('#rd-close').click();
// A late response must never write into the next modal.
detailsMode='hold';
await ssQAPage.evaluate(()=>{void openRouteDetailsModal({row_index:999,date:'2026-09-21',route:'舊視窗'},'noon',{});});
for(var attempts=0;!detailsPending.length&&attempts<50;attempts++) await ssQAPage.waitForTimeout(20);
detailsAssert.equal(detailsPending.length,1);
detailsMode='404'; await openDetails();
await detailsPending[0].fulfill({json:{success:true,details:[]}});
await ssQAPage.waitForTimeout(100);
detailsAssert.equal(await ssQAPage.locator('#rd-retry').count(),1);
detailsAssert.equal(await ssQAPage.locator('#rd-save').isDisabled(),true);
await ssQAPage.locator('#rd-close').click();
// A write transport error cannot enable a blind replay.
detailsMode='save-fail'; await openDetails();
await ssQAPage.locator('.rd-stu').check();
await ssQAPage.locator('#rd-save').click();
await ssQAPage.locator('#rd-retry').waitFor();
detailsAssert.equal(await ssQAPage.locator('#rd-save').isDisabled(),true);
detailsAssert.match(await ssQAPage.locator('#rd-body').textContent(),/請勿重複送出/);
detailsAssert.doesNotMatch(await ssQAPage.locator('#rd-body').textContent(),/沒有更動資料/);
detailsAssert.equal(detailsWrites,2);
detailsAssert.deepEqual(ssQAErrors,[]);
console.log({detailsRegression:'PASS',checks:['404/auth/schema fail closed','close on failure','retry preserves callback','legacy names readonly','unmatched IDs blocked','roster failure recoverable','late response ignored','ambiguous save not replayed','mobile fits'],productionWrites:0});
