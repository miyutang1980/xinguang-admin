// Synthetic held requests: test responsiveness rather than pretending to measure production.
var latencyAssert=(await import('node:assert/strict')).default;
var latencyMode='hold-roster', rosterHeld=[], detailsHeld=[], rosterReads=0, detailReads=0, batchReads=0;
ssQAMasters[0]['英文名']='Alice';
var latencyRows=[{...ssQAAssignments[0],'學校':'新光國小','學生中英文姓名':'測試用匿名記錄1（Alice）'}];
var finishRoster=route=>route.fulfill({json:{success:true,active:true,currentSemester:'115-1',semester:'115-1',
  semesters:['115-1','114-2'],displayNameSource:'學期班級指派!W',list:latencyRows}});
await ssQAContext.route('**/*',route=>{
  const action=new URL(route.request().url()).searchParams.get('_action');
  if(action==='semester_assignments_list'){
    rosterReads++;
    if(latencyMode==='legacy')return route.fulfill({json:{success:true,active:true,
      currentSemester:'115-1',semester:'115-1',semesters:['115-1'],list:latencyRows.map(row=>{
        const copy={...row};delete copy['學生中英文姓名'];return copy;
      })}});
    if(latencyMode==='hold-roster'||latencyMode==='hold-both'){rosterHeld.push(route);return;}
    return finishRoster(route);
  }
  if(action==='routes_get_staff')return route.fulfill({json:{success:true,staff:[{name:'合成人員'}]}});
  if(action==='routes_list')return route.fulfill({json:{success:true,rows:[{row_index:206,date:'2026-09-21',weekday:'週一',
    route:'合成接送',status:'上',noon_school:'新光國小',noon_count:0}]}});
  if(action==='routes_get_details'){
    detailReads++;
    if(latencyMode==='hold-both'||latencyMode==='timeout'){detailsHeld.push(route);return;}
    return route.fulfill({json:{success:true,details:[]}});
  }
  if(action==='routes_get_details_batch'){batchReads++;return route.fulfill({status:503,body:'synthetic outage'});}
  return route.fallback();
});
await ssQAPage.evaluate(()=>invalidatePickupRoster());
await ssQAPage.locator('#nav-pickup').click();
await ssQAPage.waitForFunction(()=>document.querySelector('#ps-table-wrap tr[data-row="206"]'));
latencyAssert.equal(rosterReads,1);
latencyAssert.equal(await ssQAPage.locator('.btn-detail').first().isDisabled(),true);
latencyAssert.equal(await ssQAPage.locator('#ps-table-wrap input:enabled, #ps-table-wrap select:enabled, #ps-table-wrap button:enabled').count(),0);
latencyAssert.match(await ssQAPage.locator('#ps-status').textContent(),/排程已取得/);
await ssQAPage.evaluate(()=>{void _loadAllStudentsOnce({});void _loadAllStudentsOnce({});});
latencyAssert.equal(rosterReads,1,'concurrent callers share one roster request');
await ssQAPage.screenshot({path:'/home/user/workspace/frontend_qa/latency-pending-desktop.png'});
await ssQAPage.setViewportSize({width:375,height:812});
await ssQAPage.screenshot({path:'/home/user/workspace/frontend_qa/latency-pending-mobile.png'});
latencyMode='ok';
await finishRoster(rosterHeld.shift());
await ssQAPage.waitForFunction(()=>!document.querySelector('.btn-detail').disabled);
// A freshly read assignment list primes the shared roster: entering pickup needs no second read.
await ssQAPage.locator('#nav-classAssign').click();
await ssQAPage.locator('#ssSearch').waitFor();
var readsAfterAssignment=rosterReads;
await ssQAPage.locator('#nav-pickup').click();
await ssQAPage.waitForFunction(()=>document.querySelector('.btn-detail')&&!document.querySelector('.btn-detail').disabled);
latencyAssert.equal(rosterReads,readsAfterAssignment);
// Both requests must start before either response is released.
latencyMode='hold-both';
await ssQAPage.evaluate(()=>{invalidatePickupRoster();void openRouteDetailsModal({
  row_index:206,date:'2026-09-21',weekday:'週一',route:'合成接送',noon_school:'新光國小'
},'noon',{});});
for(var i=0;(!detailsHeld.length||!rosterHeld.length)&&i<50;i++)await ssQAPage.waitForTimeout(20);
latencyAssert.equal(detailsHeld.length,1);
latencyAssert.equal(rosterHeld.length,1);
latencyAssert.equal(await ssQAPage.locator('#rd-save').isDisabled(),true);
await finishRoster(rosterHeld.shift());
await detailsHeld.shift().fulfill({json:{success:true,details:[]}});
await ssQAPage.waitForFunction(()=>!document.querySelector('#rd-save').disabled);
await ssQAPage.locator('#rd-close').click();
// Accelerated browser clock verifies the real configured deadline, without waiting 20 real seconds.
await ssQAPage.clock.install();
latencyMode='timeout';
await ssQAPage.evaluate(()=>{void openRouteDetailsModal({
  row_index:206,date:'2026-09-21',weekday:'週一',route:'合成接送'
},'noon',{});});
for(var i=0;!detailsHeld.length&&i<50;i++)await ssQAPage.waitForTimeout(20);
latencyAssert.equal(detailsHeld.length,1);
await ssQAPage.clock.fastForward(21000);
await ssQAPage.locator('#rd-retry').waitFor();
latencyAssert.match(await ssQAPage.locator('#rd-body').textContent(),/本班次已排學生讀取超過 20 秒/);
latencyAssert.equal(await ssQAPage.locator('#rd-save').isDisabled(),true);
await ssQAPage.screenshot({path:'/home/user/workspace/frontend_qa/latency-timeout-mobile.png'});
await ssQAPage.locator('#rd-cancel').click();
await detailsHeld.shift().fulfill({json:{success:true,details:[]}}).catch(()=>{});
// Batch outage cannot fan out into per-row requests or cache failure as empty names.
latencyMode='ok';
var detailsBeforeBatch=detailReads;
await ssQAPage.evaluate(async()=>{
  window._routeNamesCache={};
  await Promise.all([_hydrateRouteNames([{row_index:777,noon_count:2}],{}),
    _hydrateRouteNames([{row_index:777,noon_count:2}],{})]);
});
latencyAssert.equal(batchReads,1);
latencyAssert.equal(detailReads,detailsBeforeBatch);
latencyAssert.equal(await ssQAPage.evaluate(()=>window._routeNamesCache[777].noon===undefined),true);
latencyAssert.equal(await ssQAPage.evaluate(()=>window._routeNamesErrors['777:noon']),true);
// Missing W marker must show an honest compatibility source, not Chinese-only rows.
latencyMode='legacy';
await ssQAPage.evaluate(()=>invalidatePickupRoster());
await ssQAPage.locator('#nav-classAssign').click();
await ssQAPage.waitForFunction(()=>document.querySelector('#ssNotice')?.textContent.includes('主檔中英文相容模式'));
latencyAssert.match(await ssQAPage.locator('#ssNotice').textContent(),/未回傳 W|未回傳W/);
latencyAssert.match(await ssQAPage.locator('#ssList').textContent(),/Alice/);
latencyAssert.deepEqual(ssQAErrors,[]);
console.log({latencyRegression:'PASS',checks:['table visible before roster resolves','single-flight roster',
  'assignment snapshot reused','parallel modal reads','20s labelled deadline','close after timeout',
  'batch outage causes no fanout or false empty cache','legacy bilingual source labelled'],productionWrites:0});
