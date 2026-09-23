// Regression for the 10:16 screenshot: a roster fetched four minutes ago must
// not be discarded merely because the user opens a different pickup detail.
var efficiencyAssert=(await import('node:assert/strict')).default;
var efficiencyRosterReads=0,efficiencyDetailReads=0,efficiencyWrites=0,efficiencyMode='ok',efficiencyHeld=[];
var efficiencyRows=ssQAAssignments.slice(0,2).map((r,i)=>({...r,'學期':'115-1','學期狀態':'在學',
  '學校':i?'新高國小':'新光國小','學生中英文姓名':r['學生姓名']+'（Test）'}));
var efficiencyRoster=()=>({success:true,active:true,semester:'115-1',currentSemester:'115-1',
  semesters:['115-1','114-2'],displayNameSource:'學期班級指派!W',list:efficiencyRows});
await ssQAContext.route('**/*',route=>{
  const action=new URL(route.request().url()).searchParams.get('_action');
  if(action==='semester_assignments_list'){
    efficiencyRosterReads++;
    if(efficiencyMode==='hold-roster'){efficiencyHeld.push(route);return;}
    return route.fulfill({json:efficiencyRoster()});
  }
  if(action==='routes_get_staff')return route.fulfill({json:{success:true,staff:[{name:'合成人員'}]}});
  if(action==='routes_list')return route.fulfill({json:{success:true,rows:[{row_index:206,date:'2026-09-21',
    weekday:'週一',route:'合成交通車A',noon_school:'新光國小',noon_count:0,status:'上'}]}});
  if(action==='routes_get_details'){
    efficiencyDetailReads++;
    if(efficiencyMode==='details-fail')return route.fulfill({status:503,body:'synthetic outage'});
    return route.fulfill({json:{success:true,details:[]}});
  }
  if(action==='routes_save_details'){efficiencyWrites++;return route.fulfill({json:{success:true}});}
  return route.fallback();
});
await ssQAPage.evaluate(()=>invalidatePickupRoster());
await ssQAPage.locator('#nav-pickup').click();
await ssQAPage.waitForFunction(()=>document.querySelector('.btn-detail')&&!document.querySelector('.btn-detail').disabled);
efficiencyAssert.equal(efficiencyRosterReads,1);
await ssQAPage.evaluate(()=>{_allStudentsCache.at=Date.now()-4*60*1000;});
await ssQAPage.locator('.btn-detail').first().click();
await ssQAPage.waitForFunction(()=>!document.querySelector('#rd-save').disabled);
efficiencyAssert.equal(efficiencyRosterReads,1,'4-minute roster reused: no blocking roster request');
efficiencyAssert.match(await ssQAPage.locator('#rd-modal-overlay').textContent(),/5 分鐘內共用/);
await ssQAPage.locator('.rd-stu').first().check();
await ssQAPage.screenshot({path:'/home/user/workspace/frontend_qa/roster-efficient-desktop.png'});
await ssQAPage.setViewportSize({width:375,height:812});
await ssQAPage.screenshot({path:'/home/user/workspace/frontend_qa/roster-efficient-mobile.png'});
efficiencyAssert.equal(await ssQAPage.evaluate(()=>document.querySelector('#rd-modal-overlay').scrollWidth<=innerWidth),true);
await ssQAPage.clock.install();
await ssQAPage.clock.fastForward(121000); // Roster now exceeds 5 minutes.
await ssQAPage.locator('#rd-add-school').click();
await ssQAPage.locator('#rd-school-picker [data-name="新高國小"]').click();
efficiencyAssert.equal(efficiencyRosterReads,1,'modal renders from its pinned roster, not one read per school');
efficiencyRows[0]['學期狀態']='離校';
await ssQAPage.locator('#rd-save').click();
await ssQAPage.locator('#rd-retry').waitFor();
efficiencyAssert.equal(efficiencyRosterReads,2,'expired roster revalidated once before save');
efficiencyAssert.equal(efficiencyWrites,0,'changed enrollment blocks sending');
efficiencyAssert.match(await ssQAPage.locator('#rd-body').textContent(),/尚未送出儲存/);
await ssQAPage.locator('#rd-close').click();
efficiencyRows[0]['學期狀態']='在學';

// A failed roster read must not force the successful detail request to repeat.
efficiencyMode='hold-roster';
await ssQAPage.evaluate(()=>{invalidatePickupRoster();void openRouteDetailsModal({
  row_index:206,date:'2026-09-21',route:'合成交通車A',noon_school:'新光國小'
},'noon',{});});
await ssQAPage.waitForFunction(()=>document.querySelector('#rd-body').textContent.includes('已排學生已完成'));
efficiencyAssert.equal(efficiencyHeld.length,1);
await efficiencyHeld.shift().fulfill({status:503,body:'synthetic roster outage'});
await ssQAPage.locator('#rd-retry').waitFor();
var detailsBeforeRetry=efficiencyDetailReads;
efficiencyMode='ok';
await ssQAPage.locator('#rd-retry').click();
await ssQAPage.waitForFunction(()=>!document.querySelector('#rd-save').disabled);
efficiencyAssert.equal(efficiencyDetailReads,detailsBeforeRetry,'retry only failed roster, reuse <30s successful details');
await ssQAPage.locator('#rd-close').click();

// Conversely, details failing must not invalidate a valid roster.
efficiencyMode='details-fail';
await ssQAPage.locator('.btn-detail').first().click();
await ssQAPage.locator('#rd-retry').waitFor();
var rosterBeforeRetry=efficiencyRosterReads;
efficiencyMode='ok';
await ssQAPage.locator('#rd-retry').click();
await ssQAPage.waitForFunction(()=>!document.querySelector('#rd-save').disabled);
efficiencyAssert.equal(efficiencyRosterReads,rosterBeforeRetry);
await ssQAPage.locator('.rd-stu').first().check();
await ssQAPage.evaluate(()=>invalidatePickupRoster());
await ssQAPage.locator('#rd-save').click();
await ssQAPage.locator('#rd-retry').waitFor();
efficiencyAssert.equal(efficiencyWrites,0,'invalidated roster blocks stale modal writes');
efficiencyAssert.deepEqual(ssQAErrors,[]);
console.log({rosterEfficiency:'PASS',checks:['4min reuse','modal pinned snapshot','5min save-time refresh',
  'changed enrollment blocks write','retry only failed source','student-change invalidation','mobile fits'],productionWrites:0});
