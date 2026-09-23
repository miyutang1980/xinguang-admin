var recoveryAssert=(await import('node:assert/strict')).default;
var recoveryNetwork=[];
await ssQAContext.route('**/*',async route=>{
 const url=new URL(route.request().url()),action=url.searchParams.get('_action');
 if(url.hostname==='miyutang.app.n8n.cloud'&&url.pathname.includes('route-schedule'))throw new Error('Wrong n8n source');
 if(!action?.startsWith('routes_'))return route.fallback();
 recoveryNetwork.push({action,semester:url.searchParams.get('semester')});
 if(action==='routes_list')return route.fulfill({json:{success:true,rows:[
  {row_index:2,date:'2026-06-23',route:'歷史路線不可見'},
  ...Array.from({length:10},(_,i)=>({row_index:201+i,date:i<5?'2026-08-31':'2026-09-21',route:'合成路線'+(i%5),status:'上'})),
  {row_index:211,date:'2027-02-01',route:'未來路線不可見'}
 ]}});
 if(action==='routes_get_staff')return route.fulfill({json:{success:true,staff:[{name:'合成老師'}]}});
 return route.fulfill({json:{success:true,details:{}}});
});
await ssQAPage.locator('#nav-pickup').click();
await ssQAPage.waitForFunction(()=>document.querySelector('#ps-status')?.textContent.includes('共 10 列'));
await ssQAPage.locator('#ps-filter-week').selectOption('');
var tableText=await ssQAPage.locator('#ps-table-wrap').textContent();
recoveryAssert.ok(tableText.includes('2026-08-31')&&tableText.includes('2026-09-21'));
recoveryAssert.ok(!tableText.includes('不可見'));
recoveryAssert.equal(recoveryNetwork.find(r=>r.action==='routes_list').semester,'115-1');
recoveryAssert.ok(await ssQAPage.locator('#ps-table-wrap [data-noon-cell="201"]').count());
recoveryAssert.ok(await ssQAPage.locator('#ps-table-wrap [data-noon-cell="210"]').count());
await ssQAPage.screenshot({path:'/home/user/workspace/frontend_qa/pickup-term-desktop.png'});
await ssQAPage.setViewportSize({width:390,height:844});
await ssQAPage.locator('#mainContent h2').scrollIntoViewIfNeeded();
await ssQAPage.screenshot({path:'/home/user/workspace/frontend_qa/pickup-term-mobile.png'});
// Deterministic cache race check: old generations cannot repopulate cache.
var race=await ssQAPage.evaluate(async()=>{
 const original=_pickupCallGateway;const pending=[];
 try {
  _pickupCallGateway=()=>new Promise(resolve=>pending.push(resolve));
  invalidatePickupScheduleReads();
  const a=_pickupScheduleRead('test',{action:'list'}).catch(e=>e.message);
  invalidatePickupScheduleReads();
  const b=_pickupScheduleRead('test',{action:'list'});
  pending[0]({success:true,rows:[]});
  const stale=await a;
  pending[1]({success:true,rows:[{row_index:201,date:'2026-09-21',route:'fresh'}]});
  const fresh=await b;fresh.rows[0].route='mutated-client';
  const cached=await _pickupScheduleRead('test',{action:'list'});
  return {stale,route:cached.rows[0].route,requests:pending.length};
 }finally{_pickupCallGateway=original;invalidatePickupScheduleReads();}
});
recoveryAssert.ok(race.stale.includes('失效'));
recoveryAssert.equal(race.route,'fresh');recoveryAssert.equal(race.requests,2);
recoveryAssert.equal(ssQAErrors.length,0);
console.log({pickupRecovery:'PASS',checks:['ten rows across two dates','history/future hidden','original row IDs preserved','Gateway-only read','stale invalidation rejected','cache copies isolated'],productionWrites:0});
