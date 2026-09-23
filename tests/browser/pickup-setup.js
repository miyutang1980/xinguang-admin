var pickupCalls=[], pickupPendingClasses=[], pickupRows=[], pickupDialogs=[];
var pickupMode='ok', pickupAppendPending=null;
await ssQAContext.route('**/*',async route=>{
 const req=route.request(),url=new URL(req.url());
 let action=url.searchParams.get('_action');
 if(url.hostname==='miyutang.app.n8n.cloud'&&url.pathname.includes('/admin/route-schedule')){
  const direct=req.postDataJSON()?.action;
  if(direct==='list')action='routes_list';
  if(direct==='get_staff')action='routes_get_staff';
 }
 if(url.searchParams.get('action')==='classes') {pickupPendingClasses.push(route);return;}
 if(!action?.startsWith('routes_')) return route.fallback();
 pickupCalls.push(action);
 if(action==='routes_get_staff') return route.fulfill({json:{success:true,staff:[{name:'測試老師'}]}});
 if(action==='routes_list') return route.fulfill({json:pickupMode==='list-error'?{success:false,error:'測試讀取失敗'}:{success:true,rows:pickupRows}});
 if(action==='routes_get_details') return route.fulfill({json:{success:true,details:[],total_count:0}});
 if(action==='routes_append') {
  if(pickupMode==='hold-append') {pickupAppendPending=route;return;}
  const rows=JSON.parse(url.searchParams.get('rows_json'));
  const firstIndex=pickupRows.length+2;
  pickupRows.push(...rows.map((r,i)=>({...r.fields,row_index:firstIndex+i})));
  return route.fulfill({json:{success:true,appended:rows.length}});
 }
 return route.fulfill({json:{success:true,details:{}}});
});
await ssQAPage.evaluate(()=>{_currentUser={username:'offline-qa',role:'admin',permissions:ALL_PERMS};applyPermissions();window._classesLoaded=false;});
await ssQAPage.locator('#nav-classes').click();
await ssQAPage.waitForFunction(()=>document.querySelector('#mainContent').textContent.includes('讀取班別設定中'));
await ssQAPage.locator('#nav-pickup').click();
await ssQAPage.locator('#ps-add-day').waitFor();
await ssQAPage.waitForFunction(()=>document.querySelector('#ps-status')?.textContent.includes('最後同步'));
