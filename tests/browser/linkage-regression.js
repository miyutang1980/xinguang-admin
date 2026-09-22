var linkageAssert=(await import('node:assert/strict')).default;
var linkedRequests=0;
var makeLinked=(id,school,status='在學',term='115-1')=>({
  ...ssQAAssignments[0],_row:Number(id.slice(2))+1,'學生編號':id,'學生姓名':'合成學生'+id,
  '學校':school,'學期狀態':status,'學期':term,'指派編號':term+'|'+id,
  '學生中英文姓名':'合成學生'+id+'（English '+id+'）'
});
var linkedRows=[
  makeLinked('QA1','新光國小'),makeLinked('QA2',' 文心國小 '),
  makeLinked('QA3','新光國小'),makeLinked('QA4','新光國中'),
  makeLinked('QA5','幼兒園'),makeLinked('QA6','離校國小','離校'),
  makeLinked('QA7','歷史國小','在學','114-2')
];
await ssQAContext.route('**/*',route=>{
  const url=new URL(route.request().url());
  if(url.searchParams.get('_action')!=='semester_assignments_list')return route.fallback();
  linkedRequests++;
  const semester=url.searchParams.get('semester');
  return route.fulfill({json:{success:true,active:true,currentSemester:'115-1',semester,semesters:['115-1','114-2'],
    displayNameSource:semester==='115-1'?'學期班級指派!W':'',list:linkedRows.filter(r=>r['學期']===semester)}});
});
await ssQAPage.evaluate(()=>localStorage.setItem('xg_custom_schools',JSON.stringify(['新光國小','舊自訂國小'])));
pickupRows.push({row_index:206,date:'2026-09-21',weekday:'週一',route:'合成接送',status:'上',noon_school:'舊國小'});
var masterBefore=ssQACalls.filter(c=>c.action==='student_master_list').length;
await ssQAPage.locator('#ps-refresh').click();
await ssQAPage.waitForFunction(()=>document.querySelector('#ps-status').textContent.includes('共 1 列'));
linkageAssert.deepEqual((await ssQAPage.evaluate(()=>SCHOOLS_LIST.map(s=>s.name))).sort(),['文心國小','新光國小']);
linkageAssert.equal(ssQACalls.filter(c=>c.action==='student_master_list').length,masterBefore,'W-mode makes no extra master request');
var options=await ssQAPage.locator('select[data-field="noon_school"]').first().locator('option').evaluateAll(options=>options.map(o=>({value:o.value,disabled:o.disabled,selected:o.selected})));
linkageAssert.equal(options.filter(o=>o.value==='新光國小').length,1);
linkageAssert.equal(options.find(o=>o.value==='舊國小').disabled,true);
linkageAssert.equal(options.find(o=>o.value==='舊國小').selected,true,'historical value preserved without offering it as a new school');
await ssQAPage.evaluate(()=>openRouteDetailsModal({row_index:206,date:'2026-09-21',weekday:'週一',route:'合成接送',noon_school:'新光國小'},'noon',{}));
linkageAssert.equal(await ssQAPage.locator('.rd-stu').count(),2);
linkageAssert.match(await ssQAPage.locator('#rd-body').textContent(),/English QA1/);
await ssQAPage.locator('.rd-stu').first().check();
await ssQAPage.locator('#rd-student-search').fill('english qa3');
linkageAssert.equal(await ssQAPage.locator('.rd-stu:visible').count(),1);
linkageAssert.match(await ssQAPage.locator('#rd-summary').textContent(),/合計 1 人/);
await ssQAPage.locator('#rd-student-search').fill('查無此人');
linkageAssert.equal(await ssQAPage.locator('.rd-stu:visible').count(),0);
linkageAssert.match(await ssQAPage.locator('#rd-search-status').textContent(),/顯示 0/);
await ssQAPage.locator('#rd-student-search').fill('');
await ssQAPage.locator('#rd-add-school').click();
linkageAssert.deepEqual(await ssQAPage.locator('#rd-school-picker [data-name]').evaluateAll(nodes=>nodes.map(n=>n.dataset.name)),['文心國小']);
await ssQAPage.locator('#rd-school-picker [data-name]').click();
await ssQAPage.waitForFunction(()=>document.querySelectorAll('.rd-stu').length===3);
await ssQAPage.screenshot({path:'/home/user/workspace/frontend_qa/linkage-desktop.png'});
await ssQAPage.setViewportSize({width:375,height:812});
await ssQAPage.screenshot({path:'/home/user/workspace/frontend_qa/linkage-mobile.png'});
await ssQAPage.locator('#rd-close').click();
await ssQAPage.locator('#nav-classAssign').click();
await ssQAPage.locator('#ssSearch').fill('english qa3');
linkageAssert.match(await ssQAPage.locator('#ssList').textContent(),/English QA3/);
linkageAssert.doesNotMatch(await ssQAPage.locator('#ssList').textContent(),/English QA1/);
// Invalidation drops both schools and display names, forcing a fresh semester read.
await ssQAPage.evaluate(()=>invalidatePickupRoster());
linkageAssert.deepEqual(await ssQAPage.evaluate(()=>SCHOOLS_LIST),[]);
linkedRows[0]['學生中英文姓名']='合成學生QA1（Renamed）';
await ssQAPage.evaluate(()=>_loadAllStudentsOnce({}));
linkageAssert.equal(await ssQAPage.evaluate(()=>_allStudentsCache.list.find(s=>s.student_id==='QA1').display_name),'合成學生QA1（Renamed）');
linkageAssert.deepEqual(ssQAErrors,[]);
console.log({linkageRegression:'PASS',checks:['current active elementary schools only','deduplicated trimmed schools','no legacy custom list','saved old school retained readonly','one semester source without master fetch','Chinese/English search keeps selected IDs','assignment search uses W','cache refresh picks updated names'],productionWrites:0});
