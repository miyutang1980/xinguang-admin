var pickupAssert=(await import('node:assert/strict')).default;
await pickupPendingClasses[0].abort('timedout');
await new Promise((resolve,reject)=>{var start=Date.now(),timer=setInterval(()=>{if(pickupPendingClasses.length===2){clearInterval(timer);resolve();}else if(Date.now()-start>3000){clearInterval(timer);reject(new Error('Missing legacy read retry'));}},20);});
await pickupPendingClasses[1].abort('timedout');
await ssQAPage.waitForTimeout(150);
pickupAssert.equal(await ssQAPage.locator('#ps-add-day').count(),1,'stale classes timeout must not erase pickup');
pickupAssert.equal(await ssQAPage.locator('#mainContent').textContent().then(s=>s.includes('讀取班別設定失敗')),false);
// Late success is also forbidden from painting a new page.
await ssQAPage.locator('#nav-classes').click();
await ssQAPage.waitForFunction(()=>document.querySelector('#mainContent').textContent.includes('讀取班別設定中'));
await ssQAPage.locator('#nav-classAssign').click();
await ssQAPage.locator('#ssSemester').waitFor();
await pickupPendingClasses[2].fulfill({json:{success:true,data:await ssQAPage.evaluate(()=>SEMESTER_CLASSES)}});
await ssQAPage.waitForTimeout(100);
pickupAssert.equal(await ssQAPage.locator('#ssSemester').count(),1);
// Queued lazy/account renderers cannot overwrite a later navigation.
await ssQAPage.evaluate(()=>{switchSection('accounts');switchSection('students');});
await ssQAPage.locator('#ssMasterSearch').waitFor();
await ssQAPage.waitForTimeout(150);
pickupAssert.equal(await ssQAPage.locator('#ssMasterSearch').count(),1);
await ssQAPage.locator('#nav-pickup').click();
await ssQAPage.waitForFunction(()=>document.querySelector('#ps-status')?.textContent.includes('最後同步'));
ssQAPage.removeAllListeners('dialog');
ssQAPage.on('dialog',async d=>{
 pickupDialogs.push(d.message());
 if(d.type()==='prompt') await d.accept(pickupAnswers.shift()??'');
 else await d.accept();
});
var pickupAnswers=['2026-09-23'];
await ssQAPage.locator('#ps-add-day').click();
await ssQAPage.waitForFunction(()=>document.querySelector('#ps-status')?.textContent.includes('共 5 列'));
pickupAssert.equal(pickupRows.length,5);
pickupAssert.equal(pickupRows[2].route,'交通車B (RGE-2523)');
pickupAssert.equal(pickupRows[3].route,'交通車C (RGE-2522)');
pickupAssert.equal(await ssQAPage.evaluate(()=>_routeNameForDate('交通車B (RDW-1655)','2026-06-01')),'交通車B (RDW-1655)');
pickupAssert.equal(await ssQAPage.evaluate(()=>_routeNameForDate('交通車C (RFD-9763)','2026-09-28')),'交通車C (RGE-2522)');
pickupAssert.equal(pickupCalls.filter(a=>a==='routes_append').length,1);
pickupAnswers=['2026-09-23'];
await ssQAPage.locator('#ps-add-day').click();
pickupAssert.equal(pickupCalls.filter(a=>a==='routes_append').length,1,'duplicate date blocked');
pickupAnswers=['2026-09-23','離線臨時路線','3'];
await ssQAPage.locator('#ps-add-route').click();
await ssQAPage.waitForFunction(()=>document.querySelector('#ps-status')?.textContent.includes('共 6 列'));
pickupAssert.equal(pickupRows.length,6);
pickupAssert.equal(pickupRows[5].route,'離線臨時路線');
await ssQAPage.screenshot({path:'/home/user/workspace/frontend_qa/pickup-fixed-desktop.png'});
// Failed read must not be presented as an empty editable schedule; retry must work.
pickupMode='list-error';
await ssQAPage.locator('#ps-refresh').click();
await ssQAPage.waitForFunction(()=>document.querySelector('#ps-status')?.textContent.includes('載入失敗'));
pickupAssert.equal(await ssQAPage.locator('#ps-add-day').isDisabled(),true);
pickupMode='ok';
await ssQAPage.locator('#ps-table-wrap button').click();
await ssQAPage.waitForFunction(()=>document.querySelector('#ps-status')?.textContent.includes('共 6 列'));
pickupAssert.equal(await ssQAPage.locator('#ps-add-day').isEnabled(),true);
// Copying historical B/C routes into this semester changes only destination plates.
pickupRows.push(
 {row_index:20,date:'2026-06-01',weekday:'週一',status:'上',route:'交通車B (RDW-1655)',capacity:'4'},
 {row_index:21,date:'2026-06-01',weekday:'週一',status:'上',route:'交通車C (RFD-9763)',capacity:'4'}
);
await ssQAPage.locator('#ps-refresh').click();
await ssQAPage.waitForFunction(()=>document.querySelector('#ps-status')?.textContent.includes('共 8 列'));
pickupAnswers=['2026-06-01','2026-09-28'];
await ssQAPage.locator('#ps-copy-week').click();
await ssQAPage.waitForFunction(()=>document.querySelector('#ps-status')?.textContent.includes('共 10 列'));
pickupAssert.equal(pickupRows.find(r=>r.date==='2026-06-01'&&r.row_index===20).route,'交通車B (RDW-1655)');
pickupAssert.deepEqual(pickupRows.filter(r=>r.date==='2026-09-28').map(r=>r.route),['交通車B (RGE-2523)','交通車C (RGE-2522)']);
// Unconfirmed write cannot be auto-retried, even on a repeated button click.
pickupMode='hold-append';
pickupAnswers=['2026-09-24','不確定測試路線',''];
await ssQAPage.locator('#ps-add-route').click();
await ssQAPage.waitForFunction(()=>document.querySelector('#ps-status')?.textContent.includes('新增中'));
pickupAssert.equal(await ssQAPage.locator('#ps-add-route').isDisabled(),true);
await ssQAPage.evaluate(()=>document.querySelector('#ps-add-route').click());
await pickupAppendPending.abort('timedout');
await ssQAPage.waitForFunction(()=>document.querySelector('#ps-status')?.textContent.includes('無法確認新增結果'));
pickupAssert.equal(await ssQAPage.locator('#ps-add-route').isDisabled(),true);
pickupAssert.equal(pickupCalls.filter(a=>a==='routes_append').length,4);
pickupAssert.equal(ssQAErrors.length,0);
console.log({pickupRegression:'PASS',checks:['exact screenshot reproduction fixed','late classes success ignored','account/lazy navigation isolation','add day creates 5 routes','duplicate date blocked','add single route','load error disables create','retry works','uncertain write not repeated'],syntheticRows:pickupRows.length,productionWrites:0});
