// This file runs after bootstrap.js; every HTTP request is intercepted.
var stAssert=(await import('node:assert/strict')).default, stChecks=[];
var stCheck=(name,pass)=>{stAssert.ok(pass,name);stChecks.push(name);};
await ssQAPage.clock.install();
await ssQAPage.evaluate(({assignments,masters})=>{
  const original=window.fetch;
  const demo={mode:'hold',resolvers:[],calls:[],assignments,masters};
  window.__semesterTimeout=demo;
  window.fetch=(url,options)=>{
    const u=new URL(String(url)),action=u.searchParams.get('_action');
    if(!['semester_assignments_list','student_master_list','student_master_update'].includes(action))return original(url,options);
    demo.calls.push(action);
    if(action==='student_master_update'){
      if(demo.mode==='write-hang')return new Promise(resolve=>demo.resolvers.push(resolve)); // deliberately ignores abort
      return Promise.resolve({ok:true,json:async()=>({success:false,error:'No test write expected'})});
    }
    const semester=u.searchParams.get('semester')||'115-1';
    const result={success:true,active:true,currentSemester:'115-1',semester,semesters:['115-1','114-2'],
      displayNameSource:'學期班級指派!W',list:action==='student_master_list'?demo.masters:
        demo.assignments.filter(r=>r['學期']===semester).map(r=>({...r,'學生中英文姓名':r['學生姓名']+'（Demo）'}))};
    if(demo.mode==='json-hang')return Promise.resolve({ok:true,json:()=>new Promise(resolve=>demo.resolvers.push(()=>resolve(result)))});
    if(demo.mode==='hold')return new Promise(resolve=>demo.resolvers.push(()=>resolve({ok:true,json:async()=>result})));
    if(demo.mode==='denied')return Promise.resolve({ok:true,json:async()=>({success:false,error:'帳號已停用'})});
    return Promise.resolve({ok:true,json:async()=>result});
  };
},{assignments:ssQAAssignments,masters:ssQAMasters});
await ssQAPage.locator('#nav-classAssign').click();
await ssQAPage.waitForFunction(()=>window.__semesterTimeout.calls.includes('semester_assignments_list'));
await ssQAPage.clock.fastForward(15001);
stCheck('slow service explanation after 15s',(await ssQAPage.locator('#mainContent').innerText()).includes('服務回應較慢'));
await ssQAPage.clock.fastForward(7000);
stCheck('22s response not prematurely failed',!(await ssQAPage.locator('#mainContent').innerText()).includes('無法載入'));
stCheck('no edits exposed during slow read',await ssQAPage.locator('#ssAdd').count()===0);
stCheck('one pending read, no automatic retry',await ssQAPage.evaluate(()=>__semesterTimeout.calls.filter(x=>x==='semester_assignments_list').length)===1);
await ssQAPage.clock.fastForward(7000);
await ssQAPage.screenshot({path:'/home/user/workspace/transport-audit-20260924/semester-slow-desktop.png'});
await ssQAPage.evaluate(()=>__semesterTimeout.resolvers.shift()());
await ssQAPage.locator('#ssSemester').waitFor();
stCheck('late 29s success renders validated roster',(await ssQAPage.locator('#ssNotice').innerText()).includes('顯示 2 / 2 筆'));
await ssQAPage.locator('#ssRefresh').click();
await ssQAPage.clock.fastForward(45001);
await ssQAPage.locator('#ssRetry').waitFor();
stCheck('read deadline works when fetch ignores abort',(await ssQAPage.locator('.ss-error').innerText()).includes('45 秒'));
stCheck('timeout cannot expose empty roster or write controls',await ssQAPage.locator('#ssAdd').count()===0&&await ssQAPage.locator('#ssList').count()===0);
await ssQAPage.evaluate(()=>__semesterTimeout.resolvers.shift()());
stCheck('late response after deadline ignored',await ssQAPage.locator('#ssRetry').count()===1);
await ssQAPage.evaluate(()=>__semesterTimeout.mode='ok');
await ssQAPage.locator('#ssRetry').click();
await ssQAPage.locator('#ssSemester').waitFor();
stCheck('explicit retry recovers roster',await ssQAPage.locator('.ss-row-edit').count()===2);
await ssQAPage.evaluate(()=>__semesterTimeout.mode='json-hang');
await ssQAPage.locator('#ssRefresh').click();
await ssQAPage.clock.fastForward(45001);
await ssQAPage.locator('#ssRetry').waitFor();
stCheck('JSON parsing is also bounded',await ssQAPage.locator('#ssRetry').count()===1);
await ssQAPage.evaluate(()=>{__semesterTimeout.resolvers.shift()();__semesterTimeout.mode='denied';});
await ssQAPage.locator('#ssRetry').click();
await ssQAPage.waitForFunction(()=>document.querySelector('.ss-error')?.textContent.includes('帳號已停用'));
stCheck('auth errors remain blocked',await ssQAPage.locator('#ssAdd').count()===0);
await ssQAPage.evaluate(()=>__semesterTimeout.mode='hold');
await ssQAPage.locator('#ssRetry').click();
await ssQAPage.locator('#nav-dashboard').click();
var stBefore=await ssQAPage.locator('#mainContent').innerHTML();
await ssQAPage.evaluate(()=>__semesterTimeout.resolvers.shift()());
stCheck('late response cannot replace other page',await ssQAPage.locator('#mainContent').innerHTML()===stBefore);
await ssQAPage.evaluate(()=>{__semesterTimeout.mode='ok';XGStudents.invalidateViewCache();});
await ssQAPage.locator('#nav-students').click();await ssQAPage.locator('#ssMasterSearch').waitFor();
await ssQAPage.locator('.ss-row-edit').first().click();
await ssQAPage.locator('#ssEditor [data-field="英文名"]').fill('Changed demo');
await ssQAPage.evaluate(()=>__semesterTimeout.mode='write-hang');
await ssQAPage.locator('#ssSave').click();
await ssQAPage.waitForFunction(()=>__semesterTimeout.calls.includes('student_master_update'));
await ssQAPage.clock.fastForward(45001);
await ssQAPage.waitForFunction(()=>document.getElementById('ssSave')?.textContent.includes('核對結果'));
stCheck('write timeout blocks direct resubmission',await ssQAPage.locator('#ssSave').isDisabled());
stCheck('write issued once, never retried',await ssQAPage.evaluate(()=>__semesterTimeout.calls.filter(x=>x==='student_master_update').length)===1);
await ssQAPage.locator('#ssCancel').click();
await ssQAPage.evaluate(()=>{__semesterTimeout.mode='hold';XGStudents.invalidateViewCache();});
await ssQAPage.locator('#nav-classAssign').click();
await ssQAPage.setViewportSize({width:375,height:812});await ssQAPage.clock.fastForward(16000);
await ssQAPage.screenshot({path:'/home/user/workspace/transport-audit-20260924/semester-slow-mobile.png'});
stCheck('mobile no horizontal overflow',await ssQAPage.evaluate(()=>document.documentElement.scrollWidth<=innerWidth));
await ssQAPage.clock.fastForward(30000);await ssQAPage.locator('#ssRetry').waitFor();
stCheck('diagnostics do not include credentials',!(await ssQAPage.evaluate(()=>JSON.stringify(XGStudents.diagnostics()))).includes('offline-only'));
stCheck('no JavaScript errors',ssQAErrors.length===0);
await ssQAFs.writeFile('/home/user/workspace/transport-audit-20260924/semester-timeout-results.json',JSON.stringify({passed:stChecks.length,checks:stChecks,errors:ssQAErrors},null,2));
console.log({passed:stChecks.length,errors:ssQAErrors});
