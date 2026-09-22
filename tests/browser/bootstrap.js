var { chromium: ssQAChromium } = await import('playwright');
var ssQAFs = await import('node:fs/promises');
var ssQAPath = '/home/user/workspace/xinguang-admin/';
var ssQABrowser = await ssQAChromium.launch({headless:true});
var ssQAContext = await ssQABrowser.newContext({viewport:{width:1440,height:960}});
var ssQAPage = await ssQAContext.newPage();
var ssQAErrors = []; ssQAPage.on('pageerror',e=>ssQAErrors.push(e.message));
var ssQAAcceptConfirm = true; ssQAPage.on('dialog', d => ssQAAcceptConfirm ? d.accept() : d.dismiss());
var ssQAMasterHeaders=['學生編號','學生帳號','學生姓名','英文名','性別','生日','身分證字號','Email','住家地址','備註/飲食禁忌','媽媽姓名','媽媽手機','媽媽公司電話','媽媽工作單位','媽媽Email','爸爸姓名','爸爸手機','爸爸公司電話','爸爸工作單位','爸爸Email','媽媽LINE userId','爸爸LINE userId','建立時間','更新時間','弋果分校'];
var ssQAAssignmentHeaders=['指派編號','學期','學生編號','學生帳號','學生姓名','弋果班級','弋果課程','學生類別','學校','年級','小學班級','外籍教師','中籍教師','TXClass','學年度起始日','課堂時間','課程類別','教室','學期狀態','課後輔導','交通車','更新時間'];
var ssQABlank=h=>Object.fromEntries(h.map(k=>[k,'']));
var ssQAMasters=[1,2,3].map(n=>({...ssQABlank(ssQAMasterHeaders),_row:n+1,'學生編號':'QA'+n,'學生姓名':'測試用匿名記錄'+n,'學生帳號':'qa'+n,'弋果分校':'測試分校'}));
var ssQAAssignments=[1,2].map(n=>({...ssQABlank(ssQAAssignmentHeaders),_row:n+1,'指派編號':'115-1|QA'+n,'學期':'115-1','學生編號':'QA'+n,'學生姓名':'測試用匿名記錄'+n,'學生帳號':'qa'+n,'弋果班級':'CLASS-QA','TXClass':'Level-A','學期狀態':n===1?'在學':'離校','課後輔導':'托育','交通車':'Walk 走路'}));
ssQAAssignments.push({...ssQAAssignments[0],_row:4,'指派編號':'114-2|QA1','學期':'114-2','TXClass':'History-A'});
var ssQAMockMode='active', ssQACalls=[], ssQAMutations=[];
await ssQAContext.route('**/*',async route=>{
 const req=route.request(), url=new URL(req.url());
 if(url.hostname==='frontend.test') {
  let file=url.pathname==='/'?'index.html':url.pathname.slice(1);
  try { const body=await ssQAFs.readFile(ssQAPath+file); return route.fulfill({status:200,body,contentType:file.endsWith('.js')?'application/javascript':file.endsWith('.css')?'text/css':'text/html'}); } catch(e){return route.fulfill({status:404,body:'not found'});}
 }
 const action=url.searchParams.get('_action');
 if(!action) return route.fulfill({status:200,contentType:'application/json',body:JSON.stringify({success:false,error:'Offline QA only'})});
 const fields=JSON.parse(url.searchParams.get('fields_json')||'{}');
 const semester=url.searchParams.get('semester'), row=Number(url.searchParams.get('row_index'));
 ssQACalls.push({action,semester,fields,row});
 let result={success:false,error:'Unknown offline QA action'};
 if(action==='student_master_list') result={success:true,active:ssQAMockMode!=='inactive',currentSemester:'115-1',list:ssQAMasters};
 if(action==='semester_assignments_list') result={success:true,active:ssQAMockMode!=='inactive',currentSemester:'115-1',semester,semesters:['115-1','114-2'],list:ssQAAssignments.filter(r=>r['學期']===semester)};
 if(action==='student_master_create') {
  const id='QA'+(ssQAMasters.length+1); ssQAMasters.push({...ssQABlank(ssQAMasterHeaders),...fields,_row:ssQAMasters.length+2,'學生編號':id}); result={success:true,studentNo:id}; ssQAMutations.push(action);
 }
 if(action==='student_master_update') {Object.assign(ssQAMasters.find(r=>r._row===row),fields);result={success:true};ssQAMutations.push(action);}
 if(action==='semester_assignment_create') {
  const id=fields['學期']+'|'+fields['學生編號']; ssQAAssignments.push({...ssQABlank(ssQAAssignmentHeaders),...fields,_row:ssQAAssignments.length+2,'指派編號':id});result={success:true,assignmentNo:id};ssQAMutations.push(action);
 }
 if(action==='semester_assignment_update') {Object.assign(ssQAAssignments.find(r=>r._row===row),fields);result={success:true};ssQAMutations.push(action);}
 if(action==='semester_class_update') {
  const rows=ssQAAssignments.filter(r=>r['學期']===semester&&r['弋果班級']===url.searchParams.get('classCode'));rows.forEach(r=>Object.assign(r,fields));result={success:true,updated:rows.length};ssQAMutations.push(action);
 }
 if(ssQAMockMode==='missing-list' && action.endsWith('_list')) result={success:true};
 if(ssQAMockMode==='error' && action.endsWith('_list')) result={success:false,error:'Gateway update pending'};
 return route.fulfill({status:200,contentType:'application/json',body:JSON.stringify(result)});
});
await ssQAPage.goto('http://frontend.test/',{waitUntil:'load'});
await ssQAPage.evaluate(()=>{_currentUser={username:'offline-qa',role:'admin',permissions:ALL_PERMS};sessionStorage.setItem('xg_session_pwd','offline-only');document.getElementById('loginScreen').style.display='none';document.getElementById('appScreen').style.display='flex';});
await ssQAPage.locator('#nav-students').click();
await ssQAPage.locator('#ssMasterSearch').waitFor();
console.log({ready:await ssQAPage.locator('#ssNotice').textContent(),ssQAErrors});
