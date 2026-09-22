// Run bootstrap.js first. Every network request remains intercepted.
var loginAssert = (await import('node:assert/strict')).default;
var authRoutes=[],authFallbackRoutes=[],authGatewayMode='hold';
await ssQAContext.route('**/*',async route=>{
 const req=route.request(), url=new URL(req.url());
 if(url.searchParams.get('_action')==='auth_login' || (req.method()==='POST' && req.postDataJSON()?._action==='auth_login')) {
  authRoutes.push(route);
  if(authGatewayMode==='abort'){setTimeout(()=>route.abort('timedout').catch(()=>{}),100);}
  return;
 }
 if(url.hostname==='miyutang.app.n8n.cloud' && url.pathname==='/webhook/auth-login') {
  authFallbackRoutes.push(route); return;
 }
 return route.fallback();
});
await ssQAContext.addInitScript(()=>{
 window.__XG_AUTH_GATEWAY_TIMEOUT_MS=80;
 window.__XG_AUTH_FALLBACK_TIMEOUT_MS=300;
 localStorage.setItem('xg_remember',JSON.stringify({username:'qa-login',password:'synthetic-only',expires:Date.now()+60000}));
});
var authOK={success:true,user:{username:'qa-login',name:'離線登入測試',role:'admin',permissions:['dashboard','students']}};
await ssQAPage.reload({waitUntil:'load'});
loginAssert.equal(await ssQAPage.evaluate(()=>AUTH_GATEWAY_TIMEOUT_MS),80);
await ssQAPage.waitForFunction(()=>!!_loginFlight);
await ssQAPage.evaluate(()=>{doLogin();doLogin();checkSession();});
await ssQAPage.waitForTimeout(450); // Covers both former 200ms/300ms startup timers.
loginAssert.equal(authRoutes.length,1,'Auto-login and repeated clicks must share one request');
loginAssert.equal(await ssQAPage.locator('.login-btn').isDisabled(),true);
await authRoutes[0].fulfill({json:authOK});
await ssQAPage.locator('#appScreen').waitFor({state:'visible'});
await ssQAPage.locator('#nav-students').click();
await ssQAPage.locator('#ssMasterSearch').waitFor();
loginAssert.match(await ssQAPage.locator('#pageTitle').textContent(),/學生管理/);
loginAssert.match(await ssQAPage.locator('#mainContent').innerText(),/測試用匿名記錄1/);
await ssQAPage.evaluate(()=>{initApp();doLogin();checkSession();});
loginAssert.equal(await ssQAPage.evaluate(()=>currentSection),'students');
await ssQAPage.locator('#nav-classAssign').click();
await ssQAPage.locator('#ssSemester').waitFor();
loginAssert.match(await ssQAPage.locator('#pageTitle').textContent(),/學期班級指派/);
loginAssert.match(await ssQAPage.locator('#mainContent').innerText(),/Level-A/);
await ssQAPage.locator('#ssSemester').selectOption('114-2');
await ssQAPage.waitForFunction(()=>document.querySelector('#mainContent').innerText.includes('History-A'));
await ssQAPage.waitForTimeout(1000); // Automatic post-login sync must not reset the chosen page.
loginAssert.equal(await ssQAPage.evaluate(()=>currentSection),'classAssign');
loginAssert.equal(authRoutes.length,1);
await ssQAPage.screenshot({path:'/home/user/workspace/frontend_qa/login-fixed-desktop.png'});
await ssQAPage.setViewportSize({width:390,height:844});
await ssQAPage.screenshot({path:'/home/user/workspace/frontend_qa/login-fixed-mobile.png'});
await ssQAPage.setViewportSize({width:1440,height:960});
// Logout invalidates an unresolved attempt. Its response must not restore a session.
await ssQAPage.evaluate(()=>{doLogout();document.querySelector('#loginUser').value='qa-old';document.querySelector('#loginPass').value='synthetic-only';doLogin();});
await ssQAPage.waitForFunction(()=>!!_loginFlight);
await ssQAPage.evaluate(()=>doLogout());
await authRoutes[1].fulfill({json:authOK});
await ssQAPage.waitForTimeout(100);
loginAssert.equal(await ssQAPage.evaluate(()=>_loggedIn),false);
loginAssert.equal(await ssQAPage.locator('#loginUser').inputValue(),'');
// A rejected attempt releases the single-flight lock, so a manual retry can succeed.
await ssQAPage.evaluate(()=>{document.querySelector('#loginUser').value='qa-login';document.querySelector('#loginPass').value='synthetic-only';doLogin();});
await ssQAPage.waitForFunction(()=>!!_loginFlight);
await authRoutes[2].fulfill({json:{success:false,error:'測試拒絕登入'}});
await ssQAPage.waitForFunction(()=>!_loginFlight);
loginAssert.equal(await ssQAPage.evaluate(()=>_loggedIn),false);
await ssQAPage.evaluate(()=>{doLogin();doLogin();});
await ssQAPage.waitForFunction(()=>!!_loginFlight);
await authRoutes[3].fulfill({json:authOK});
await ssQAPage.locator('#appScreen').waitFor({state:'visible'});
await ssQAPage.locator('#nav-classAssign').click();
await ssQAPage.locator('#ssSemester').waitFor();
loginAssert.equal(authRoutes.length,4);
loginAssert.equal(authFallbackRoutes.length,0);
// A stalled Gateway must release to the already-supported n8n fallback quickly.
authGatewayMode='abort';
await ssQAPage.evaluate(()=>{doLogout();document.querySelector('#loginUser').value='qa-login';document.querySelector('#loginPass').value='synthetic-only';doLogin();});
for(var fallbackWait=0;authFallbackRoutes.length<1&&fallbackWait<50;fallbackWait++)await ssQAPage.waitForTimeout(20);
loginAssert.equal(authFallbackRoutes.length,1);
await authFallbackRoutes[0].fulfill({json:authOK});
await ssQAPage.locator('#appScreen').waitFor({state:'visible'});
loginAssert.equal(await ssQAPage.evaluate(()=>_loggedIn),true);
loginAssert.equal(await ssQAPage.locator('.login-btn').isDisabled(),false);
loginAssert.equal(ssQAErrors.length,0);
console.log({loginRegression:'PASS',checks:['single-flight automatic/manual login','master route + records','assignment route + records','semester history','repeat initialization preserves route','background sync preserves route','logout ignores stale response','failed login permits retry','5s Gateway deadline releases fallback'],authRequests:authRoutes.length,productionWrites:0});
