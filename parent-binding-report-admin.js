/* Read-only, same-origin frame. No credentials in URLs, messages or new storage. */
(function(){
  'use strict';
  let frame=null;
  const allowed=()=>typeof _loggedIn!=='undefined'&&_loggedIn&&_currentUser&&
    ['admin','staff'].includes(String(_currentUser.role).toLowerCase());
  function reset(){
    if(frame){try{frame.contentWindow.clearBindingReport?.();}catch(_){}frame.remove();frame=null;}
  }
  window.XGParentBindingReport={
    allowed,reset,
    session(child){
      if(!allowed()||currentSection!=='parentBindingReport'||!frame?.isConnected||frame.contentWindow!==child)return null;
      const adminPass=sessionStorage.getItem('xg_session_pwd')||'';
      if(!adminPass)return null;
      return {adminUser:_currentUser.username,adminPass,epoch:_authEpoch};
    }
  };
  window.renderParentBindingReport=function(container){
    reset();container.replaceChildren();
    if(!allowed()){container.textContent='僅限管理員或行政查閱家長資料。';return;}
    frame=document.createElement('iframe');frame.id='parentBindingReportFrame';
    frame.title='115-1 學生與家長綁定唯讀對照';frame.referrerPolicy='no-referrer';
    frame.style.cssText='width:100%;height:calc(100dvh - 140px);min-height:480px;border:0;display:block';
    frame.src='./parent-binding-report/?embed=1&v=20261001-1';
    container.append(frame);
  };
  window.addEventListener('pagehide',reset);
})();
