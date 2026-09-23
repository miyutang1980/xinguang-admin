'use strict';
const GW_URL = 'https://script.google.com/macros/s/AKfycbw-7_a_OfUVlgegcLxkux_9dr9UlYSVKhi3uQjV-0sr2X2TpRRmCXtSM7jbIqMHK4hNww/exec';
const $=id=>document.getElementById(id);
let credentials=null, roster=[], working=false;
const embedded = new URLSearchParams(location.search).get('embed') === '1';
const requests = new Set();
let alive = true, needsRefresh = false;
function parentSession() {
  try { return window.parent !== window && window.parent.XGParentBinding?.session(window); } catch (_) { return null; }
}
function activeSession() {
  const session = embedded ? parentSession() : credentials;
  if (!alive || !session) throw Error('後台登入已失效，請回到後台重新登入。');
  return session;
}
const el=(tag,value='',cls='')=>{const n=document.createElement(tag);n.textContent=value;if(cls)n.className=cls;return n;};
async function api(action,data={}){
  const session = activeSession();
  const controller=new AbortController(),timer=setTimeout(()=>controller.abort(),35000);
  requests.add(controller);
  try{
    const response=await fetch(GW_URL,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},
      body:JSON.stringify({...data,adminUser:session.adminUser,adminPass:session.adminPass,_action:action}),signal:controller.signal,cache:'no-store',referrerPolicy:'no-referrer'});
    const latest=activeSession();
    if(latest.adminUser!==session.adminUser || latest.epoch!==session.epoch)throw Error('登入已變更，已忽略舊回應。');
    if(!response.ok)throw Error('服務連線失敗，請稍後重新讀取。');
    const r=await response.json();
    const after=activeSession();
    if(after.adminUser!==session.adminUser || after.epoch!==session.epoch)throw Error('登入已變更，已忽略舊回應。');
    if(!r.success)throw Error(r.error||'操作失敗，請核對 Gateway 部署。');return r;
  }catch(e){if(e.name==='AbortError')throw Error('連線超時。請先重新讀取狀態，不要改選其他孩子重送。');throw e;}
  finally{clearTimeout(timer);requests.delete(controller);}
}
function setBusy(b){working=b;document.querySelectorAll('button,input,select,textarea').forEach(n=>n.disabled=b);}
async function load(){
  $('applications').replaceChildren();$('reviewHistory').replaceChildren();roster=[];
  const r=await api('parent_binding_review_list');
  if(!Array.isArray(r.applications)||!Array.isArray(r.roster))throw Error('請先部署家長中心 v3 Gateway。');
  roster=r.roster;
  const pending=r.applications.filter(a=>a.status==='待審核'), history=r.applications.filter(a=>a.status!=='待審核');
  $('applications').replaceChildren();$('reviewHistory').replaceChildren();
  $('pendingCount').textContent=`待審核 ${pending.length} 份`;
  if(!pending.length)$('applications').append(el('div','目前沒有待審申請。家長送出後，按「重新讀取」即可查看。','panel muted'));
  pending.forEach(a=>$('applications').append(applicationCard(a,true)));
  history.forEach(a=>$('reviewHistory').append(applicationCard(a,false)));
  $('authPanel').hidden=true;$('reviewPanel').hidden=false;
  needsRefresh=false;
}
function applicationCard(a,pending){
  const card=el('article','','panel review-card'),head=el('div','','staff-toolbar');
  head.append(el('h3',a.parentName+' · '+a.relationship),el('span',a.status,'badge'));
  card.append(head,el('p',`聯絡電話：${a.phone}`),el('p',`申請時間：${a.createdAt}`,'help'),el('p',a.applicationId,'application-id'));
  if(!pending){
    card.append(el('p','申請孩子：'+a.children.join('、')));
    if(a.mappings.length)card.append(el('p','核准對應：'+a.mappings.map(m=>`${m.name}（${m.studentNo}）`).join('、')));
    card.append(el('p',`處理人：${a.reviewer||'—'} · ${a.reviewedAt||''}`,'help'));
    if(a.note)card.append(el('p','說明：'+a.note));
    if(a.status==='已核准'){
      const revoke=el('button','撤銷這份申請新增的授權','secondary');revoke.type='button';
      revoke.onclick=async()=>{
        const note=window.prompt('請填寫撤銷原因。這會撤銷本申請全部孩子的新增授權，不會移除主檔或其他申請的既有綁定。');
        if(!note?.trim())return;
        if(window.confirm(`確定撤銷 ${a.parentName} 這份申請的全部 ${a.children.length} 位孩子授權？`))await review(a,{decision:'revoke',note:note.trim()});
      };card.append(revoke);
    }
    return card;
  }
  const form=el('form'), selectors=[];
  a.children.forEach((name,i)=>{
    const label=el('label',`孩子 ${i+1}：申請填寫「${name}」`),select=el('select');
    select.required=true;select.setAttribute('aria-label',`對應孩子 ${i+1}`);
    const placeholder=el('option','請核對並選擇本學期學生');placeholder.value='';select.append(placeholder);
    roster.forEach((k,index)=>{const option=el('option',`${k.name} · ${k.studentNo} · ${k.cls||'未填班級'}`);option.value=String(index);select.append(option);});
    label.append(select);form.append(label);selectors.push(select);
  });
  const verifyLabel=el('label','','makeup'),verify=el('input');verify.type='checkbox';verify.required=true;
  verifyLabel.append(verify,el('span','已透過校方既有聯絡管道核驗身分，並逐位確認親子／照顧關係。'));
  const noteLabel=el('label','審核備註（退回必填，家長可見）'),note=el('textarea');note.maxLength=300;note.rows=2;noteLabel.append(note);
  const actions=el('div','','actions'),approve=el('button',`一次核准 ${a.children.length} 位孩子`,'primary'),reject=el('button','退回補正','secondary');
  approve.type='submit';reject.type='button';actions.append(approve,reject);form.append(verifyLabel,noteLabel,actions);
  form.onsubmit=async e=>{
    e.preventDefault();if(working)return;
    if(needsRefresh){$('result').textContent='上次操作結果尚待確認，請先按「重新讀取」。';return;}
    const selected=selectors.map(s=>roster[Number(s.value)]);
    if(selectors.some(s=>s.value==='')||new Set(selected.map(k=>k.studentNo)).size!==selected.length){$('result').textContent='請完整對應不同的孩子，不可重複選擇。';return;}
    if(!window.confirm(`家長：${a.parentName}（${a.relationship}）\n${selected.map((k,i)=>`${a.children[i]} → ${k.name}（${k.studentNo}）`).join('\n')}\n\n確定全部核准？家長重新讀取後即可為這些孩子請假。`))return;
    await review(a,{decision:'approve',identityVerified:verify.checked,mappings:selected.map(k=>({studentNo:k.studentNo,sheet:k.sheet})),note:note.value.trim()});
  };
  reject.onclick=async()=>{
    if(!note.value.trim()){$('result').textContent='請填寫退回原因。';note.focus();return;}
    if(window.confirm(`退回 ${a.parentName} 的整份申請？\n原因：${note.value.trim()}`))await review(a,{decision:'reject',note:note.value.trim()});
  };
  card.append(form);return card;
}
async function review(a,data){
  if(working)return;
  if(needsRefresh){$('result').textContent='上次操作結果尚待確認，請先按「重新讀取」。';return;}
  setBusy(true);$('result').textContent='正在處理整份申請…';
  let confirmed=false;
  try{
    const r=await api('parent_binding_review',{applicationId:a.applicationId,...data});confirmed=true;
    $('result').textContent=`${a.parentName}：${r.application.status}。未發送 LINE 訊息；家長可重新讀取狀態。`;
    await load();
  }catch(e){needsRefresh=true;$('result').textContent=(confirmed?'處理已成功，但列表更新失敗。':'')+e.message+' 請重新讀取確認狀態。';}
  finally{setBusy(false);}
}
$('authForm').onsubmit=async e=>{
  e.preventDefault();if(working)return;credentials=Object.fromEntries(new FormData(e.target));e.target.elements.adminPass.value='';
  setBusy(true);$('result').textContent='正在核驗校方帳號並讀取…';
  try{await load();$('result').textContent='已讀取最新申請。';}catch(e){credentials=null;roster=[];$('result').textContent=e.message;}
  finally{setBusy(false);}
};
$('refreshReviews').onclick=async()=>{if(working)return;setBusy(true);try{await load();$('result').textContent='已重新讀取最新狀態。';}catch(e){$('result').textContent=e.message;}finally{setBusy(false);}};
$('signOut').onclick=()=>{credentials=null;roster=[];$('applications').replaceChildren();$('reviewHistory').replaceChildren();$('codeResult').textContent='';$('authForm').reset();$('authPanel').hidden=false;$('reviewPanel').hidden=true;$('result').textContent='已清除本頁登入資料。';};
$('issueForm').onsubmit=async e=>{
  e.preventDefault();if(working)return;
  const data=Object.fromEntries(new FormData(e.target));
  if(!window.confirm(`確定為 ${data.studentNo} 家長欄位 ${data.slot} 核發一次性綁定碼？`))return;
  setBusy(true);$('codeResult').textContent='正在核發…';
  try{const r=await api('parent_issue_code',data);$('codeResult').textContent=`${r.name}（${r.studentNo}）／${r.code}；30 分鐘有效，請勿張貼群組。`;}
  catch(e){$('codeResult').textContent=e.message;}finally{setBusy(false);}
};
window.addEventListener('pagehide',()=>{alive=false;credentials=null;roster=[];requests.forEach(c=>c.abort());requests.clear();});
if(embedded){
  document.documentElement.classList.add('embedded');
  $('authPanel').hidden=true;$('reviewPanel').hidden=false;
  $('result').textContent='沿用後台登入，正在讀取待審申請…';
  const observer=new ResizeObserver(()=>{
    try { window.parent.XGParentBinding?.resize(window,document.querySelector('main').getBoundingClientRect().height); } catch (_) {}
  });
  observer.observe(document.querySelector('main'));
  setBusy(true);
  load().then(()=>{$('result').textContent='已讀取最新申請。';})
    .catch(e=>{$('result').textContent=e.message;})
    .finally(()=>setBusy(false));
}
