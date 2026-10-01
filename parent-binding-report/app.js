(function(){
 'use strict';
 const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
 let students=[],apps=[];
 function freeze(o){Object.values(o).forEach(v=>{if(v&&typeof v==='object'&&!Object.isFrozen(v))freeze(v);});return Object.freeze(o);}

 const nameKey=v=>String(v??'').normalize('NFKC').trim();
 function phoneKey(value){
   let v=String(value??'').normalize('NFKC').replace(/[\s()（）-]/g,'');
   if(/^\+?8869\d{8}$/.test(v))v='0'+v.replace(/^\+?886/,'');
   return /^09\d{8}$/.test(v)?v:null;
 }
 function compareName(a,b){if(!nameKey(a)||!nameKey(b))return 'missing';return nameKey(a)===nameKey(b)?'same':'different';}
 function comparePhone(a,b){const x=phoneKey(a),y=phoneKey(b);return !x||!y?'missing':x===y?'same':'different';}
 const relationSlot=r=>r==='爸爸'?'father':r==='媽媽'?'mother':null;
 function comparison(student,application){
   const slot=relationSlot(application?.relation),base=slot?student[slot]:null;
   return {name:base?compareName(base.name,application.parent):'missing',phone:base?comparePhone(base.phone,application.phone):'missing',base,
     parent:application?.parent||'',phoneValue:application?.phone||''};
 }
 function pendingFor(student){return apps.filter(a=>a.status==='待審核'&&a.claims.some(n=>nameKey(n)===nameKey(student.name)));}
 function candidates(application){
   return application.claims.map(name=>({name,students:students.filter(s=>nameKey(s.name)===nameKey(name))}));
 }
 function build(){
   return students.map(student=>{
     const bindings=new Map();
     for(const [slot,relation] of [['father','爸爸'],['mother','媽媽']]){
       const g=student[slot];
       if(g.key)bindings.set(g.key,{key:g.key,relation,parent:g.name||'姓名未填',sources:['主檔既有綁定'],application:null});
     }
     for(const a of apps){
       if(a.status!=='已核准'||!a.key||!a.mappings.includes(student.id))continue;
       const old=bindings.get(a.key);
       if(old){old.sources.push('已核准申請 '+a.id);old.application=a;old.parent=a.parent.trim();old.relation=a.relation;}
       else bindings.set(a.key,{key:a.key,relation:a.relation,parent:a.parent.trim(),sources:['已核准申請 '+a.id],application:a});
     }
     const parents=[...bindings.values()].map(b=>({...b,comparison:comparison(student,b.application)}));
     const pending=pendingFor(student),bound=parents.length>0;
     const needsReview=pending.length>0||student.father.invalidBinding||student.mother.invalidBinding||parents.some(p=>p.comparison.name!=='same'||p.comparison.phone!=='same');
     return {...student,parents,pending,bound,needsReview};
   });
 }
 let report=[],filter='all',query='',lastFocus=null;
 function badge(state,label){
   const c={same:'good',different:'warn',missing:'neutral'}[state],text={same:'一致',different:'不同',missing:'資料不足'}[state];
   return '<span class="pill '+c+'">'+esc(label+' '+text)+'</span>';
 }
 function guardianCell(g){return '<strong>'+esc(g.name||'姓名未填')+'</strong><small>'+esc(g.phone||'手機未填')+'</small>';}
 function visible(){
   const q=nameKey(query).toLowerCase(),digits=q.replace(/[\s()-]/g,'');
   return report.filter(s=>{
     if(filter==='bound'&&!s.bound||filter==='unbound'&&s.bound||filter==='review'&&!s.needsReview)return false;
     const blob=[s.id,s.name,s.english,s.cls,s.father.name,s.father.phone,s.mother.name,s.mother.phone,
       ...s.parents.flatMap(p=>[p.parent,p.application?.phone||'']),...s.pending.flatMap(p=>[p.parent,p.phone])].join(' ').toLowerCase();
     return !q||blob.includes(q)||(digits&&blob.replace(/[\s()-]/g,'').includes(digits));
   });
 }
 function render(){
   if(!hasData)return;
   document.getElementById('metric-total').textContent=report.length;
   document.getElementById('metric-bound').textContent=report.filter(s=>s.bound).length;
   document.getElementById('metric-unbound').textContent=report.filter(s=>!s.bound).length;
   document.getElementById('metric-pending').textContent=apps.filter(a=>a.status==='待審核').length;
   const rows=visible();
   document.getElementById('table-count').textContent='顯示 '+rows.length+' / '+report.length+' 位學生';
   document.getElementById('empty').hidden=rows.length>0;
   document.getElementById('empty').textContent='沒有符合條件的學生。';
   document.getElementById('student-rows').innerHTML=rows.map(s=>'<tr data-student="'+esc(s.id)+'">'+
    '<td><strong>'+esc(s.name)+'</strong><small>'+esc(s.english+' · '+s.cls)+'</small><small>'+esc(s.id)+'</small>'+
    (s.pending.length?'<small class="warn pill">'+(s.pending.some(a=>candidates(a).some(c=>c.students.length>1))?'同名申請未歸戶':'有待審候選對應')+'</small>':'')+'</td>'+
    '<td>'+guardianCell(s.father)+'</td><td>'+guardianCell(s.mother)+'</td>'+
    '<td>'+(s.parents.length?s.parents.map(p=>'<div class="person">'+esc(p.relation+'：'+p.parent)+'</div>').join(''):'<span class="muted">無有效綁定家長</span>')+'</td>'+
    '<td>'+(s.parents.length?s.parents.map(p=>'<div class="check-person"><small>'+esc(p.relation)+'</small><div class="checks">'+badge(p.comparison.name,'姓名')+badge(p.comparison.phone,'電話')+'</div></div>').join(''):'<span class="pill neutral">尚無有效資料可比對</span>')+'</td>'+
    '<td><span class="pill '+(s.bound?'good':'neutral')+'">'+(s.bound?'已綁定':'未綁定')+'</span><small>'+s.parents.length+' 位有效家長</small></td>'+
    '<td><button type="button" class="detail-link" data-detail="'+esc(s.id)+'" aria-label="查看 '+esc(s.name)+' '+esc(s.id)+' 明細">查看明細</button></td></tr>').join('');
   document.querySelectorAll('[data-detail]').forEach(b=>b.addEventListener('click',()=>showDetails(b.dataset.detail,b)));
   document.querySelectorAll('[data-filter]').forEach(b=>b.setAttribute('aria-pressed',String(b.dataset.filter===filter)));
 }
 function matchField(label,state,schoolValue,submitted){
   return '<div class="match-field '+({same:'good',different:'warn',missing:'neutral'}[state])+'"><strong>'+label+'：'+({same:'一致',different:'不同',missing:'資料不足'}[state])+'</strong>'+
    '<div>校方留存：'+esc(schoolValue||'未提供')+'</div><div>申請填寫：'+esc(submitted||'無申請資料')+'</div></div>';
 }
 function parentDetail(student,p){
   const c=p.comparison,a=p.application;
   return '<article class="parent-detail"><div class="parent-detail-head"><h3>'+esc(p.relation+' · '+p.parent)+'</h3><span class="pill good">有效綁定</span></div>'+
     '<div class="match-grid">'+matchField('姓名',c.name,c.base?.name,c.parent)+matchField('電話',c.phone,c.base?.phone,c.phoneValue)+'</div>'+
     '<p class="source-line">來源：'+esc(p.sources.join(' ＋ '))+' · '+(a?'審核日期：'+esc(a.reviewed):'無申請資料可交叉比對；不自動標綠')+'</p></article>';
 }
 function showDetails(id,button){
   const s=report.find(s=>s.id===id);if(!s)return;lastFocus=button;
   document.getElementById('detail-title').textContent=s.name+'（'+s.english+'） · '+s.id;
   const history=apps.filter(a=>a.mappings.includes(s.id)||a.claims.some(n=>nameKey(n)===nameKey(s.name)));
   document.getElementById('detail-body').innerHTML='<p class="explanation">這是唯讀分析。姓名／電話一致僅代表資料相符，不是核准依據；本頁不會改變學生或綁定關係。</p>'+
    '<section><h3>學生與校方留存資料</h3><p class="muted">115-1 · '+esc(s.cls)+' · 在學</p><div class="badge-row"><span class="pill '+(s.bound?'good':'neutral')+'">'+(s.bound?'已綁定':'未綁定')+'</span><span>'+s.parents.length+' 位有效家長；學生人數只算 1 人</span></div>'+
    '<div class="master-grid"><div class="master-box"><strong>爸爸</strong>'+guardianCell(s.father)+'</div><div class="master-box"><strong>媽媽</strong>'+guardianCell(s.mother)+'</div></div></section>'+
    '<section><h3>誰已綁定</h3>'+(s.parents.length?s.parents.map(p=>parentDetail(s,p)).join(''):'<p class="muted">沒有有效綁定。待審或僅同名的申請不算已綁定。</p>')+'</section>'+
    '<section><h3>相關申請紀錄（包含待審及歷史）</h3>'+history.map(a=>{
      const pending=a.status==='待審核',ambiguous=candidates(a).some(c=>c.students.length!==1),mapped=a.mappings.includes(s.id),c=comparison(s,a);
      return '<article class="history"><div><strong>'+esc(a.parent+' · '+a.relation)+'</strong> <span class="pill '+(pending?'warn':a.status==='已核准'?'good':'neutral')+'">'+a.status+'</span></div>'+
       '<p>電話：'+esc(a.phone||'未填')+' · '+esc(a.id)+'</p><p>申請孩子：'+esc(a.claims.join('、'))+'</p>'+
       '<p class="subtle">'+(mapped?'正式核准對應曾包含此學生編號':ambiguous?'同名／多重候選，尚未歸戶，不表示屬於此學生':'僅按孩子姓名列為候選，尚未正式對應')+'</p>'+
       (!ambiguous||mapped?'<div class="checks">'+badge(c.name,'姓名')+badge(c.phone,'電話')+'</div>':'<span class="pill neutral">尚無確定對應，不標示比對正確</span>')+
       '<p class="subtle">申請 '+esc(a.at)+' · '+(a.reviewed?'審核 '+esc(a.reviewed):'尚未審核')+'</p></article>';
    }).join('')+'</section>';
   document.getElementById('detail-dialog').showModal();document.getElementById('close-dialog').focus();
 }
 function renderPending(){
   const pending=apps.filter(a=>a.status==='待審核');
   document.getElementById('pending-count').textContent=pending.length+' 份申請';
   document.getElementById('pending-list').innerHTML=pending.map(a=>{
     const c=candidates(a),unique=c.length===1&&c[0].students.length===1,target=unique?c[0].students[0]:null,match=target?comparison(target,a):null;
     return '<article class="pending-card" data-application="'+esc(a.id)+'"><strong>'+esc(a.parent+' · '+a.relation)+'</strong> <span class="pill warn">待審核</span>'+
     '<p>申請電話：'+esc(a.phone||'未填')+'</p><p>申請孩子：'+esc(a.claims.join('、'))+'</p>'+
     '<p>'+(unique?'候選：'+esc(target.id+' · '+target.cls)+'（未正式綁定）':'同名孩子 '+c.reduce((n,x)=>n+x.students.length,0)+' 位，尚未歸戶')+'</p>'+
     (match?'<div class="checks">'+badge(match.name,'姓名')+badge(match.phone,'電話')+'</div>':'<span class="pill neutral">無確定對應，不標綠</span>')+
     '<p class="subtle">'+esc(a.at)+' · '+esc(a.id)+'</p></article>';
   }).join('');
 }
 document.querySelectorAll('[data-filter]').forEach(b=>b.addEventListener('click',()=>{filter=b.dataset.filter;render();}));
 document.getElementById('search').addEventListener('input',e=>{query=e.target.value;render();});
 document.getElementById('close-dialog').addEventListener('click',()=>document.getElementById('detail-dialog').close());
 document.getElementById('detail-dialog').addEventListener('close',()=>lastFocus?.focus());
 // API is intentionally native POST in its own frame, not the legacy GET wrapper.
 const API='https://script.google.com/macros/s/AKfycbw-7_a_OfUVlgegcLxkux_9dr9UlYSVKhi3uQjV-0sr2X2TpRRmCXtSM7jbIqMHK4hNww/exec';
 let busy=false,alive=true,hasData=false,controller=null;
 function currentSession(){
   try{
     if(window.parent===window)return null;
     return window.parent.XGParentBindingReport?.session(window)||null;
   }catch(_){return null;}
 }
 function sameSession(initial){
   const now=currentSession();
   return alive&&now&&initial&&now.adminUser===initial.adminUser&&now.epoch===initial.epoch;
 }
 function erase(message){
   hasData=false;students=[];apps=[];report=[];
   for(const id of ['student-rows','pending-list','detail-body'])document.getElementById(id).replaceChildren();
   for(const id of ['metric-total','metric-bound','metric-unbound','metric-pending'])document.getElementById(id).textContent='—';
   document.getElementById('pending-count').textContent='尚未取得';
   document.getElementById('table-count').textContent='尚未完成資料核對';
   document.getElementById('empty').hidden=false;
   document.getElementById('empty').textContent=message;
   document.getElementById('read-status').textContent=message;
   document.getElementById('detail-dialog').close();
 }
 function validateData(data){
   if(data.version!=='parent-binding-report-v1'||data.readOnly!==true||data.semester!=='115-1'||
     !Array.isArray(data.students)||!Array.isArray(data.applications)||!data.totals)throw Error('對照資料或 Gateway 版本不符，不能視為零人');
   const ids=new Set(),applicationIds=new Set();
   data.students.forEach(s=>{
     if(!s||typeof s.id!=='string'||!s.id||ids.has(s.id)||typeof s.name!=='string'||!s.father||!s.mother)throw Error('學生欄位缺漏或編號重複');
     ids.add(s.id);
     ['father','mother'].forEach(slot=>{
       const p=s[slot];if(typeof p.name!=='string'||typeof p.phone!=='string'||typeof p.key!=='string'||('uid' in p))throw Error('家長欄位格式不符');
     });
   });
   data.applications.forEach(a=>{
     if(!a||!a.id||applicationIds.has(a.id)||!Array.isArray(a.claims)||!Array.isArray(a.mappings)||
       a.claims.some(n=>typeof n!=='string')||a.mappings.some(id=>typeof id!=='string')||typeof a.parent!=='string'||
       typeof a.phone!=='string'||typeof a.relation!=='string'||typeof a.key!=='string'||
       !['待審核','已核准','已退回','已撤銷'].includes(a.status))throw Error('申請資料欄位不符，停止統計');
     applicationIds.add(a.id);
   });
 }
 async function loadReport(){
   if(busy)return;
   const session=currentSession();
   if(!session){erase('請從正式後台以管理員或行政帳號開啟此頁，不提供獨立登入。');return;}
   busy=true;document.getElementById('reload').disabled=true;
   erase('正在取得 115-1 在學名冊、家長資料及綁定申請；尚未確認，不代表零人。');
   controller=new AbortController();let deadline;
   const timeout=new Promise((_,reject)=>{deadline=setTimeout(()=>{controller?.abort();reject(Error('讀取超過 35 秒，尚未取得完整資料；請手動重讀，不自動重試。'));},35000);});
   try{
     const response=await Promise.race([(async()=>{
       const r=await fetch(API,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},
         body:JSON.stringify({_action:'parent_binding_report',semester:'115-1',adminUser:session.adminUser,adminPass:session.adminPass}),
         cache:'no-store',referrerPolicy:'no-referrer',signal:controller.signal});
       if(!r.ok)throw Error('唯讀服務連線失敗，未完成資料核對');
       return r.json();
     })(),timeout]);
     if(!sameSession(session))throw Error('登入狀態已變更，已清除舊回應');
     if(response.success!==true)throw Error(response.error||'唯讀資料未成功取得');
     validateData(response);students=response.students;apps=response.applications;
     freeze(students);freeze(apps);report=build();
     const counts={students:report.length,bound:report.filter(s=>s.bound).length,unbound:report.filter(s=>!s.bound).length,
       pending:apps.filter(a=>a.status==='待審核').length};
     if(Object.keys(counts).some(k=>counts[k]!==response.totals[k]))throw Error('前後端統計不一致，已停止顯示');
     if(!sameSession(session))throw Error('登入已失效');
     hasData=true;render();renderPending();
     document.getElementById('read-status').textContent='資料取得：'+String(response.generatedAt||'未提供')+
       ' · 待審含未歸戶申請；統計不隨搜尋改變。新申請或審核變更後，按「重新讀取」。';
   }catch(e){
     erase(!sameSession(session)?'登入已失效，資料已清除。':e.message||'尚未取得資料，請重新讀取。');
   }finally{
     clearTimeout(deadline);controller=null;busy=false;
     document.getElementById('reload').disabled=!currentSession();
   }
 }
 window.clearBindingReport=function(){alive=false;controller?.abort();erase('已離開唯讀對照頁，資料已清除。');};
 window.addEventListener('pagehide',window.clearBindingReport);
 document.getElementById('reload').addEventListener('click',loadReport);
 loadReport();

})();
