(function(){
 'use strict';
 const VERSION='pickup-exceptions-v1',esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
 let ready=false,rules=[],state=null,owner=null,lastRows=[];
 const auth=()=>({adminUser:_currentUser?.username||'',adminPass:sessionStorage.getItem('xg_session_pwd')||''});
 async function call(action,params={}){
   const epoch=_authEpoch,user=_currentUser;
   const r=await gwCallJson(action,{...auth(),...params},true);
   if(epoch!==_authEpoch||user!==_currentUser)throw Error('登入已改變，請重新載入');
   if(!r?.success)throw Error(r?.error||'未取得正確回應');
   return r;
 }
 function currentRules(){return owner===_currentUser?rules:[];}
 function makeButton(){
   const info=document.getElementById('ps-policy-info');
   if(!info)return;
   const existing=document.getElementById('rx-open');
   if(existing){existing.disabled=!ready;existing.title=ready?'保留學生排程，依期間／星期／班次標示不接':'需先部署期間不接送 Gateway';return;}
   const b=document.createElement('button');b.id='rx-open';b.type='button';b.textContent='期間不接送規則';
   b.className='rx-button';b.onclick=open;b.disabled=!ready;
   b.title=ready?'保留學生排程，依期間／星期／班次標示不接':'需先部署期間不接送 Gateway';
   info.append(document.createElement('br'),b);
   const note=document.createElement('span');note.className='rx-legend';note.textContent='線別固定配色；同線續頁同色。暫停接送仍保留原排程。';info.append(note);
 }
 function mark(rows){
   lastRows=rows;
   makeButton();
   const list=currentRules();
   rows.forEach(row=>{
     const tr=document.querySelector('#ps-table-wrap tr[data-row="'+Number(row.row_index)+'"]');if(!tr)return;
     const color=_rxRouteColor(row.route),cell=tr.children[3];
     if(cell){cell.style.borderLeft='5px solid '+color;cell.style.color=color;}
     ['noon','pm'].forEach(session=>{
       const td=tr.querySelector('[data-'+session+'-cell]');if(!td)return;
       td.querySelector('.rx-trip-note')?.remove();
       const snapshot=typeof _pickupCachedDetails==='function'?_pickupCachedDetails(row,session):null;
       if(!ready||!snapshot)return;
       const students=snapshot.details.flatMap(d=>d.student_ids.map((id,i)=>({id,name:d.student_names[i]})));
       const skipped=students.map(s=>({...s,reason:_rxReason(list,row.date,session,s.id)})).filter(s=>s.reason);
       if(!skipped.length)return;
       const n=document.createElement('div');n.className='rx-trip-note';
       n.textContent='排定 '+students.length+' 人｜依期間規則要接 '+(students.length-skipped.length)+' 人｜暫停 '+skipped.length+' 人\n'+
         skipped.map(s=>s.name+'：不接／'+s.reason).join('\n')+'\n臨時請假另計，當日通知以 Gateway 核對為準';
       td.append(n);
     });
   });
 }
 async function open(){
   if(!ready){alert('請先更新期間不接送 Gateway');return;}
   document.getElementById('rx-modal')?.remove();
   const modal=document.createElement('div');modal.id='rx-modal';modal.className='rx-overlay';
   modal.innerHTML='<section class="rx-panel" role="dialog" aria-modal="true" aria-label="期間不接送規則">'+
     '<header><div><h2>期間不接送規則</h2><p>Recurring no-pickup rules</p></div><button id="rx-close" type="button">關閉</button></header>'+
     '<div class="rx-content"><p>學生仍保留原排程與容量，僅從符合條件當日的「要接／接回」名單扣除。截止日包含當天，隔日恢復一般判定；其他請假或停駛仍有效。</p>'+
     '<p class="rx-warning">建立或停用不立即推播。規則會在當次交通群通知中套用；已發出的卡片、完成紀錄不回寫。當日通知已發出後的急件，請行政另行通知司機。</p>'+
     '<div id="rx-status" role="status">讀取規則及當期名冊…</div><form id="rx-form" hidden>'+
     '<label>學生 Student<select id="rx-student" required></select></label>'+
     '<div class="rx-grid"><label>起日 Start<input id="rx-start" type="date" required></label><label>迄日 End（含當天）<input id="rx-end" type="date" required></label></div>'+
     '<fieldset><legend>每週適用星期</legend>'+['日','一','二','三','四','五','六'].map((v,i)=>'<label class="rx-day"><input type="checkbox" name="weekday" value="'+i+'">週'+v+'</label>').join('')+'</fieldset>'+
     '<label>班次 Session<select id="rx-session"><option value="pm">下午 Afternoon</option><option value="noon">中午 Noon</option><option value="all">全天 All sessions</option></select></label>'+
     '<label>原因 Reason<input id="rx-reason" maxlength="120" required placeholder="例如：社團活動 / Club activity"></label>'+
     '<p>套用該學生所有路線中的指定班次，不因換車漏判。新增重疊規則會被阻擋；要變更請先停用舊規則，再建立新規則。</p>'+
     '<button id="rx-save" class="rx-button" type="submit">確認並建立規則</button></form>'+
     '<h3>既有規則</h3><div id="rx-list"></div><button id="rx-refresh" class="rx-button" type="button">重新取得規則</button></div></section>';
   document.body.append(modal);
   let busy=false,verified=false;const epoch=_authEpoch;
   const el=id=>modal.querySelector('#'+id),valid=()=>modal.isConnected&&epoch===_authEpoch;
   const setStatus=(text,error=false)=>{if(valid()){el('rx-status').textContent=text;el('rx-status').className=error?'rx-error':'';}};
   el('rx-close').onclick=()=>{if(!busy)modal.remove();};
   async function refresh(){
     if(busy)return;busy=true;verified=false;el('rx-save').disabled=true;el('rx-refresh').disabled=true;
     try{
       const [r,students]=await Promise.all([call('routes_exceptions_list'),_loadAllStudentsOnce({})]);
       if(!valid())return;if(r.version!==VERSION||!Array.isArray(r.rules)||!r.revision)throw Error('後端版本或清單格式不符');
       state=r;rules=r.rules;owner=_currentUser;verified=true;mark(lastRows);
       el('rx-student').innerHTML='<option value="">請選學生，不預設任何人</option>'+students.filter(s=>s.status==='active').map(s=>
         '<option value="'+esc(s.student_id)+'">'+esc((s.display_name||s.chinese_name)+'｜'+s.school_name+' '+s.class_code+'｜'+s.student_id)+'</option>').join('');
       el('rx-start').min='2026-08-31';el('rx-end').max=_rpSpec().through;
       if(!el('rx-start').value)el('rx-start').value=r.today;
       el('rx-list').replaceChildren();
       if(!rules.length)el('rx-list').textContent='尚無規則。未替任何學生建立暫停安排。';
       rules.forEach(rule=>{
         const item=document.createElement('article');item.className='rx-rule';
         const status=rule.status==='cancelled'?'已停用':rule.end<r.today?'已到期':rule.start>r.today?'尚未開始':'有效';
         item.innerHTML='<strong>'+esc(rule.name)+'｜'+esc(rule.school)+'</strong><p>'+esc(rule.start+' ～ '+rule.end)+'（含迄日）｜'+
           esc(rule.weekdays.map(d=>'週'+'日一二三四五六'[d]).join('、'))+'｜'+esc({noon:'中午',pm:'下午',all:'全天'}[rule.session])+'</p>'+
           '<p>不接原因：'+esc(rule.reason)+' · '+status+'</p><small>'+esc('最後更新 '+rule.updatedAt+'／'+rule.operator)+'</small>';
         if(rule.status==='active'&&rule.end>=r.today){
           const b=document.createElement('button');b.type='button';b.textContent='停用規則';b.className='rx-button';
           b.onclick=async()=>{
             if(busy||!verified||!confirm('停用 '+rule.name+' 的此筆不接送規則？\n原排程保留；其他請假仍有效，已發通知不修改。'))return;
             busy=true;verified=false;el('rx-save').disabled=true;b.disabled=true;
             try{await call('routes_exceptions_cancel',{id:rule.id,revision:rule.revision});setStatus('已停用，請重新取得規則並重新載入班表');}
             catch(e){setStatus('結果待核對：'+e.message+'。請重新取得規則，不要重送。',true);}
             finally{busy=false;el('rx-refresh').disabled=false;}
           };item.append(b);
         }
         el('rx-list').append(item);
       });
       el('rx-form').hidden=false;setStatus('規則清單已核對。新增不會立即發送 LINE。');
     }catch(e){setStatus(e.message,true);}
     finally{busy=false;if(valid()){el('rx-save').disabled=!verified;el('rx-refresh').disabled=false;}}
   }
   el('rx-refresh').onclick=refresh;
   el('rx-form').onsubmit=async event=>{
     event.preventDefault();if(busy||!verified)return;
     const rule={studentId:el('rx-student').value,start:el('rx-start').value,end:el('rx-end').value,
       weekdays:[...modal.querySelectorAll('[name=weekday]:checked')].map(n=>Number(n.value)),session:el('rx-session').value,reason:el('rx-reason').value.trim()};
     if(!rule.studentId||!rule.weekdays.length||!rule.reason||rule.start>rule.end){setStatus('請完整選學生、星期、日期與原因',true);return;}
     const label=el('rx-student').selectedOptions[0].textContent;
     if(!confirm(label+'\n'+rule.start+' ～ '+rule.end+'（含迄日）\n'+rule.weekdays.map(d=>'週'+'日一二三四五六'[d]).join('、')+'／'+el('rx-session').selectedOptions[0].textContent+
       '\n不接原因：'+rule.reason+'\n保留原排程，不立即推播。是否建立？'))return;
     busy=true;verified=false;el('rx-save').disabled=true;el('rx-refresh').disabled=true;
     try{await call('routes_exceptions_save',{rule,expectedRevision:state.revision,requestId:crypto.randomUUID()});setStatus('規則已建立。請重新取得規則並重新載入班表；不立即推播。');}
     catch(e){setStatus('結果待核對：'+e.message+'。請先重新取得規則，不要重送。',true);}
     finally{busy=false;el('rx-refresh').disabled=false;}
   };
   refresh();
 }
 window.XGPickupExceptions={
   accept(result){owner=_currentUser;ready=result.pickupExceptionsVersion===VERSION;rules=ready&&Array.isArray(result.pickupExceptions)?result.pickupExceptions:[];},
   decorate:mark,
   annotateModal(row,session){
     document.querySelectorAll('#rd-modal-overlay .rx-student-note').forEach(n=>n.remove());
     document.querySelectorAll('#rd-modal-overlay .rd-stu').forEach(cb=>{
       const reason=_rxReason(currentRules(),row.date,session,cb.dataset.sid);if(!reason)return;
       const note=document.createElement('small');note.className='rx-student-note';note.textContent='本日不接：'+reason+'（保留勾選）';
       cb.closest('label').style.flexWrap='wrap';note.style.flexBasis='100%';cb.closest('label').append(note);
     });
   }
 };
})();
