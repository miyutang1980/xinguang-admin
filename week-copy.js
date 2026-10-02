/* Fixed week-copy preview only. No API calls or permission changes here. */
(function(){
  'use strict';
  const node=(tag,text)=>{const n=document.createElement(tag);if(text!=null)n.textContent=String(text);return n;};
  window.XGWeekCopy={confirm(plan,source,target,isCurrent,host){
    return new Promise(resolve=>{
      const dialog=node('dialog');dialog.id='week-copy-preview';
      dialog.setAttribute('aria-labelledby','week-copy-heading');
      dialog.style.cssText='width:min(940px,calc(100vw - 32px));max-width:calc(100vw - 32px);max-height:85dvh;padding:20px;border:1px solid #c6dfce;border-radius:8px;color:#243d35;background:white;box-sizing:border-box;overflow:auto';
      const heading=node('h2','整週接送複製預覽');heading.id='week-copy-heading';heading.style.marginTop='0';dialog.append(heading);
      dialog.append(node('p',source+' → '+target));
      dialog.append(node('p','新增 '+plan.added+' 列；填入空白 '+plan.filled+' 列；學生安排 '+plan.students+' 筆（不是不重複學生人數）。'));
      const note=node('p','固定車型、時間與容量以目標週為準。既有安排整列保留；不新增暫緩的 15:25 半巴，不複製完成狀態，不發通知。2026/10/9 複製後直接設為「休」，只保留名單、不接送；其他假日仍跳過。');
      note.style.cssText='background:#eef7f1;padding:12px;line-height:1.6';dialog.append(note);
      if(plan.conflicts.length){
        const warning=node('p','有 '+plan.conflicts.length+' 項衝突，無法執行。');warning.style.color='#a52b22';dialog.append(warning);
        const list=node('ul');plan.conflicts.forEach(t=>list.append(node('li',t)));dialog.append(list);
      }
      const wrap=node('div');wrap.style.overflowX='auto';
      const table=node('table');table.style.cssText='border-collapse:collapse;width:100%;min-width:640px;font-size:14px';
      const head=node('tr');['日期／路線','方式／狀態','時段','學校順序','預排／期間不接／其餘'].forEach(t=>head.append(node('th',t)));table.append(head);
      plan.summary.forEach(row=>row.trips.forEach(t=>{
        const tr=node('tr');
        [row.date+' '+row.route,(row.isNew?'新增':'填入空白')+(row.status==='休'?' · 休（不接送）':''),t.time,t.schools.join(' → ')||'未安排',t.scheduled+' / '+t.paused+' / '+t.required]
          .forEach(value=>tr.append(node('td',value)));
        table.append(tr);
      }));
      table.querySelectorAll('td,th').forEach(cell=>cell.style.cssText='padding:8px;text-align:left;border-bottom:1px solid #dce7e0;vertical-align:top');
      wrap.append(table);dialog.append(wrap);
      function section(title,items){if(!items.length)return;dialog.append(node('h3',title));const ul=node('ul');items.forEach(x=>ul.append(node('li',x)));dialog.append(ul);}
      section('提醒（不阻擋）',plan.warnings.map(r=>r.date+' '+r.route+'：'+r.reason));
      section('保留／跳過（'+plan.skipped.length+'）',plan.skipped.map(r=>r.date+' '+r.route+'：'+r.reason));
      section('依目標日期核對期間不接送',plan.paused.map(r=>r.date+' '+r.route+' '+(r.session==='noon'?'中午':'下午')+' '+r.studentId+'：'+r.reason));
      dialog.append(node('p',plan.pauseNote||''));
      dialog.append(node('p','期間不接學生仍保留在名單。「其餘」不代表已扣除目標日所有請假；當日請假及通知依原流程另行判定。'));
      const actions=node('div');actions.style.cssText='display:flex;flex-wrap:wrap;gap:12px;margin-top:18px';
      const cancel=node('button','取消，不複製'),apply=node('button','確認複製');
      cancel.type=apply.type='button';cancel.id='week-copy-cancel';apply.id='week-copy-apply';
      cancel.style.cssText=apply.style.cssText='padding:10px 16px;border:1px solid #176347;border-radius:4px;cursor:pointer';
      apply.style.background='#176347';apply.style.color='white';
      apply.disabled=plan.conflicts.length>0||!(plan.added+plan.filled);
      if(apply.disabled){apply.style.opacity='.45';apply.style.cursor='not-allowed';}
      actions.append(cancel,apply);dialog.append(actions);
      let done=false;
      const finish=value=>{if(done)return;done=true;observer.disconnect();dialog.close();dialog.remove();resolve(value);};
      const observer=new MutationObserver(()=>{if(!host.isConnected||!isCurrent())finish(false);});
      observer.observe(document.body,{childList:true,subtree:true});
      cancel.onclick=()=>finish(false);
      apply.onclick=()=>finish(!apply.disabled&&isCurrent());
      dialog.addEventListener('cancel',e=>{e.preventDefault();finish(false);});
      document.body.append(dialog);dialog.showModal();cancel.focus();
    });
  }};
})();
