/* Fixed schedule UI. Backend is authoritative; no writes during load. */
(function(){
  'use strict';
  let ready=false;
  const escape=s=>String(s||'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  window.XGRoutePolicy={
    accept(result){
      ready=!!result.fixedPolicy&&result.fixedPolicy.version===_rpSpec().version&&result.fixedPolicy.effective===_rpSpec().effective;
      if(!ready)throw Error('固定接送規則後端尚未更新，已停止編輯；請先部署 Gateway');
    },
    isFixed:date=>_rpApplies(date),
    templates(date){
      if(!ready)throw Error('固定規則尚未載入');
      const rows=_rpTemplates(date);
      if(!rows.length)throw Error('此日不排車：'+_rpDay(date).reason);
      return rows.map(fields=>({fields}));
    },
    canCopy(start){
      const d=_rpDate(start);d.setUTCDate(d.getUTCDate()+4);
      return d.toISOString().slice(0,10)<_rpSpec().effective;
    },
    checkDetails(row,session,count){
      const rule=_rpRule(row.date,row.route);if(!rule.enforced)return;
      if(!ready)throw Error('固定規則尚未載入');
      if(rule.closed||row.status==='休'||!rule[session])throw Error(rule.reason||'本日沒有此班次');
      const errors=_rpRowErrors(row,true);if(errors.length)throw Error('固定欄位不符，請先核對：'+errors.join('；'));
      if(rule.capacity!=null&&count>rule.capacity)throw Error('本趟上限 '+rule.capacity+' 人，超載不得儲存');
    },
    decorate(el,rows){
      rows.forEach(row=>{
        const tr=el.querySelector('tr[data-row="'+Number(row.row_index)+'"]');
        if(!tr||!_rpApplies(row.date))return;
        const rule=_rpRule(row.date,row.route),errors=row.policyErrors||_rpRowErrors(row,false);
        tr.querySelectorAll('[data-field]').forEach(control=>{
          const field=control.dataset.field,session=field.startsWith('noon_')?'noon':field.startsWith('pm_')?'pm':'';
          const fixed=['noon_vehicle','pm_vehicle','noon_time','pm_time','noon_count','pm_count'].includes(field);
          if(fixed||rule.closed||(session&&!rule[session])||errors.length||row.noon_returned_at||row.pm_returned_at)control.disabled=true;
          if(fixed)control.title='固定規則或系統依學生明細計算，不可直接修改';
        });
        tr.querySelectorAll('.btn-detail').forEach(b=>{
          if(rule.closed||!rule[b.dataset.session]||errors.length||row.status==='休'||row.noon_returned_at||row.pm_returned_at)b.disabled=true;
        });
        tr.querySelectorAll('.btn-route-del,.btn-day-del').forEach(b=>{b.disabled=true;b.title='固定班表保留紀錄，請改設休';});
        // A conflicting uncompleted row can only be made inactive, never scheduled.
        const status=tr.querySelector('[data-field="status"]');
        if(status&&(rule.closed||errors.length)&&!row.noon_returned_at&&!row.pm_returned_at){
          status.disabled=false;[...status.options].forEach(o=>{if(o.value==='上')o.disabled=true;});
        }
        const routeCell=tr.children[3];
        if(routeCell&&!routeCell.querySelector('.rp-tag')){
          const tag=document.createElement('div');tag.className='rp-tag';tag.style.cssText='font-size:12px;color:'+(errors.length?'#b91c1c':'#176347');
          tag.textContent=errors.length?'規則衝突：'+errors.join('；'):'固定班表'+(rule.extra?' · 14:45 獨立趟次':'');
          routeCell.appendChild(tag);
        }
        if(tr.children[4]&&rule.capacity===null&&!rule.closed)tr.children[4].textContent='無上限';
      });
    },
    bind(el,api,loadAll,isCurrent){
      if(!document.getElementById('pickupPolicyMenu')){
        const toggle=document.createElement('button');toggle.id='pickupPolicyMenu';toggle.textContent='展開後台選單';
        toggle.setAttribute('aria-expanded','false');
        toggle.onclick=()=>{const open=document.body.classList.toggle('pickup-policy-menu-open');toggle.setAttribute('aria-expanded',String(open));toggle.textContent=open?'收起後台選單':'展開後台選單';};
        document.querySelector('.sidebar').prepend(toggle);
      }
      const box=document.createElement('div');box.id='ps-policy-info';
      box.style.cssText='padding:12px 16px;margin:0 0 14px;background:#eef7f1;border:1px solid #c6dfce;border-radius:6px;font-size:14px;line-height:1.7';
      box.innerHTML='<strong>固定接送規則｜2026/09/29 起</strong><br>A 6／B 4／C 4／小巴 8／半巴 4 人；走路無上限。週二僅 15:25，週三僅 12:25；週五另有六條 14:45 加班。國定假日、補假、週末不排車。<br>半巴為新增路線，不合併 A車加。歷史資料不改；固定容量、車型、時間與系統人數不可手改。';
      el.querySelector('#ps-status').before(box);
      const button=document.createElement('button');button.id='ps-fixed-create';button.textContent='建立固定班表（先預覽）';
      button.style.cssText='padding:8px 12px;background:#176347;color:white;border:0;border-radius:5px;cursor:pointer;margin-top:8px';
      box.appendChild(document.createElement('br'));box.appendChild(button);
      let busy=false;
      button.onclick=async()=>{
        if(busy||!ready||!isCurrent())return;
        const start=prompt('固定班表起日（不得早於 2026-09-29）',_rpSpec().effective);if(!start)return;
        const end=prompt('固定班表迄日（至本學期 2027-01-20）',_rpSpec().through);if(!end)return;
        busy=true;button.disabled=true;button.textContent='正在預覽，不會寫入…';
        let wrote=false;
        try{
          const plan=await api({action:'policy_plan',start,end});if(!isCurrent())return;
          if(!plan.success)throw Error(plan.error||'預覽失敗');
          if(plan.conflicts.length)throw Error('既有排程有衝突，未寫入：\n'+plan.conflicts.slice(0,12).join('\n'));
          if(!plan.inserted){alert('固定班表已存在，未新增任何資料。');return;}
          const summary='日期：'+start+'～'+end+'\n新增 '+plan.inserted+' 列固定趟次框架；保留 '+plan.existing+' 列既有資料。\n略過 '+plan.skippedDays.length+' 個週末／假日。\n\n不會自動加入學生、不複製完成狀態，不更改既有名單。\n確認建立？';
          if(!confirm(summary))return;
          wrote=true;button.textContent='建立中，請勿重複送出…';
          const r=await api({action:'policy_apply',start,end,fingerprint:plan.fingerprint});
          if(!r.success)throw Error(r.error||'建立結果不明');
          if(!isCurrent())return;
          await loadAll();alert('已建立 '+r.inserted+' 列固定框架。請逐班次選學生及安排接送人員。');
        }catch(e){if(isCurrent())alert(e.message+(wrote?'\n請先重新載入核對，不要直接重複建立。':''));}
        finally{busy=false;if(isCurrent()){button.disabled=false;button.textContent='建立固定班表（先預覽）';}}
      };
    }
  };
})();
