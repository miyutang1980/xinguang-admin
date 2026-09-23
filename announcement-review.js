/* Deep links are navigation only. Server-side authorization must also be enforced
 * by Gateway; these UI checks are not a substitute for server access control. */
(function () {
  'use strict';
  let generation = 0, record = null, intent = 'approve', busy = false, blockedRow = null;
  const params = new URLSearchParams(location.search);
  const validRow = value => /^[1-9]\d*$/.test(String(value)) &&
    Number.isSafeInteger(Number(value)) && Number(value) >= 2;
  const allowed = () => typeof _loggedIn !== 'undefined' && _loggedIn &&
    typeof _currentUser !== 'undefined' && _currentUser &&
    effectivePermissions(_currentUser).includes('announce');
  const esc = value => String(value == null ? '' : value).replace(/[&<>"']/g,
    ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
  const host = () => document.getElementById('annReview');
  const panel = html => '<section class="card" style="padding:24px;margin-bottom:20px;border:2px solid #0F4A3E;overflow-wrap:anywhere">' + html + '</section>';
  async function get(row) {
    const response = await gwCall('announce_get', {row: Number(row)}, {cache:'no-store'});
    const result = await response.json();
    if (!result.success || !result.announcement) throw new Error(result.error || '找不到這份公告');
    if (Number(result.announcement.row) !== Number(row)) throw new Error('公告紀錄不符，請回列表重新選擇');
    return result.announcement;
  }
  function mount() {
    generation++;
    record = null;
    if (!allowed() || !host()) return;
    if (!document.getElementById('annMobileMenu')) {
      const sidebar = document.querySelector('.sidebar');
      if (sidebar) {
        const button = document.createElement('button');
        button.id = 'annMobileMenu';
        button.className = 'btn btn-outline';
        button.textContent = '展開後台選單';
        button.setAttribute('aria-expanded', 'false');
        button.onclick = () => {
          const expanded = document.body.classList.toggle('ann-menu-expanded');
          button.setAttribute('aria-expanded', String(expanded));
          button.textContent = expanded ? '收起後台選單' : '展開後台選單';
        };
        sidebar.insertBefore(button, sidebar.firstChild);
      }
    }
    if (params.get('section') !== 'announce' || !params.has('review_row')) return;
    const rows = params.getAll('review_row');
    if (rows.length !== 1 || !validRow(rows[0])) {
      host().innerHTML = panel('<h2>公告連結格式不完整</h2><p>請從下方公告列表選擇，或重新開啟主管收到的通知。</p>');
      return;
    }
    return open(Number(rows[0]), params.get('review_action') === 'reject' ? 'reject' : 'approve');
  }
  async function open(row, nextIntent) {
    if (!allowed() || !host() || !validRow(row) || busy) return;
    const container = host(), owner = _currentUser, ticket = ++generation;
    record = null;
    intent = nextIntent === 'reject' ? 'reject' : 'approve';
    container.innerHTML = panel('<p role="status">正在讀取公告全文，尚未執行審核…</p>');
    try {
      const result = await get(row);
      if (ticket !== generation || owner !== _currentUser || !allowed() || host() !== container) return;
      record = result;
      draw();
      container.scrollIntoView({behavior: 'auto', block: 'start'});
    } catch (error) {
      if (ticket !== generation || !allowed() || host() !== container) return;
      container.innerHTML = panel('<h2>暫時無法讀取公告</h2><p role="alert">' + esc(error.message) +
        '</p><button class="btn btn-outline" id="annReviewRetry">重新讀取</button>');
      document.getElementById('annReviewRetry').onclick = () => open(row, intent);
    }
  }
  function draw() {
    if (!host() || !record || !allowed()) return;
    const a = record, pending = a.status === '待審核' && blockedRow !== Number(a.row);
    host().innerHTML = panel(
      '<p style="color:#52645e;font-size:14px">公告審核 / ' + esc(a.no) + '</p>' +
      '<h2 style="margin:12px 0;font-size:24px">' + esc(a.title) + '</h2>' +
      '<p>狀態：<strong>' + esc(a.status) + '</strong>　類別：' + esc(a.category) + '</p>' +
      '<p>提報人：' + esc(a.submitter) + '　對象：' + esc(a.target || '未指定') + '</p>' +
      '<div style="white-space:pre-wrap;line-height:1.85;padding:20px 0;border-top:1px solid #d9e3df;border-bottom:1px solid #d9e3df">' + esc(a.content) + '</div>' +
      (pending ?
        '<p>請核對全文。開啟連結不會自動同意或拒絕。</p>' +
        '<p style="font-size:14px">同意後系統會嘗試產生 PDF 並通知提報人，不會自動推送全體員工群。</p>' +
        '<label for="annReviewReason">拒絕原因（拒絕時必填）</label>' +
        '<textarea id="annReviewReason" rows="3" style="display:block;box-sizing:border-box;width:100%;margin:8px 0 16px;padding:12px;font:inherit;border:1px solid #a6bab1;border-radius:8px"></textarea>' +
        '<div style="display:flex;gap:12px;flex-wrap:wrap"><button id="annReviewApprove" class="btn btn-green" style="min-height:44px">確認同意發布</button>' +
        '<button id="annReviewReject" class="btn btn-outline" style="min-height:44px">確認拒絕</button></div>' :
        '<p>此公告已處理，或上次送出結果尚待查核。此處不提供重複審核。</p>') +
      '<p id="annReviewMessage" role="status" style="white-space:pre-wrap"></p>');
    if (pending) {
      document.getElementById('annReviewApprove').onclick = () => decide('approve');
      document.getElementById('annReviewReject').onclick = () => decide('reject');
      if (intent === 'reject') document.getElementById('annReviewReason').focus({preventScroll:true});
    }
  }
  async function decide(action) {
    if (!allowed() || busy || !record || record.status !== '待審核' || blockedRow === Number(record.row)) return;
    const snapshot = record, container = host(), owner = _currentUser;
    const reason = (document.getElementById('annReviewReason').value || '').trim();
    const message = document.getElementById('annReviewMessage');
    const approver = String(owner.name || '').trim();
    if (!approver) { message.textContent = '登入帳號缺少姓名，請聯絡管理員。'; return; }
    if (action === 'reject' && !reason) { message.textContent = '請填寫拒絕原因。'; return; }
    busy = true;
    container.querySelectorAll('button,textarea').forEach(el => el.disabled = true);
    let sent = false;
    try {
      const fresh = await get(snapshot.row);
      if (!allowed() || owner !== _currentUser || host() !== container) return;
      const changed = ['row','no','title','content','submitter','target','category','status'].some(k => fresh[k] !== snapshot[k]);
      if (changed) {
        record = fresh;
        draw();
        document.getElementById('annReviewMessage').textContent = '公告內容或狀態已變更，請重新核對。尚未送出審核。';
        return;
      }
      const text = action === 'approve' ?
        '確定同意發布？系統會嘗試產生 PDF 並通知提報人，不會自動推送全體員工群。' :
        '確定拒絕此公告？\n拒絕原因：' + reason;
      if (!confirm(snapshot.no + '｜' + snapshot.title + '\n審核人：' + approver + '\n' + text)) return;
      sent = true;
      const response = await gwCall(action === 'approve' ? 'announce_approve' : 'announce_reject',
        {row:Number(snapshot.row), approver, ...(action === 'reject' ? {reason} : {})});
      const result = await response.json();
      if (!result.success) throw new Error(result.error || '審核未完成');
      blockedRow = Number(snapshot.row);
      if (!allowed() || owner !== _currentUser || host() !== container) return;
      record = {...snapshot, status:action === 'approve' ? '已發布' : '已拒絕'};
      draw();
      document.getElementById('annReviewMessage').textContent = action === 'approve' ?
        (result.pdfUrl ? '審核已通過。PDF 與通知狀態請於公告列表查核。' : '審核已通過，但未取得 PDF 連結。請於公告列表確認 PDF 與通知狀態。') : '已拒絕此公告。';
      _annRenderBody();
    } catch (error) {
      if (sent) blockedRow = Number(snapshot.row);
      if (!allowed() || owner !== _currentUser || host() !== container) return;
      if (sent) draw();
      document.getElementById('annReviewMessage').textContent = sent ?
        '送出後未能確認結果，請先查核公告列表，不要重複送出。' + error.message :
        '無法確認最新狀態，尚未送出審核。' + error.message;
    } finally {
      busy = false;
      if (host() === container && allowed()) container.querySelectorAll('button,textarea').forEach(el => el.disabled = false);
    }
  }
  window.AnnReview = {mount, open};
})();
