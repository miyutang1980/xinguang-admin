/* Student master and semester assignments. No student data is persisted in the browser. */
(function () {
  'use strict';
  const CURRENT = '115-1';
  const DISPLAY_NAME = '學生中英文姓名'; // Derived W column, never an editable/persisted field.
  const MASTER = ['學生編號','學生帳號','學生姓名','英文名','性別','生日','身分證字號','Email',
    '住家地址','備註/飲食禁忌','媽媽姓名','媽媽手機','媽媽公司電話','媽媽工作單位','媽媽Email',
    '爸爸姓名','爸爸手機','爸爸公司電話','爸爸工作單位','爸爸Email','媽媽LINE userId','爸爸LINE userId','建立時間','更新時間','弋果分校'];
  // 17 yellow columns: semester + the 16 semester-specific enrollment fields.
  const YELLOW = ['學期','弋果班級','弋果課程','學生類別','學校','年級','小學班級','外籍教師',
    '中籍教師','TXClass','學年度起始日','課堂時間','課程類別','教室','學期狀態','課後輔導','交通車'];
  const ASSIGN = ['指派編號','學期','學生編號','學生帳號','學生姓名', ...YELLOW.slice(1), '更新時間'];
  const CLASS_FIELDS = ['TXClass','弋果課程','課程類別','中籍教師','外籍教師','學年度起始日','課堂時間','教室'];
  const state = { master: [], assignments: [], semester: CURRENT, semesters: [CURRENT, '114-2'],
    keyword: '', status: '', classCode: '', masterKeyword: '', request: 0, ready: false, active: false, kind: '', owner: null };
  const esc = value => String(value == null ? '' : value).replace(/[&<>"']/g, c =>
    ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'}[c]));
  const val = (row, key) => String(row[key] == null ? '' : row[key]);
  const unique = values => [...new Set(values.filter(Boolean))].sort((a,b) => a.localeCompare(b, 'zh-Hant'));
  const same = (a,b) => JSON.stringify(a) === JSON.stringify(b);
  const session = () => typeof _currentUser === 'object' ? _currentUser : null;
  const owns = () => state.owner && state.owner === session();
  const visible = () => owns() && (typeof currentSection === 'undefined' ||
    currentSection === (state.kind === 'master' ? 'students' : 'classAssign'));
  let dialog = null;
  let focusBeforeDialog = null;
  let saving = false;
  let pendingRender = null;
  const timings = [];

  // Memory-only diagnostics: never retain URLs, credentials, rows or payloads.
  function recordTiming(action, started, ok) {
    timings.push({ action, milliseconds: Math.round(performance.now() - started), ok });
    if (timings.length > 30) timings.shift();
  }

  function progress(button, initial) {
    const started = performance.now();
    let label = initial;
    const paint = () => {
      if (button.isConnected) button.textContent = `${label} · ${Math.floor((performance.now() - started) / 1000)} 秒`;
    };
    paint();
    const timer = setInterval(paint, 1000);
    return {
      stage(text) { label = text; paint(); },
      stop() { clearInterval(timer); },
      elapsed() { return ((performance.now() - started) / 1000).toFixed(1); }
    };
  }

  function auth() {
    const u = session();
    if (!u || !u.username || !['admin','staff','teacher'].includes(u.role)) throw new Error('請以管理員、行政或教師帳號重新登入。');
    let password = '';
    try { password = sessionStorage.getItem('xg_session_pwd') || localStorage.getItem('xg_admin_pwd') || ''; } catch (_) {}
    if (!password) throw new Error('登入憑證已失效，請登出後重新登入。');
    return { adminUser: u.username, adminPass: password };
  }

  async function api(action, payload = {}) {
    const started = performance.now();
    let ok = false;
    const params = Object.assign({}, payload, auth());
    const query = new URLSearchParams({ _action: action });
    Object.keys(params).forEach(key => query.set(typeof params[key] === 'object' ? key + '_json' : key,
      typeof params[key] === 'object' ? JSON.stringify(params[key]) : String(params[key])));
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 45000);
    try {
      const response = await fetch(GW_URL + '?' + query, { cache: 'no-store', signal: controller.signal });
      if (!response.ok) throw new Error('服務連線失敗（HTTP ' + response.status + '）。');
      const result = await response.json();
      if (!result || Array.isArray(result) || result.success !== true) {
        throw new Error(result && result.error ? String(result.error) : '後端尚未升級或回應格式不符，已停止操作。');
      }
      ok = true;
      return result;
    } catch (error) {
      if (error.name === 'AbortError') throw new Error('連線逾時；若剛才正在儲存，結果尚未確認，請重新整理核對後再操作。');
      if (error instanceof SyntaxError) throw new Error('後端回傳的不是有效資料，請管理員確認 Gateway 部署版本。');
      if (error instanceof TypeError) throw new Error('無法連線；若剛才正在儲存，請重新整理核對結果，勿重複送出。');
      throw error;
    } finally {
      clearTimeout(timer); recordTiming(action, started, ok);
      if (!action.endsWith('_list') && typeof window.invalidatePickupRoster === 'function') {
        window.invalidatePickupRoster();
      }
    }
  }

  function validateList(result, kind, semester) {
    const headers = kind === 'master' ? MASTER : ASSIGN;
    if (typeof result.active !== 'boolean' || result.currentSemester !== CURRENT) {
      throw new Error('後端缺少明確啟用狀態或營運學期不符，已停止操作。請管理員核對 Gateway 設定。');
    }
    if (!Array.isArray(result.list)) throw new Error('後端缺少 list 清單；不是零筆資料。請升級 Gateway。');
    const seen = new Set();
    result.list.forEach(row => {
      if (!row || !Number.isInteger(Number(row._row)) || Number(row._row) < 2 ||
          headers.some(key => !Object.prototype.hasOwnProperty.call(row, key)) || !row['學生編號'] ||
          !row['學生姓名']) throw new Error('資料欄位或列號不完整，已停用編輯，請管理員檢查資料移轉。');
      const key = kind === 'master' ? row['學生編號'] : row['指派編號'];
      if (!key || seen.has(key)) throw new Error('資料識別碼重複或空白，請管理員先修復，禁止覆寫。');
      seen.add(key);
      if (kind === 'assignment' && (row['學期'] !== semester || key !== semester + '|' + row['學生編號'])) {
        throw new Error('後端回傳其他學期或不一致的指派資料，已停止操作。');
      }
    });
    if (kind === 'assignment' && (result.semester !== semester || !Array.isArray(result.semesters))) {
      throw new Error('後端學期回應不完整，請升級 Gateway；不會回退到舊名冊。');
    }
    return result.list;
  }
  async function readList(kind, semester = state.semester) {
    const result = await api(kind === 'master' ? 'student_master_list' : 'semester_assignments_list',
      kind === 'master' ? {} : { semester });
    validateList(result, kind, semester);
    return result;
  }
  function requireActive(result) {
    if (result.active !== true) throw new Error('新資料模型尚未啟用，已停止寫入。請管理員完成 Gateway 部署、資料核對與啟用。');
  }

  const banner = (title, text, yellow = false) => `<div class="ss-banner${yellow ? ' ss-yellow' : ''}"><h3>${esc(title)}</h3><p>${esc(text)}</p></div>`;
  function button(text, id, primary = false) {
    return `<button type="button" class="btn ${primary ? 'btn-green' : 'btn-outline'}" id="${id}">${esc(text)}</button>`;
  }
  function table(headers, rows) {
    return `<div class="ss-table-scroll" tabindex="0" aria-label="資料表，可橫向捲動"><table class="data-table"><thead><tr>${
      headers.map(h => `<th scope="col">${esc(h)}</th>`).join('')}</tr></thead><tbody>${rows ||
      `<tr><td colspan="${headers.length}" class="ss-empty">沒有符合條件的資料</td></tr>`}</tbody></table></div>`;
  }
  const option = (value, selected, label = value) => `<option value="${esc(value)}"${value === selected ? ' selected' : ''}>${esc(label)}</option>`;

  function shell(el, kind) {
    const historical = state.semester !== CURRENT;
    el.classList.add('ss-host');
    el.innerHTML = banner(kind === 'master' ? '學生主檔｜跨學期沿用' : `學期班級指派｜${state.semester}${historical ? ' 歷史唯讀' : ' 新學期'}`,
      kind === 'master' ? '綠色資料：學生與家長基本資料，只建一次。班級、老師、接送與就學狀態請到「學期班級指派」維護。' :
        historical ? '歷史資料僅供查閱，不能新增、修改或整班升級。切換回 115-1 才能維護新學期。' :
          '黃色資料：每位學生每學期一筆。115-1 是目前營運學期；114-2 保留歷史，不覆蓋、不搬移。', kind !== 'master') +
      `<div class="ss-tools" id="ssTools"></div><div id="ssNotice" role="status" aria-live="polite"></div><div class="table-wrap" id="ssList"></div>`;
    const tools = el.querySelector('#ssTools');
    if (kind === 'master') {
      tools.innerHTML = `<label>搜尋主檔<input id="ssMasterSearch" type="search" placeholder="姓名／學號／帳號／聯絡資料" value="${esc(state.masterKeyword)}"></label>` +
        button('新增學生主檔', 'ssAdd', true) + button('重新整理', 'ssRefresh') +
        `<a class="btn btn-outline" href="STUDENT_MANUAL.md" target="_blank" rel="noopener">操作手冊</a>`;
      tools.querySelector('#ssMasterSearch').oninput = event => { state.masterKeyword = event.target.value; drawRows(el); };
    } else {
      tools.innerHTML = `<label>學期<select id="ssSemester">${state.semesters.map(s => option(s, state.semester,
        s + (s === CURRENT ? ' · 新學期（可編輯）' : ' · 歷史（唯讀）'))).join('')}</select></label>
        <label>學期狀態<select id="ssStatus">${option('', state.status, '全部狀態')}${
          unique(state.assignments.map(r => r['學期狀態'])).map(s => option(s, state.status)).join('')}</select></label>
        <label>班代號（弋果班級）<select id="ssClass">${option('', state.classCode, '全部班代號')}${
          unique(state.assignments.map(r => r['弋果班級'])).map(s => option(s, state.classCode)).join('')}</select></label>
        <label>搜尋指派<input id="ssSearch" type="search" placeholder="中文／英文名／學號／老師／TXClass" value="${esc(state.keyword)}"></label>` +
        (historical ? '' : button('新增學期指派', 'ssAdd', true) + button('整班升級／調整', 'ssBulk')) +
        button('重新整理', 'ssRefresh');
      tools.querySelector('#ssSemester').onchange = event => {
        state.semester = event.target.value; state.status = ''; state.classCode = ''; state.keyword = '';
        render(el, 'assignment');
      };
      tools.querySelector('#ssStatus').onchange = event => { state.status = event.target.value; drawRows(el); };
      tools.querySelector('#ssClass').onchange = event => { state.classCode = event.target.value; drawRows(el); };
      tools.querySelector('#ssSearch').oninput = event => { state.keyword = event.target.value; drawRows(el); };
      if (!historical) tools.querySelector('#ssBulk').onclick = () => openBulk(el);
    }
    tools.querySelector('#ssRefresh').onclick = () => render(el, kind);
    if (tools.querySelector('#ssAdd')) tools.querySelector('#ssAdd').onclick = () => openEditor(el, null);
    if (!state.active) {
      ['#ssAdd','#ssBulk'].forEach(id => { const b = tools.querySelector(id); if (b) b.remove(); });
      el.querySelector('#ssTools').insertAdjacentHTML('beforebegin', banner('資料已暫存 · 尚未啟用',
        '新資料模型尚未啟用，目前僅供核對，不能新增或修改。請管理員完成 Gateway 部署、資料移轉驗收及啟用；既有營運名冊不會被此頁覆寫。', true));
    }
    drawRows(el);
  }

  function drawRows(el) {
    if (!owns() || !state.ready) return;
    const master = state.kind === 'master';
    const all = master ? state.master : state.assignments;
    const keyword = (master ? state.masterKeyword : state.keyword).trim().toLocaleLowerCase();
    const filtered = all.filter(row => (master || ((!state.status || row['學期狀態'] === state.status) &&
      (!state.classCode || row['弋果班級'] === state.classCode))) &&
      (!keyword || (master ? MASTER : ASSIGN.concat(DISPLAY_NAME)).map(h => val(row, h)).join(' ').toLocaleLowerCase().includes(keyword)));
    const headers = master ? ['學生編號','學生姓名','英文名','學生帳號','媽媽姓名','媽媽手機','爸爸姓名','爸爸手機','操作'] :
      ['學生編號',DISPLAY_NAME,'弋果班級','TXClass','弋果課程','學期狀態','學校','年級','小學班級','中籍教師','外籍教師','交通車','操作'];
    el.querySelector('#ssNotice').textContent = `顯示 ${filtered.length} / ${all.length} 筆` +
      (!all.length ? (master ? ' · 主檔尚未建檔；如與預期不符，請先確認移轉結果。' :
        ` · ${state.semester} 尚無指派；不會自動複製歷史資料。`) : '');
    el.querySelector('#ssList').innerHTML = table(headers, filtered.map(row => `<tr>${
      headers.slice(0,-1).map(h => { const text = h === DISPLAY_NAME ? row[h] || row['學生姓名'] : row[h]; return `<td title="${esc(text)}">${esc(text || '—')}</td>`; }).join('')}
      <td><button type="button" class="btn btn-outline ss-row-edit" data-row="${Number(row._row)}">${
        !state.active ? '查看（尚未啟用）' : master || state.semester === CURRENT ? '查看／編輯' : '查看歷史'}</button></td></tr>`).join(''));
    el.querySelectorAll('.ss-row-edit').forEach(btn => btn.onclick = () =>
      openEditor(el, all.find(row => Number(row._row) === Number(btn.dataset.row))));
  }

  function render(el, kind, verified = null) {
    // Coalesce only simultaneous page loads, not pre-write or post-write reads.
    if (!verified && pendingRender && pendingRender.el === el && pendingRender.kind === kind &&
        pendingRender.semester === state.semester && pendingRender.owner === session() &&
        pendingRender.request === state.request) return pendingRender.promise;
    const load = { el, kind, semester: state.semester, owner: session() };
    load.promise = renderView(el, kind, verified).finally(() => {
      if (pendingRender === load) pendingRender = null;
    });
    load.request = state.request;
    pendingRender = load;
    return load.promise;
  }

  async function renderView(el, kind, verified) {
    const request = ++state.request;
    const owner = session();
    const semester = state.semester;
    closeDialog(true);
    state.kind = kind; state.ready = false; state.active = false; state.owner = owner;
    state.master = []; state.assignments = [];
    el.classList.add('ss-host');
    el.innerHTML = banner(kind === 'master' ? '學生主檔' : '學期班級指派', '正在讀取正式資料，請稍候…', kind !== 'master');
    const status = progress(el.querySelector('.ss-banner p'), '正在讀取正式資料');
    try {
      // Use the just-verified server response; do not issue a fourth round trip.
      const result = verified || await readList(kind, semester);
      validateList(result, kind, semester);
      if (request !== state.request || owner !== session() || !visible()) return;
      if (kind === 'master') state.master = result.list;
      else {
        state.assignments = result.list;
        state.semesters = unique([CURRENT, '114-2', ...result.semesters]).sort().reverse();
      }
      state.active = result.active === true;
      state.ready = true;
      shell(el, kind);
      if (!verified) el.querySelector('#ssNotice').textContent += ` · 讀取 ${status.elapsed()} 秒`;
      if (window.matchMedia('(max-width:768px)').matches) {
        requestAnimationFrame(() => {
          if (!visible()) return;
          const nav = document.querySelector('.sidebar-menu');
          if (nav) nav.scrollLeft = 0;
        });
      }
    } catch (error) {
      if (request !== state.request || owner !== session() || !visible()) return;
      el.innerHTML = banner('無法載入｜已停用寫入', '這不是空名冊。未取得完整資料前，不提供新增、修改或整班升級。', true) +
        `<div class="ss-error" role="alert">${esc(error.message)}<p>請確認登入與網路；若後端尚未升級，請管理員完成 Gateway 部署。不要改用舊名冊覆寫。</p></div>` +
        button('重新讀取', 'ssRetry');
      el.querySelector('#ssRetry').onclick = () => render(el, kind);
    } finally { status.stop(); }
  }

  function showDialog(title, html) {
    closeDialog(true);
    focusBeforeDialog = document.activeElement;
    dialog = document.createElement('dialog');
    dialog.className = 'ss-dialog';
    dialog.setAttribute('aria-labelledby', 'ssDialogTitle');
    dialog.innerHTML = `<div class="modal-header"><h3 id="ssDialogTitle">${esc(title)}</h3><button type="button" class="modal-close" id="ssClose" aria-label="關閉">×</button></div><div class="modal-body">${html}</div>`;
    document.body.appendChild(dialog);
    dialog.querySelector('#ssClose').onclick = () => closeDialog();
    dialog.addEventListener('cancel', event => { event.preventDefault(); closeDialog(); });
    dialog.showModal();
    return dialog;
  }
  function closeDialog(force = false) {
    if (saving && !force) return;
    if (dialog) { dialog.close(); dialog.remove(); dialog = null; }
    if (focusBeforeDialog && focusBeforeDialog.isConnected) focusBeforeDialog.focus();
  }
  function message(form, text) {
    form.querySelector('.ss-form-error').textContent = text;
  }
  const hints = {
    '學生編號': '系統產生，建立後不可修改', '建立時間': '由系統建立', '更新時間': '由系統更新',
    '學期': '固定 115-1，歷史學期不可更改', '弋果班級': '穩定班代號；不是課程級數',
    'TXClass': '課程級數／班級顯示名稱；整班升級請使用整班調整',
    '年級': '就讀學校年級，不等同 TXClass', '生日': '保留原格式；建議 YYYY-MM-DD',
    '學年度起始日': '建議 YYYY-MM-DD'
  };
  const suggestions = {
    '性別': ['男','女','其他'], '學期狀態': ['在學','在讀','停學','休學','退學','離校','畢業','結業','未入學'],
    '交通車': ['交通車','家長接送','Walk 走路','公務車'], '課後輔導': ['NO','課輔','托育']
  };
  function field(key, value, index, readonly, yellow, required = false) {
    const id = 'ssField' + index;
    const list = suggestions[key];
    return `<label class="ss-field${yellow ? ' ss-field-yellow' : ''}" for="${id}"><span>${esc(key)}${required ? ' *' : ''}</span>
      ${key === '備註/飲食禁忌' || key === '住家地址' ? `<textarea id="${id}" data-field="${esc(key)}"${readonly ? ' readonly' : ''} rows="2">${esc(value)}</textarea>` :
        `<input id="${id}" data-field="${esc(key)}" value="${esc(value)}" type="text"${readonly ? ' readonly' : ''}${required ? ' required' : ''}${list && !readonly ? ` list="${id}Options"` : ''} autocomplete="off">`}
      ${list && !readonly ? `<datalist id="${id}Options">${list.map(s => option(s, '')).join('')}</datalist>` : ''}
      ${hints[key] ? `<small>${esc(hints[key])}</small>` : ''}</label>`;
  }
  function getFields(form, allowed) {
    const fields = {};
    form.querySelectorAll('[data-field]').forEach(input => {
      if (allowed.includes(input.dataset.field)) fields[input.dataset.field] = input.value.trim();
    });
    return fields;
  }
  function mustEdit(kind, semester) {
    if (!owns() || !state.ready || state.kind !== kind || !visible()) throw new Error('頁面已變更，請重新載入後再編輯。');
    if (!state.active) throw new Error('新資料模型尚未啟用，禁止寫入。');
    if (kind === 'assignment' && (semester !== CURRENT || state.semester !== CURRENT)) throw new Error('歷史學期為唯讀，不能修改。');
    auth();
  }

  async function openEditor(el, original) {
    if (!owns() || !state.ready || saving) return;
    const kind = state.kind;
    const master = kind === 'master';
    const historical = !master && state.semester !== CURRENT;
    const readonly = historical || !state.active;
    const semester = state.semester;
    const request = state.request;
    let choices = [];
    if (!original && readonly) return;
    if (!master && !original) {
      const trigger = el.querySelector('#ssAdd');
      trigger.disabled = true;
      try {
        const result = await readList('master');
        requireActive(result);
        if (request !== state.request || !visible()) return;
        choices = result.list.filter(row => !state.assignments.some(a => a['學生編號'] === row['學生編號']));
        if (!choices.length) {
          el.querySelector('#ssNotice').textContent = '沒有可新增的學生：請先在學生主檔建檔，或確認該生是否已有本學期指派。';
          return;
        }
      } catch (error) {
        if (request === state.request && visible()) el.querySelector('#ssNotice').textContent = '無法載入學生主檔，已阻擋新增：' + error.message;
        return;
      } finally { trigger.disabled = false; }
    }
    const row = original ? {...original} : master ? {} : { '學期': CURRENT, '學期狀態': '在學' };
    const headers = master ? MASTER : ASSIGN;
    const body = `<form id="ssEditor">
      ${!master && original && row[DISPLAY_NAME] ? `<p class="ss-help">學生中英文姓名（Google W 欄自動合併，唯讀）：${esc(row[DISPLAY_NAME])}</p>` : ''}
      <p class="ss-help">${!state.active ? '新資料模型尚未啟用，此處僅供核對，所有欄位唯讀。' : master ? '原 24 個主檔欄位及弋果分校完整呈現；學生編號與兩個系統時間為唯讀，其餘 22 欄可編輯。' :
        historical ? '歷史快照唯讀；此處不會修改 115-1 或學生主檔。' : '17 個黃色學期欄位；學期固定 115-1。學生身分由主檔帶入。個別轉班改「弋果班級」，整班升級請用列表的「整班升級／調整」。'}</p>
      ${!master && !original ? `<label class="ss-field">選擇已建檔學生 *
        <input id="ssStudentSearch" type="search" placeholder="先輸入中文、英文名或學號篩選">
        <select id="ssStudentChoice" required>${option('', '', '請選擇學生')}${choices.map(r => option(r['學生編號'], '', r['學生編號'] + ' · ' + r['學生姓名'] + (r['英文名'] ? '（' + r['英文名'] + '）' : '') + (r['學生帳號'] ? ' · ' + r['學生帳號'] : ''))).join('')}</select></label>` : ''}
      <div class="ss-fields">${headers.map((key,index) => field(key, row[key] || '', index,
        readonly || (master ? ['學生編號','建立時間','更新時間'].includes(key) : !YELLOW.includes(key) || key === '學期'),
        !master && YELLOW.includes(key), master && key === '學生姓名')).join('')}</div>
      <div class="ss-form-error" role="alert"></div><div class="ss-actions">${button(readonly ? '關閉' : '取消', 'ssCancel')}${
        readonly ? '' : '<button class="btn btn-green" id="ssSave" type="submit">儲存' + (original ? '修改' : master ? '主檔' : '指派') + '</button>'}</div></form>`;
    const modal = showDialog((historical ? '歷史指派 · ' + semester : original ? '查看／編輯' : '新增') +
      (master ? '學生主檔' : '學期指派'), body);
    const form = modal.querySelector('#ssEditor');
    form.querySelector('#ssCancel').onclick = () => closeDialog();
    if (!master && !readonly) {
      const classInput = form.querySelector('[data-field="弋果班級"]');
      classInput.setAttribute('list', 'ssExistingClasses');
      classInput.insertAdjacentHTML('afterend', `<datalist id="ssExistingClasses">${
        unique(state.assignments.map(r => r['弋果班級'])).map(c => option(c, '')).join('')}</datalist>`);
      classInput.onchange = () => {
        const peers = classRows(state.assignments, classInput.value.trim()).filter(r => !original || r._row !== original._row);
        CLASS_FIELDS.forEach(key => {
          const values = unique(peers.map(r => val(r,key)));
          form.querySelector(`[data-field="${key}"]`).value = values.length === 1 ? values[0] : '';
        });
        message(form, peers.length ? '已帶入目標班代號的共用資料，請核對後再儲存。TXClass 如有歧異請先使用整班調整。' :
          '此班代號尚無其他指派，已清除舊班共用欄位；請填入新班資料並核對。');
      };
    }
    if (!master && !original) {
      const select = form.querySelector('#ssStudentChoice');
      form.querySelector('#ssStudentSearch').oninput = event => {
        const query = event.target.value.trim().toLocaleLowerCase();
        const selected = select.value;
        select.innerHTML = option('', '', '請選擇學生') + choices.filter(r =>
          r['學生編號'] === selected || [r['學生編號'],r['學生姓名'],r['英文名'],r['學生帳號']].join(' ').toLocaleLowerCase().includes(query))
          .map(r => option(r['學生編號'], selected, r['學生編號'] + ' · ' + r['學生姓名'] + (r['英文名'] ? '（' + r['英文名'] + '）' : ''))).join('');
      };
      select.onchange = () => {
        const student = choices.find(r => r['學生編號'] === select.value);
        ['學生編號','學生帳號','學生姓名'].forEach(key => { form.querySelector(`[data-field="${key}"]`).value = student ? student[key] : ''; });
        form.querySelector('[data-field="指派編號"]').value = student ? CURRENT + '|' + student['學生編號'] : '';
      };
    }
    let reconcileRequired = false;
    form.onsubmit = async event => {
      event.preventDefault();
      if (readonly || saving || reconcileRequired) return;
      const save = form.querySelector('#ssSave');
      let writeAttempted = false;
      let status;
      try {
        mustEdit(kind, semester);
        const allowed = master ? MASTER.filter(k => !['學生編號','建立時間','更新時間'].includes(k)) : YELLOW.slice(1);
        const fields = getFields(form, allowed);
        if (master) {
          fields['學生姓名'] = fields['學生姓名'].trim();
          if (!fields['學生姓名']) throw new Error('學生姓名不可空白。');
        } else if (!original) {
          const student = choices.find(r => r['學生編號'] === form.querySelector('#ssStudentChoice').value);
          if (!student) throw new Error('請先從學生主檔選擇學生。');
          Object.assign(fields, { '學期': CURRENT, '學生編號': student['學生編號'], '學生帳號': student['學生帳號'], '學生姓名': student['學生姓名'] });
        }
        const changed = original ? Object.fromEntries(Object.entries(fields).filter(([k,v]) => v !== val(original,k))) : fields;
        if (!Object.keys(changed).length) { message(form, '沒有變更，尚未送出。'); return; }
        saving = true; form.querySelectorAll('input,select,textarea,button').forEach(i => i.disabled = true);
        status = progress(save, '1/3 核對最新資料'); message(form, '');
        // Re-read before row_index writes to avoid targeting moved rows and stale edits.
        const latest = await readList(kind, semester);
        requireActive(latest);
        mustEdit(kind, semester);
        if (request !== state.request || !form.isConnected) throw new Error('頁面已變更，未送出儲存。請重新開啟編輯。');
        if (original) {
          const identity = master ? '學生編號' : '指派編號';
          const fresh = latest.list.find(r => r[identity] === original[identity]);
          if (!fresh || Number(fresh._row) !== Number(original._row) ||
              headers.some(k => val(fresh,k) !== val(original,k))) throw new Error('資料已被其他人修改或列號已變動。請關閉視窗、重新整理後重試。');
        } else if (!master && latest.list.some(r => r['學生編號'] === fields['學生編號'])) {
          throw new Error('此學生本學期已有指派，請重新整理後編輯既有紀錄。');
        }
        const action = master ? 'student_master_' + (original ? 'update' : 'create') :
          'semester_assignment_' + (original ? 'update' : 'create');
        const payload = { fields: changed };
        if (original) payload.row_index = Number(original._row);
        status.stage('2/3 寫入及同步');
        writeAttempted = true;
        const result = await api(action, payload);
        if (!original && (master ? !result.studentNo : !result.assignmentNo)) {
          throw new Error('後端未回傳建立識別碼，結果尚未確認；請關閉視窗並重新整理，勿直接重送。');
        }
        // A nominal success is not enough: verify the persisted fields by a fresh read.
        status.stage('3/3 讀回驗證');
        const checked = await readList(kind, semester);
        const identity = master ? '學生編號' : '指派編號';
        const id = original ? original[identity] : master ? result.studentNo : result.assignmentNo;
        const persisted = checked.list.find(r => r[identity] === id);
        if (!persisted || Object.entries(changed).some(([k,v]) => val(persisted,k) !== String(v))) {
          throw new Error('儲存回應已收到，但重新讀取尚未核對成功。請關閉視窗、重新整理確認，勿重複新增。');
        }
        if (request === state.request && visible() && form.isConnected) {
          closeDialog(true);
          await render(el, kind, checked);
          if (state.ready && visible()) el.querySelector('#ssNotice').textContent += ` · 已儲存並讀回確認，共 ${status.elapsed()} 秒`;
        }
      } catch (error) {
        reconcileRequired = writeAttempted;
        if (form.isConnected) message(form, error.message + (writeAttempted ? '\n請關閉視窗並重新整理核對後再編輯；本視窗已停用再次送出。' : ''));
      } finally {
        if (status) status.stop();
        saving = false; form.querySelectorAll('input,select,textarea,button').forEach(i => i.disabled = false);
        save.disabled = reconcileRequired;
        save.textContent = reconcileRequired ? '請先關閉並核對結果' : '儲存';
      }
    };
  }

  function classRows(rows, code) { return rows.filter(r => r['弋果班級'] === code); }
  function snapshot(rows) { return rows.map(row => ASSIGN.map(h => val(row,h)).concat(Number(row._row))).sort((a,b) => a[0].localeCompare(b[0])); }
  function openBulk(el) {
    try { mustEdit('assignment', state.semester); } catch (_) { return; }
    if (saving) return;
    const request = state.request;
    let reconcileRequired = false;
    const codes = unique(state.assignments.map(r => r['弋果班級']));
    if (!codes.length) { el.querySelector('#ssNotice').textContent = '本學期沒有班代號，請先建立指派及班代號。'; return; }
    const modal = showDialog('整班升級／調整 · ' + CURRENT, `<form id="ssBulkForm">
      <p class="ss-help">依穩定的「弋果班級」班代號更新本學期全班所有指派（含暫停／離校），不受列表搜尋與狀態篩選影響。不修改學生主檔、班代號、小學年級或 114-2 歷史。</p>
      <label class="ss-field">班代號（弋果班級）<select id="ssBulkClass">${codes.map(c => option(c, state.classCode)).join('')}</select></label>
      <p class="ss-help">只勾選要修改的欄位。升級請勾選 TXClass 並輸入目標級數；不會自動推算。勾選後留空代表清空該欄。</p>
      <div class="ss-fields">${CLASS_FIELDS.map((key,index) => `<div class="ss-bulk-field"><label class="ss-check"><input type="checkbox" data-apply="${esc(key)}">修改 ${esc(key)}</label>${field(key, '', index, false, true)}</div>`).join('')}</div>
      <div id="ssBulkPreview"></div><div class="ss-form-error" role="alert"></div>
      <div class="ss-actions">${button('取消', 'ssCancel')}${button('預覽全班變更', 'ssPreview', true)}</div></form>`);
    const form = modal.querySelector('#ssBulkForm');
    form.onsubmit = event => event.preventDefault();
    form.querySelector('#ssCancel').onclick = () => closeDialog();
    let preview = null;
    const invalidate = () => { preview = null; form.querySelector('#ssBulkPreview').innerHTML = ''; message(form, ''); };
    form.querySelectorAll('input,select').forEach(input => input.addEventListener('input', invalidate));
    form.querySelector('#ssPreview').onclick = async () => {
      if (saving || reconcileRequired) return;
      const previewButton = form.querySelector('#ssPreview');
      try {
        mustEdit('assignment', CURRENT);
        invalidate();
        const code = form.querySelector('#ssBulkClass').value;
        const applied = [...form.querySelectorAll('[data-apply]:checked')].map(i => i.dataset.apply);
        if (!applied.length) throw new Error('請至少勾選一個要更新的欄位。');
        const fields = getFields(form, applied);
        previewButton.disabled = true; previewButton.textContent = '讀取最新全班名單…';
        const latest = await readList('assignment', CURRENT);
        requireActive(latest);
        mustEdit('assignment', CURRENT);
        if (!form.isConnected) return;
        if (form.querySelector('#ssBulkClass').value !== code ||
            !same([...form.querySelectorAll('[data-apply]:checked')].map(i => i.dataset.apply), applied) ||
            !same(getFields(form, applied), fields)) throw new Error('預覽讀取期間欄位已變更，請重新預覽。');
        const rows = classRows(latest.list, code);
        if (!rows.length) throw new Error('此班代號已無指派，請重新整理。');
        const changedRows = rows.filter(r => applied.some(k => val(r,k) !== fields[k]));
        if (!changedRows.length) throw new Error('選取欄位與全班目前資料相同，沒有需要更新的資料。');
        preview = { code, fields, rows, snapshot: snapshot(rows) };
        form.querySelector('#ssBulkPreview').innerHTML = `<div class="ss-banner ss-yellow"><h3>送出前確認</h3>
          <p>${CURRENT} · 班代號 ${esc(code)} · 全班 ${rows.length} 筆，實際有差異 ${changedRows.length} 筆。</p>
          <p>本次不動主檔、不動其他班、不動 114-2。只更新下列勾選欄位。</p></div>` +
          table(['學生編號','學生姓名',...applied], rows.map(r => `<tr><td>${esc(r['學生編號'])}</td><td>${esc(r['學生姓名'])}</td>${
            applied.map(k => `<td class="ss-diff">${esc(r[k] || '（空白）')} → <strong>${esc(fields[k] || '（清空）')}</strong></td>`).join('')}</tr>`).join('')) +
          `<label class="ss-check ss-confirm"><input type="checkbox" id="ssAcknowledged">我已核對全班名單與前後差異，確認只修改 ${CURRENT} 的 ${esc(code)}。</label>` +
          button('確認更新全班', 'ssCommitBulk', true);
        const commit = form.querySelector('#ssCommitBulk');
        commit.disabled = true;
        form.querySelector('#ssAcknowledged').onchange = event => { commit.disabled = !event.target.checked; };
        commit.onclick = async () => {
          if (!preview || saving || reconcileRequired || !form.querySelector('#ssAcknowledged').checked) return;
          const approved = preview;
          let writeAttempted = false;
          let status;
          try {
            mustEdit('assignment', CURRENT);
            if (!confirm(`確認更新 ${CURRENT}／班代號 ${approved.code} 的全班 ${approved.rows.length} 筆？\n\n欄位：${Object.keys(approved.fields).join('、')}\n114-2 歷史與學生主檔不變。`)) return;
            saving = true; form.querySelectorAll('input,select,button').forEach(i => i.disabled = true);
            status = progress(commit, '1/3 核對全班資料');
            const fresh = await readList('assignment', CURRENT);
            requireActive(fresh);
            mustEdit('assignment', CURRENT);
            if (request !== state.request || !form.isConnected) throw new Error('頁面已變更，未送出全班更新。');
            if (!same(snapshot(classRows(fresh.list, approved.code)), approved.snapshot)) {
              throw new Error('全班資料在預覽後已變更。請重新預覽，禁止套用過期名單。');
            }
            status.stage('2/3 寫入及同步全班');
            writeAttempted = true;
            const result = await api('semester_class_update', { semester: CURRENT, classCode: approved.code, fields: approved.fields });
            if (!Number.isInteger(result.updated) || result.updated < 1) {
              throw new Error('後端未提供有效更新筆數，結果尚未確認。請關閉視窗、重新整理核對，勿重複送出。');
            }
            status.stage('3/3 讀回全班驗證');
            const checked = await readList('assignment', CURRENT);
            const saved = classRows(checked.list, approved.code);
            if (saved.length !== approved.rows.length || approved.rows.some(old =>
              !saved.some(r => r['指派編號'] === old['指派編號'])) ||
                saved.some(r => Object.entries(approved.fields).some(([k,v]) => val(r,k) !== v))) {
              throw new Error('後端已回應，但全班更新結果尚未核對一致。請重新整理並請管理員檢查，勿重複送出。');
            }
            if (request === state.request && visible() && form.isConnected) {
              closeDialog(true);
              await render(el, 'assignment', checked);
              if (state.ready && visible()) el.querySelector('#ssNotice').textContent += ` · 已核對 ${approved.code} 全班 ${saved.length} 筆，共 ${status.elapsed()} 秒；歷史學期不變。`;
            }
          } catch (error) {
            reconcileRequired = writeAttempted;
            if (form.isConnected) { invalidate(); message(form, error.message + (writeAttempted ? '\n請關閉視窗並重新整理核對，勿重複送出。' : '')); }
          } finally {
            if (status) status.stop();
            saving = false; form.querySelectorAll('input,select,button').forEach(i => i.disabled = false);
            form.querySelector('#ssPreview').disabled = reconcileRequired;
          }
        };
      } catch (error) { if (form.isConnected) message(form, error.message); }
      finally { previewButton.disabled = false; previewButton.textContent = '預覽全班變更'; }
    };
  }

  window.XGStudents = {
    diagnostics: () => timings.map(entry => Object.assign({}, entry)),
    renderMaster: el => render(el, 'master'),
    renderAssignments: el => render(el, 'assignment'),
    reset: () => {
      pendingRender = null; timings.length = 0;
      ++state.request; state.ready = false; state.active = false; state.owner = null; state.master = []; state.assignments = [];
      state.semester = CURRENT; state.semesters = [CURRENT,'114-2'];
      state.keyword = ''; state.masterKeyword = ''; state.status = ''; state.classCode = '';
      closeDialog(true);
    }
  };
})();
