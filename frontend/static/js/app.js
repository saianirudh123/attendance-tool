/* ══════════════════════════════════════════════════════════════════════════
   GLOBAL STATE
   ══════════════════════════════════════════════════════════════════════════ */
let pdfFile           = null;
let xlsxFile          = null;
let registerFile      = null;
let downloadUrl       = null;
let sessionId         = null;
let currentSummary    = null;
let hasUnsavedChanges = false;
let isFrozen          = false;
let currentUser        = null;
let currentPendingId   = null;
window.empData        = [];   // live mutable employee data

const DAY_NAMES = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'];

/* ══════════════════════════════════════════════════════════════════════════
   FILE INPUT & DRAG-DROP
   ══════════════════════════════════════════════════════════════════════════ */
document.getElementById('pdf-input').addEventListener('change', function () {
  if (!this.files.length) return;
  pdfFile = this.files[0];
  document.getElementById('pdf-filename').textContent = pdfFile.name;
  document.getElementById('pdf-btn-text').textContent = 'Change PDF';
  document.getElementById('pdf-card').classList.add('has-file');
  checkReady();
});

document.getElementById('xlsx-input').addEventListener('change', function () {
  if (!this.files.length) return;
  xlsxFile = this.files[0];
  document.getElementById('xlsx-filename').textContent = xlsxFile.name;
  document.getElementById('xlsx-btn-text').textContent = 'Change XLSX';
  document.getElementById('xlsx-card').classList.add('has-file');
  checkReady();
});

document.getElementById('register-input').addEventListener('change', function () {
  if (!this.files.length) return;
  registerFile = this.files[0];
  document.getElementById('register-filename').textContent = registerFile.name;
  document.getElementById('register-btn-text').textContent = 'Change Register';
  document.getElementById('register-card').classList.add('has-file');
  checkRegisterReady();
});

['pdf-card', 'xlsx-card', 'register-card'].forEach(id => {
  const card = document.getElementById(id);
  card.addEventListener('dragover', e => { e.preventDefault(); card.style.borderColor = 'var(--accent)'; });
  card.addEventListener('dragleave', () => { if (!card.classList.contains('has-file')) card.style.borderColor = ''; });
  card.addEventListener('drop', e => {
    e.preventDefault();
    const files = e.dataTransfer.files;
    if (!files.length) return;
    const inputId = id === 'pdf-card' ? 'pdf-input' : (id === 'xlsx-card' ? 'xlsx-input' : 'register-input');
    document.getElementById(inputId).files = files;
    document.getElementById(inputId).dispatchEvent(new Event('change'));
  });
});

function checkReady() {
  document.getElementById('process-btn').disabled = !(pdfFile && xlsxFile);
}

function checkRegisterReady() {
  document.getElementById('import-register-btn').disabled = !registerFile;
}

function isAdmin() {
  return currentUser && currentUser.role === 'admin';
}

async function authFetch(url, options = {}) {
  const res = await fetch(url, options);
  if (res.status === 401) {
    window.location.href = '/login';
    throw new Error('Login required');
  }
  return res;
}

async function loadCurrentUser() {
  const res = await fetch('/api/auth/me');
  const data = await res.json();
  if (!data.authenticated) {
    window.location.href = '/login';
    return null;
  }
  currentUser = data.user;
  renderUserChrome();
  return currentUser;
}

function renderUserChrome() {
  const chip = document.getElementById('user-chip');
  const logout = document.getElementById('logout-link');
  const pendingBtn = document.getElementById('pending-toggle-btn');
  if (chip && currentUser) {
    chip.textContent = `${currentUser.role.toUpperCase()} · ${currentUser.email}`;
    chip.classList.remove('hidden');
  }
  if (logout) logout.classList.remove('hidden');
  if (pendingBtn && isAdmin()) pendingBtn.classList.remove('hidden');
}

function applyRoleControls() {
  const freezeZone = document.getElementById('freeze-zone');
  const submitZone = document.getElementById('submit-zone');
  if (!freezeZone || !submitZone) return;
  if (isAdmin()) {
    freezeZone.classList.remove('hidden');
    submitZone.classList.add('hidden');
  } else {
    freezeZone.classList.add('hidden');
    submitZone.classList.remove('hidden');
  }
}

/* ══════════════════════════════════════════════════════════════════════════
   PROCESS FILES
   ══════════════════════════════════════════════════════════════════════════ */
async function processFiles() {
  const btn      = document.getElementById('process-btn');
  const loader   = document.getElementById('btn-loader');
  const procText = document.getElementById('process-text');
  const errorBox = document.getElementById('error-box');

  errorBox.classList.add('hidden');
  btn.disabled = true;
  procText.textContent = 'Processing…';
  loader.classList.remove('hidden');

  const fd = new FormData();
  fd.append('punch_pdf',  pdfFile);
  fd.append('leave_xlsx', xlsxFile);

  try {
    const res  = await authFetch('/api/process', { method: 'POST', body: fd });
    const data = await res.json();
    if (!res.ok) throw new Error(data.detail || `Server error ${res.status}`);
    sessionId      = data.session_id;
    downloadUrl    = data.download_url;
    currentSummary = data.summary;
    currentPendingId = null;
    renderResults(data.summary);
  } catch (err) {
    errorBox.textContent = '⚠ ' + err.message;
    errorBox.classList.remove('hidden');
  } finally {
    btn.disabled = false;
    procText.textContent = 'Process Files';
    loader.classList.add('hidden');
  }
}

async function parseResponseError(res) {
  const data = await res.json().catch(() => ({}));
  return data.detail || data.message || `Server error ${res.status}`;
}

async function importRegisterFile() {
  const btn      = document.getElementById('import-register-btn');
  const loader   = document.getElementById('import-loader');
  const procText = document.getElementById('import-register-text');
  const errorBox = document.getElementById('error-box');

  errorBox.classList.add('hidden');
  btn.disabled = true;
  procText.textContent = 'Loading…';
  loader.classList.remove('hidden');

  const fd = new FormData();
  fd.append('register_xlsx', registerFile);

  try {
    let res = await authFetch('/api/process', { method: 'POST', body: fd });
    if (res.status === 404) {
      res = await authFetch('/api/import-register', { method: 'POST', body: fd });
    }
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.detail || data.message || `Server error ${res.status}`);
    sessionId      = data.session_id;
    downloadUrl    = data.download_url;
    currentSummary = data.summary;
    currentPendingId = null;
    renderResults(data.summary);
    showToast('Register loaded from workbook', 'success');
  } catch (err) {
    errorBox.textContent = '⚠ ' + err.message;
    errorBox.classList.remove('hidden');
  } finally {
    btn.disabled = false;
    procText.textContent = 'Load Register';
    loader.classList.add('hidden');
  }
}

/* ══════════════════════════════════════════════════════════════════════════
   BUSINESS LOGIC  (mirrors Python processor.py)
   ══════════════════════════════════════════════════════════════════════════ */
function timeToMins(t) {
  if (!t) return null;
  const [h, m] = t.split(':').map(Number);
  return h * 60 + m;
}

function punchDurationMins(inTime, outTime) {
  const inM  = timeToMins(inTime);
  const outM = timeToMins(outTime);
  if (inM === null || outM === null) return 0;
  return outM >= inM ? outM - inM : (24 * 60 - inM) + outM;
}

function extraPunchMins(extraPunches) {
  return (extraPunches || []).reduce((sum, p) => (
    sum + punchDurationMins(p.in_time, p.out_time)
  ), 0);
}

function minsToStr(mins) {
  if (mins === null || mins === undefined) return '—';
  const sign = mins < 0 ? '-' : '';
  const abs  = Math.abs(Math.round(mins));
  return `${sign}${Math.floor(abs / 60)}:${String(abs % 60).padStart(2, '0')}`;
}

function getThresholds(location) {
  if (location.toLowerCase().includes('factory')) {
    return { sh_in:555, sh_out:1065, lh_in:510, lh_out:1125, lm_lo:600, lm_hi:810, hd_thr:810 };
  }
  return { sh_in:585, sh_out:1095, lh_in:555, lh_out:1155, lm_lo:630, lm_hi:840, hd_thr:840 };
}

function computeDayRow(empType, inTime, outTime, dayType, location, extraPunches = []) {
  const th   = getThresholds(location);
  const inM  = timeToMins(inTime);
  const outM = timeToMins(outTime);
  const extraMins = extraPunchMins(extraPunches);
  const hasExtra = extraMins > 0 || (extraPunches || []).some(p => p.in_time || p.out_time);
  const status = (inTime || outTime || hasExtra) ? 'Present' : 'Absent';

  if (empType === 'Labour') {
    const hw = (status === 'Present' && inM !== null && outM !== null) ? outM - inM : 0;
    const net = hw + extraMins;
    const ot = status === 'Present' ? net - 9 * 60 : 0;
    return { status, short_mins:0, long_mins:0, net_mins:net, overtime_mins:ot,
             late_mark:'', half_day:'', extra_mins:extraMins };
  }

  let halfDay = '';
  if (dayType === 'Working Day' && status === 'Present') {
    if (inM !== null && inM >= th.hd_thr)        halfDay = 'Half Day';
    else if (outM !== null && outM <= th.hd_thr)  halfDay = 'Half Day';
  }

  let shortMins = 0, longMins = 0, netMins = 0;
  if (!halfDay) {
    if (dayType === 'Working Day' && status === 'Present' && inM !== null && outM !== null) {
      shortMins = Math.max(0, inM - th.sh_in) + Math.max(0, th.sh_out - outM);
      longMins  = Math.max(0, th.lh_in - inM) + Math.max(0, outM - th.lh_out);
      netMins   = longMins - shortMins;
    } else if ((dayType === 'Company Holiday' || dayType === 'Week Off') &&
               status === 'Present' && inM !== null && outM !== null) {
      netMins = outM - inM;
    }
  }
  netMins += extraMins;

  let lateMark = '';
  if (dayType === 'Working Day' && status === 'Present' && inM !== null) {
    if (inM >= th.lm_lo && inM < th.lm_hi) lateMark = 'Late Mark';
  }

  return { status, short_mins:shortMins, long_mins:longMins, net_mins:netMins,
           overtime_mins:0, late_mark:lateMark, half_day:halfDay, extra_mins:extraMins };
}

function computeSummary(emp) {
  const daily = emp.daily;
  const wdPresent = daily.filter(d => d.day_type === 'Working Day' && d.status === 'Present');
  const wdAbsent  = daily.filter(d => d.day_type === 'Working Day' && d.status === 'Absent');
  const halfDays  = daily.filter(d => d.half_day === 'Half Day');
  const totalNetMins = daily.reduce((s, d) => s + d.net_mins, 0);

  if (emp.type === 'Labour') {
    const allPresent  = daily.filter(d => d.status === 'Present');
    const totalOtMins = daily.reduce((s, d) => s + d.overtime_mins, 0);
    return {
      present_days: allPresent.length,
      absent_days:  wdAbsent.length,
      half_days: 0,
      total_net_hours: totalNetMins / 60,
      working_days: daily.filter(d => d.day_type === 'Working Day').length,
      total_overtime_hours: totalOtMins / 60,
      ot_days: totalOtMins / (6 * 60),
    };
  }

  const el_o = emp.el_opening || 0;
  const cl_o = emp.cl_opening || 0;
  const sl_o = emp.sl_opening || 0;
  const sl_a = emp.sl_applied || 0;

  const netHours  = totalNetMins / 60;
  const elAccrual = netHours < 0 ? Math.ceil(netHours / 9) : Math.floor(netHours / 9);

  const presentDays = wdPresent.length;
  const absentDays  = wdAbsent.length;
  const numHalf     = halfDays.length;
  const totalAbsent = absentDays + 0.5 * numHalf;

  const clAdj     = Math.min(1, cl_o, Math.max(0, totalAbsent));
  const slAdj     = Math.min(sl_a, sl_o);
  const adjElOpen = Math.max(0, el_o + elAccrual);
  const elAdj     = Math.min(adjElOpen, Math.max(0, totalAbsent - clAdj - slAdj));
  const lwpNegEl  = (el_o + elAccrual) < 0 ? -(el_o + elAccrual) : 0;
  const lwpExcess = Math.max(0, totalAbsent - clAdj - slAdj - elAdj);

  return {
    present_days: presentDays, absent_days: absentDays, half_days: numHalf,
    total_net_hours: netHours,
    el_accrual: elAccrual, cl_adj: clAdj, sl_adj: slAdj, el_adj: elAdj,
    total_lwp: lwpNegEl + lwpExcess,
    cl_closing: Math.max(0, cl_o - clAdj),
    el_closing: Math.max(0, adjElOpen - elAdj),
    sl_closing: Math.max(0, sl_o - slAdj),
  };
}

/* ══════════════════════════════════════════════════════════════════════════
   RENDER RESULTS
   ══════════════════════════════════════════════════════════════════════════ */
function renderResults(summary) {
  document.getElementById('upload-section').classList.add('hidden');
  document.getElementById('results-section').classList.remove('hidden');
  document.getElementById('results-title').textContent = `Register Generated · ${summary.month}`;

  // Reset action button states
  hasUnsavedChanges = false;
  isFrozen          = false;
  const acceptBtn = document.getElementById('accept-btn');
  acceptBtn.classList.add('hidden');
  acceptBtn.classList.remove('pending', 'saved');
  document.getElementById('accept-btn-text').textContent = 'Accept Changes';
  _resetFreezeZone();
  document.getElementById('download-sub').textContent =
    'Complete workbook · all employees · all formulas live';
  applyRoleControls();

  // Store mutable copy
  window.empData = summary.employees.map(e => ({
    ...e,
    daily: e.daily.map(d => ({
      ...d,
      extra_punches: (d.extra_punches || []).map(p => ({ ...p })),
      extra_mins: d.extra_mins || 0,
    })),
  }));

  // Stats bar
  const employees = summary.employees || [];
  const fullTime  = employees.filter(e => e.type !== 'Labour').length;
  const labour    = employees.filter(e => e.type === 'Labour').length;
  document.getElementById('stats-bar').innerHTML = [
    { value: summary.total_emp, label: 'Total Workforce' },
    { value: fullTime,          label: 'Full-time' },
    { value: labour,            label: 'Labour' },
    { value: (summary.holidays || []).length, label: 'Holidays' },
  ].map(s => `
    <div class="stat-item">
      <div class="stat-value">${s.value}</div>
      <div class="stat-label">${s.label}</div>
    </div>`).join('');

  // Employee cards
  document.getElementById('emp-grid').innerHTML =
    window.empData.map((emp, i) => buildCardHTML(emp, i)).join('');

  document.getElementById('download-btn').onclick = downloadFile;
}

/* ══════════════════════════════════════════════════════════════════════════
   CARD BUILDER
   ══════════════════════════════════════════════════════════════════════════ */
function buildCardHTML(emp, i) {
  const isFactory   = emp.location && emp.location.toLowerCase().includes('factory');
  const isLabour    = emp.type === 'Labour';
  const avatarClass = isLabour ? 'labour' : (isFactory ? 'factory' : 'corporate');
  const badgeClass  = isLabour ? 'badge-labour' : (isFactory ? 'badge-factory' : 'badge-corporate');
  const initials    = emp.name.split(' ').map(w => w[0]).join('').substring(0, 2).toUpperCase();
  const summary     = computeSummary(emp);

  // Use Excel emp_id; fall back to PDF code if not set
  const displayId = emp.emp_id || emp.code || '';

  // Meta line: EmpId · designation · department · area · week-off day
  const metaParts = [
    displayId         ? displayId : null,
    emp.designation   ? emp.designation : null,
    emp.department    ? emp.department  : null,
    emp.area          ? emp.area        : null,
    emp.week_off !== undefined ? `WO:${DAY_NAMES[emp.week_off]}` : null,
  ].filter(Boolean);

  const metaHtml = metaParts.length
    ? `<div class="emp-meta">${metaParts.map((p, idx) =>
        idx === 0 ? p : `<span class="emp-meta-sep">·</span>${p}`).join('')}</div>`
    : '';

  // Quick stats (with IDs for dynamic updates)
  const qsHtml = `
    <div class="emp-quick-stats">
      <div class="emp-qs-item">
        <span class="emp-qs-val green" id="qs-present-${i}">${summary.present_days}</span>
        <span class="emp-qs-label">Present</span>
      </div>
      <div class="emp-qs-item">
        <span class="emp-qs-val red" id="qs-absent-${i}">${summary.absent_days}</span>
        <span class="emp-qs-label">Absent</span>
      </div>
      <div class="emp-qs-item" id="qs-half-item-${i}" style="${(!isLabour && summary.half_days > 0) ? '' : 'display:none'}">
        <span class="emp-qs-val purple" id="qs-half-${i}">${summary.half_days}</span>
        <span class="emp-qs-label">Half Days</span>
      </div>
    </div>`;

  // Leave pills
  const pillsHtml = !isLabour ? `
    <div class="emp-leave-pills" id="pills-${i}">${buildLeavePills(summary)}</div>` : `
    <div class="emp-leave-pills" id="pills-${i}">
      <div class="leave-pill">
        <span class="pill-label">OT</span>
        <span class="pill-val">${minsToStr(summary.total_overtime_hours * 60)}</span>
      </div>
    </div>`;

  return `
    <div class="emp-card" id="card-${i}" style="animation-delay:${i * 0.04}s">
      <div class="emp-card-header" onclick="toggleCard(${i})">
        <div class="emp-avatar ${avatarClass}">${initials}</div>
        <div class="emp-identity">
          <div class="emp-name">${emp.name}</div>
          ${metaHtml}
        </div>
        <div class="emp-badges">
          <span class="badge ${badgeClass}">${emp.location || 'Factory'}</span>
          <span class="badge badge-fulltime">${emp.type}</span>
        </div>
        ${qsHtml}
        ${pillsHtml}
        <div class="emp-toggle">
          <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
            <path d="M3 6l5 5 5-5" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
          </svg>
        </div>
      </div>
      <div class="emp-body" id="body-${i}">
        <div class="emp-body-inner">
          ${buildInfoBar(emp)}
          ${buildTableHTML(emp, i)}
          <div class="emp-summaries" id="summaries-${i}">
            ${buildSummariesHTML(emp, summary, i)}
          </div>
        </div>
      </div>
    </div>`;
}

/* ── Employee info bar (inside expanded card) ────────────────────────────── */
function buildInfoBar(emp) {
  const displayId = emp.emp_id || emp.code || '';
  const fields = [
    displayId       ? { label: 'Emp ID',      val: displayId }        : null,
    emp.designation ? { label: 'Designation', val: emp.designation }  : null,
    emp.department  ? { label: 'Department',  val: emp.department }   : null,
    emp.area        ? { label: 'Area',         val: emp.area }        : null,
    { label: 'Location', val: emp.location || '—' },
    { label: 'Week Off', val: DAY_NAMES[emp.week_off] || '—' },
  ].filter(Boolean);

  if (!fields.length) return '';
  return `<div class="emp-info-bar">${fields.map(f => `
    <div class="emp-info-item">
      <span class="emp-info-label">${f.label}</span>
      <span class="emp-info-val">${f.val}</span>
    </div>`).join('')}</div>`;
}

/* ══════════════════════════════════════════════════════════════════════════
   LEAVE PILLS / TABLES
   ══════════════════════════════════════════════════════════════════════════ */
function fmtLeave(v) {
  if (v === null || v === undefined) return '—';
  if (typeof v !== 'number') return v;
  return parseFloat(v.toFixed(2)).toString();
}

function buildLeavePills(summary) {
  const pills = [
    { label: 'EL', val: summary.el_closing ?? '—', lwp: false },
    { label: 'CL', val: summary.cl_closing ?? '—', lwp: false },
    { label: 'SL', val: summary.sl_closing ?? '—', lwp: false },
  ];
  if (summary.total_lwp > 0) pills.push({ label: 'LWP', val: summary.total_lwp, lwp: true });
  return pills.map(p => `
    <div class="leave-pill${p.lwp ? ' lwp' : ''}">
      <span class="pill-label">${p.label}</span>
      <span class="pill-val">${fmtLeave(p.val)}</span>
    </div>`).join('');
}

function buildLeaveTableRows(emp, summary) {
  const fmt = v => (v === undefined || v === null) ? '—' : fmtLeave(v);
  const rows = [
    { label:'EL', opening:emp.el_opening, adj:summary.el_accrual,  closing:summary.el_closing },
    { label:'CL', opening:emp.cl_opening, adj:-summary.cl_adj,     closing:summary.cl_closing },
    { label:'SL', opening:emp.sl_opening, adj:-summary.sl_adj,     closing:summary.sl_closing },
  ];
  const leaveRows = rows.map(r => `
    <tr>
      <td>${r.label}</td>
      <td>${fmt(r.opening)}</td>
      <td class="${r.adj > 0 ? 'cell-net-pos' : r.adj < 0 ? 'cell-net-neg' : ''}">${r.adj > 0 ? '+' : ''}${fmt(r.adj)}</td>
      <td class="closing">${fmt(r.closing)}</td>
    </tr>`).join('');
  const lwpRow = summary.total_lwp > 0 ? `
    <tr class="lwp-row">
      <td>LWP</td><td>—</td><td>—</td><td class="closing">${fmt(summary.total_lwp)}</td>
    </tr>` : '';
  return leaveRows + lwpRow;
}

/* ══════════════════════════════════════════════════════════════════════════
   TABLE BUILDER
   ══════════════════════════════════════════════════════════════════════════ */
function buildTableHTML(emp, i) {
  const isLabour = emp.type === 'Labour';
  const headers  = isLabour
    ? ['Date','Day Type','Status','In','Out','Extra In/Out','Extra Hrs','Hrs Worked','OT','Note']
    : ['Date','Day Type','Status','In','Out','Extra In/Out','Extra Hrs','Short','Long','Net','Late','Half','Note'];

  const rows = emp.daily.map((d, di) => buildRowHTML(emp, i, d, di, isLabour)).join('');
  return `
    <div class="att-table-wrap">
      <table class="att-table">
        <thead><tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr></thead>
        <tbody id="tbody-${i}">${rows}</tbody>
      </table>
    </div>`;
}

function isIncomplete(d) {
  return (d.in_time && !d.out_time) || (!d.in_time && d.out_time) ||
    (d.extra_punches || []).some(p => (p.in_time && !p.out_time) || (!p.in_time && p.out_time));
}

function noteHTML(d) {
  if (d.in_time && !d.out_time)  return '<span class="note-warn">⚠ Out-Time missing</span>';
  if (!d.in_time && d.out_time)  return '<span class="note-warn">⚠ In-Time missing</span>';
  if ((d.extra_punches || []).some(p => p.in_time && !p.out_time)) return '<span class="note-warn">⚠ Extra Out missing</span>';
  if ((d.extra_punches || []).some(p => !p.in_time && p.out_time)) return '<span class="note-warn">⚠ Extra In missing</span>';
  return '';
}

function rowClass(d) {
  if (isIncomplete(d))                  return 'row-incomplete';
  if (d.day_type === 'Week Off')        return 'row-weekoff';
  if (d.day_type === 'Company Holiday') return 'row-holiday';
  if (d.status === 'Present')           return 'row-present';
  return 'row-absent';
}

function statusClass(status, dayType) {
  if (dayType === 'Week Off')        return 'cell-weekoff';
  if (dayType === 'Company Holiday') return 'cell-holiday';
  if (status === 'Present')          return 'cell-present';
  return 'cell-absent';
}

function dayTypeLabel(dt) {
  if (dt === 'Week Off')        return '<span class="cell-weekoff">Week Off</span>';
  if (dt === 'Company Holiday') return '<span class="cell-holiday">Holiday</span>';
  return 'Working';
}

function netClass(mins) {
  if (mins > 0) return 'cell-net-pos';
  if (mins < 0) return 'cell-net-neg';
  return '';
}

function timeCell(time, empIdx, dayIdx, field) {
  // ALL time cells are editable inputs — whether or not data already exists.
  // Pre-existing values are shown with class "has-value" (green tint) so the
  // user can distinguish filled vs empty slots at a glance.
  const val = (time !== null && time !== undefined && time !== '') ? time : '';
  const cls = val ? 'time-edit has-value' : 'time-edit';
  return `<input type="time" class="${cls}"
            data-emp="${empIdx}" data-day="${dayIdx}" data-field="${field}"
            value="${val}"
            onchange="handleTimeEdit(this)">`;
}

function extraPunchesHTML(d, empIdx, dayIdx) {
  const punches = d.extra_punches || [];
  const pairs = punches.map((p, pi) => `
    <div class="extra-punch-pair">
      ${extraTimeCell(p.in_time, empIdx, dayIdx, pi, 'in_time')}
      ${extraTimeCell(p.out_time, empIdx, dayIdx, pi, 'out_time')}
      <button type="button" class="extra-remove-btn"
              title="Remove extra punch"
              onclick="removeExtraPunch(${empIdx}, ${dayIdx}, ${pi})">×</button>
    </div>`).join('');
  return `
    <div class="extra-punches" id="extra-punches-${empIdx}-${dayIdx}">
      ${pairs}
      <button type="button" class="extra-add-btn"
              onclick="addExtraPunch(${empIdx}, ${dayIdx})">+ Add</button>
    </div>`;
}

function extraTimeCell(time, empIdx, dayIdx, punchIdx, field) {
  const val = (time !== null && time !== undefined && time !== '') ? time : '';
  const cls = val ? 'time-edit extra-time has-value' : 'time-edit extra-time';
  return `<input type="time" class="${cls}"
            data-emp="${empIdx}" data-day="${dayIdx}"
            data-punch="${punchIdx}" data-field="${field}"
            value="${val}"
            onchange="handleExtraTimeEdit(this)">`;
}

function buildRowHTML(emp, i, d, di, isLabour) {
  const rc = rowClass(d);
  const sc = statusClass(d.status, d.day_type);
  const inCell  = timeCell(d.in_time,  i, di, 'in_time');
  const outCell = timeCell(d.out_time, i, di, 'out_time');
  const extraCells = extraPunchesHTML(d, i, di);
  const extraMins = extraPunchMins(d.extra_punches);

  const dateCol = `<td><div class="date-cell"><span class="date-day">${d.date}</span><span class="date-dow">${d.dow}</span></div></td>`;

  if (isLabour) {
    return `<tr class="${rc}" id="row-${i}-${di}">
      ${dateCol}
      <td>${dayTypeLabel(d.day_type)}</td>
      <td id="c-status-${i}-${di}" class="${sc}">${d.status}</td>
      <td id="c-in-${i}-${di}">${inCell}</td>
      <td id="c-out-${i}-${di}">${outCell}</td>
      <td id="c-extra-punches-${i}-${di}">${extraCells}</td>
      <td id="c-extra-${i}-${di}" class="${extraMins > 0 ? 'cell-net-pos' : ''}">${extraMins ? minsToStr(extraMins) : '—'}</td>
      <td id="c-hw-${i}-${di}">${d.status === 'Present' ? minsToStr(d.net_mins) : '—'}</td>
      <td id="c-ot-${i}-${di}" class="${d.overtime_mins > 0 ? 'cell-net-pos' : ''}">${d.status === 'Present' ? minsToStr(d.overtime_mins) : '—'}</td>
      <td id="c-note-${i}-${di}" class="note-cell">${noteHTML(d)}</td>
    </tr>`;
  }

  return `<tr class="${rc}" id="row-${i}-${di}">
    ${dateCol}
    <td>${dayTypeLabel(d.day_type)}</td>
    <td id="c-status-${i}-${di}" class="${sc}">${d.status}</td>
    <td id="c-in-${i}-${di}">${inCell}</td>
    <td id="c-out-${i}-${di}">${outCell}</td>
    <td id="c-extra-punches-${i}-${di}">${extraCells}</td>
    <td id="c-extra-${i}-${di}" class="${extraMins > 0 ? 'cell-net-pos' : ''}">${extraMins ? minsToStr(extraMins) : '—'}</td>
    <td id="c-short-${i}-${di}">${minsToStr(d.short_mins)}</td>
    <td id="c-long-${i}-${di}">${minsToStr(d.long_mins)}</td>
    <td id="c-net-${i}-${di}" class="${netClass(d.net_mins)}">${minsToStr(d.net_mins)}</td>
    <td id="c-late-${i}-${di}" class="cell-late">${d.late_mark ? '★' : ''}</td>
    <td id="c-half-${i}-${di}" class="cell-half">${d.half_day ? '½' : ''}</td>
    <td id="c-note-${i}-${di}" class="note-cell">${noteHTML(d)}</td>
  </tr>`;
}

/* ══════════════════════════════════════════════════════════════════════════
   SUMMARY PANELS
   ══════════════════════════════════════════════════════════════════════════ */
function buildSummariesHTML(emp, summary, i) {
  const isLabour = emp.type === 'Labour';

  const monthlySummaryRows = isLabour ? `
    <div class="summary-row"><span class="summary-row-label">Working Days</span><span class="summary-row-val" id="s-wd-${i}">${summary.working_days}</span></div>
    <div class="summary-row"><span class="summary-row-label">Present Days</span><span class="summary-row-val green" id="s-present-${i}">${summary.present_days}</span></div>
    <div class="summary-row"><span class="summary-row-label">Cumulative OT</span><span class="summary-row-val accent" id="s-ot-${i}">${minsToStr(summary.total_overtime_hours * 60)}</span></div>
    <div class="summary-row"><span class="summary-row-label">OT Days</span><span class="summary-row-val" id="s-otd-${i}">${summary.ot_days.toFixed(2)}</span></div>
  ` : `
    <div class="summary-row"><span class="summary-row-label">Present Days</span><span class="summary-row-val green" id="s-present-${i}">${summary.present_days}</span></div>
    <div class="summary-row"><span class="summary-row-label">Absent Days</span><span class="summary-row-val red" id="s-absent-${i}">${summary.absent_days}</span></div>
    <div class="summary-row"><span class="summary-row-label">Half Days</span><span class="summary-row-val purple" id="s-half-${i}">${summary.half_days}</span></div>
    <div class="summary-row"><span class="summary-row-label">Net Hours</span><span class="summary-row-val accent" id="s-net-${i}">${minsToStr(summary.total_net_hours * 60)}</span></div>
  `;

  const leavePanelHtml = !isLabour ? `
    <div class="summary-panel">
      <div class="summary-panel-title">Leave Summary</div>
      <table class="leave-table" id="leave-table-${i}">
        <thead><tr><th>Type</th><th>Opening</th><th>Adj</th><th>Closing</th></tr></thead>
        <tbody>${buildLeaveTableRows(emp, summary)}</tbody>
      </table>
    </div>` : '';

  return `
    <div class="summary-panel">
      <div class="summary-panel-title">Monthly Summary</div>
      <div class="summary-rows" id="monthly-summary-${i}">${monthlySummaryRows}</div>
    </div>
    ${leavePanelHtml}`;
}

/* ══════════════════════════════════════════════════════════════════════════
   DYNAMIC RECALCULATION  (Feature 2)
   ══════════════════════════════════════════════════════════════════════════ */
function handleTimeEdit(input) {
  const empIdx = parseInt(input.dataset.emp);
  const dayIdx = parseInt(input.dataset.day);
  const field  = input.dataset.field;
  const value  = input.value || null;

  const emp = window.empData[empIdx];
  const row = emp.daily[dayIdx];

  // 1. Keep visual class in sync (green = has value, amber = empty)
  if (value) {
    input.classList.add('has-value');
  } else {
    input.classList.remove('has-value');
  }

  // 2. Update stored data
  row[field] = value;

  // 3. Recompute all day metrics
  const metrics = computeDayRow(emp.type, row.in_time, row.out_time, row.day_type,
                                emp.location, row.extra_punches);
  Object.assign(row, metrics);

  // 4. Update every computed cell in this row
  updateRowCells(empIdx, dayIdx);

  // 5. Recompute monthly + leave summary
  const summary = computeSummary(emp);

  // 6. Write summary back into window.empData so freeze captures correct closing
  //    balances (el_closing, cl_closing, sl_closing, total_lwp, present_days, etc.)
  Object.assign(window.empData[empIdx], summary);

  // 7. Push updated summary to DOM
  updateSummaryDOM(empIdx, emp, summary);

  // 8. Flag unsaved changes
  markUnsaved();
}

function handleExtraTimeEdit(input) {
  const empIdx = parseInt(input.dataset.emp);
  const dayIdx = parseInt(input.dataset.day);
  const punchIdx = parseInt(input.dataset.punch);
  const field = input.dataset.field;
  const value = input.value || null;

  const emp = window.empData[empIdx];
  const row = emp.daily[dayIdx];
  row.extra_punches = row.extra_punches || [];
  row.extra_punches[punchIdx] = row.extra_punches[punchIdx] || { in_time:null, out_time:null };
  row.extra_punches[punchIdx][field] = value;

  if (value) input.classList.add('has-value');
  else input.classList.remove('has-value');

  recalcEditedDay(empIdx, dayIdx);
}

function addExtraPunch(empIdx, dayIdx) {
  const row = window.empData[empIdx].daily[dayIdx];
  row.extra_punches = row.extra_punches || [];
  row.extra_punches.push({ in_time:null, out_time:null });
  updateExtraPunchInputs(empIdx, dayIdx);
  recalcEditedDay(empIdx, dayIdx);
}

function removeExtraPunch(empIdx, dayIdx, punchIdx) {
  const row = window.empData[empIdx].daily[dayIdx];
  row.extra_punches = row.extra_punches || [];
  row.extra_punches.splice(punchIdx, 1);
  updateExtraPunchInputs(empIdx, dayIdx);
  recalcEditedDay(empIdx, dayIdx);
}

function recalcEditedDay(empIdx, dayIdx) {
  const emp = window.empData[empIdx];
  const row = emp.daily[dayIdx];
  const metrics = computeDayRow(emp.type, row.in_time, row.out_time, row.day_type,
                                emp.location, row.extra_punches);
  Object.assign(row, metrics);
  updateRowCells(empIdx, dayIdx);
  const summary = computeSummary(emp);
  Object.assign(window.empData[empIdx], summary);
  updateSummaryDOM(empIdx, emp, summary);
  markUnsaved();
}

function updateExtraPunchInputs(empIdx, dayIdx) {
  const d = window.empData[empIdx].daily[dayIdx];
  const el = document.getElementById(`c-extra-punches-${empIdx}-${dayIdx}`);
  if (el) el.innerHTML = extraPunchesHTML(d, empIdx, dayIdx);
}

function updateRowCells(i, di) {
  const d        = window.empData[i].daily[di];
  const isLabour = window.empData[i].type === 'Labour';
  const sc       = statusClass(d.status, d.day_type);

  const statusEl = document.getElementById(`c-status-${i}-${di}`);
  if (statusEl) { statusEl.textContent = d.status; statusEl.className = sc; }

  const rowEl = document.getElementById(`row-${i}-${di}`);
  if (rowEl) rowEl.className = rowClass(d);

  if (isLabour) {
    const extraEl = document.getElementById(`c-extra-${i}-${di}`);
    if (extraEl) {
      extraEl.textContent = d.extra_mins ? minsToStr(d.extra_mins) : '—';
      extraEl.className = d.extra_mins > 0 ? 'cell-net-pos' : '';
    }
    setText(`c-hw-${i}-${di}`, d.status === 'Present' ? minsToStr(d.net_mins) : '—');
    const otEl = document.getElementById(`c-ot-${i}-${di}`);
    if (otEl) {
      otEl.textContent = d.status === 'Present' ? minsToStr(d.overtime_mins) : '—';
      otEl.className   = d.overtime_mins > 0 ? 'cell-net-pos' : '';
    }
  } else {
    const extraEl = document.getElementById(`c-extra-${i}-${di}`);
    if (extraEl) {
      extraEl.textContent = d.extra_mins ? minsToStr(d.extra_mins) : '—';
      extraEl.className = d.extra_mins > 0 ? 'cell-net-pos' : '';
    }
    setText(`c-short-${i}-${di}`, minsToStr(d.short_mins));
    setText(`c-long-${i}-${di}`,  minsToStr(d.long_mins));
    const netEl = document.getElementById(`c-net-${i}-${di}`);
    if (netEl) { netEl.textContent = minsToStr(d.net_mins); netEl.className = netClass(d.net_mins); }
    const lateEl = document.getElementById(`c-late-${i}-${di}`);
    if (lateEl) lateEl.textContent = d.late_mark ? '★' : '';
    const halfEl = document.getElementById(`c-half-${i}-${di}`);
    if (halfEl) halfEl.textContent = d.half_day ? '½' : '';
  }

  const noteEl = document.getElementById(`c-note-${i}-${di}`);
  if (noteEl) noteEl.innerHTML = noteHTML(d);
}

function updateSummaryDOM(i, emp, summary) {
  const isLabour = emp.type === 'Labour';

  // ── Collapsed-mode quick stats (always in DOM inside card header) ──────────
  // Update by direct element reference to guarantee refresh regardless of card state
  const qsPresent = document.getElementById(`qs-present-${i}`);
  const qsAbsent  = document.getElementById(`qs-absent-${i}`);
  if (qsPresent) qsPresent.textContent = summary.present_days;
  if (qsAbsent)  qsAbsent.textContent  = summary.absent_days;

  if (!isLabour) {
    const qsHalf     = document.getElementById(`qs-half-${i}`);
    const qsHalfItem = document.getElementById(`qs-half-item-${i}`);
    if (qsHalf)     qsHalf.textContent        = summary.half_days;
    if (qsHalfItem) qsHalfItem.style.display  = summary.half_days > 0 ? '' : 'none';
  }

  // Detailed summary rows
  if (isLabour) {
    setText(`s-present-${i}`, summary.present_days);
    setText(`s-ot-${i}`,      minsToStr(summary.total_overtime_hours * 60));
    setText(`s-otd-${i}`,     summary.ot_days.toFixed(2));
  } else {
    setText(`s-present-${i}`, summary.present_days);
    setText(`s-absent-${i}`,  summary.absent_days);
    setText(`s-half-${i}`,    summary.half_days);
    setText(`s-net-${i}`,     minsToStr(summary.total_net_hours * 60));

    const lt = document.getElementById(`leave-table-${i}`);
    if (lt) {
      const tbody = lt.querySelector('tbody');
      if (tbody) tbody.innerHTML = buildLeaveTableRows(emp, summary);
    }
  }

  // Leave pills
  const pillsEl = document.getElementById(`pills-${i}`);
  if (pillsEl) {
    if (isLabour) {
      pillsEl.innerHTML = `
        <div class="leave-pill">
          <span class="pill-label">OT</span>
          <span class="pill-val">${minsToStr(summary.total_overtime_hours * 60)}</span>
        </div>`;
    } else {
      pillsEl.innerHTML = buildLeavePills(summary);
    }
  }
}

function setText(id, text) {
  const el = document.getElementById(id);
  if (el) el.textContent = text;
}

function toggleCard(i) {
  document.getElementById(`card-${i}`).classList.toggle('expanded');
}

/* ══════════════════════════════════════════════════════════════════════════
   UNSAVED-CHANGE STATE
   ══════════════════════════════════════════════════════════════════════════ */
function markUnsaved() {
  hasUnsavedChanges = true;
  const btn = document.getElementById('accept-btn');
  btn.classList.remove('hidden', 'saved');
  btn.classList.add('pending');
  document.getElementById('accept-btn-text').textContent = 'Accept Changes';
  document.getElementById('download-sub').textContent =
    'Unsaved edits — accept changes before downloading';
}

function markSaved() {
  hasUnsavedChanges = false;
  const btn = document.getElementById('accept-btn');
  btn.classList.remove('pending');
  btn.classList.add('saved');
  document.getElementById('accept-btn-text').textContent = 'Saved ✓';
  document.getElementById('download-sub').textContent =
    'Changes accepted · ready to freeze or download';
}

/* ══════════════════════════════════════════════════════════════════════════
   ACCEPT CHANGES
   ══════════════════════════════════════════════════════════════════════════ */
async function acceptChanges() {
  const btn     = document.getElementById('accept-btn');
  const loader  = document.getElementById('accept-loader');
  const btnText = document.getElementById('accept-btn-text');

  btn.disabled = true;
  btnText.textContent = 'Saving…';
  loader.classList.remove('hidden');

  try {
    const res = await authFetch(`/api/rebuild/${sessionId}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        employees: window.empData,
        year:      currentSummary.year,
        month_num: currentSummary.month_num,
        holidays:  currentSummary.holidays,
      }),
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.detail || `Server error ${res.status}`);
    }
    lockFilledInputs();
    markSaved();
    showToast('Changes accepted and recorded', 'success');
  } catch (err) {
    showToast('Failed to save: ' + err.message, 'error');
    btn.classList.add('pending');
    btnText.textContent = 'Accept Changes';
  } finally {
    btn.disabled = false;
    loader.classList.add('hidden');
  }
}

function lockFilledInputs() {
  // Inputs stay editable — just refresh the has-value class after saving
  // so any newly-filled cells switch from amber to green styling.
  document.querySelectorAll('input.time-edit').forEach(input => {
    if (input.value) {
      input.classList.add('has-value');
    } else {
      input.classList.remove('has-value');
    }
  });
}

async function submitForApproval() {
  const btn = document.getElementById('submit-btn');
  const loader = document.getElementById('submit-loader');
  const btnText = document.getElementById('submit-btn-text');

  btn.disabled = true;
  btnText.textContent = 'Submitting…';
  loader.classList.remove('hidden');

  try {
    const res = await authFetch(`/api/submit/${sessionId}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        employees: window.empData,
        year:      currentSummary.year,
        month_num: currentSummary.month_num,
        month:     currentSummary.month,
        holidays:  currentSummary.holidays,
      }),
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.detail || `Server error ${res.status}`);
    }
    hasUnsavedChanges = false;
    lockFilledInputs();
    markSaved();
    btnText.textContent = 'Submitted';
    showToast('Register submitted for admin approval', 'success');
  } catch (err) {
    btn.disabled = false;
    btnText.textContent = 'Submit for Approval';
    showToast('Submit failed: ' + err.message, 'error');
  } finally {
    loader.classList.add('hidden');
  }
}

/* ══════════════════════════════════════════════════════════════════════════
   FREEZE REGISTER  (Feature 3)
   ══════════════════════════════════════════════════════════════════════════ */
function openFreezeModal() {
  const warnEl = document.getElementById('freeze-modal-warn');

  // Build summary for modal
  const summary = document.getElementById('freeze-modal-summary');
  summary.innerHTML = [
    { label: 'Period',    val: currentSummary.month },
    { label: 'Employees', val: currentSummary.total_emp },
    { label: 'Holidays',  val: (currentSummary.holidays || []).length },
  ].map(r => `
    <div class="modal-summary-row">
      <span>${r.label}</span>
      <span>${r.val}</span>
    </div>`).join('');

  if (hasUnsavedChanges) {
    warnEl.textContent = '⚠ You have unsaved edits. Freezing will first accept all changes.';
    warnEl.classList.remove('hidden');
  } else {
    warnEl.classList.add('hidden');
  }

  document.getElementById('freeze-modal').classList.remove('hidden');
}

function closeFreezeModal() {
  document.getElementById('freeze-modal').classList.add('hidden');
}

async function confirmFreeze() {
  const confirmBtn = document.getElementById('freeze-confirm-btn');
  const loader     = document.getElementById('freeze-loader');
  const btnText    = document.getElementById('freeze-confirm-text');

  confirmBtn.disabled = true;
  btnText.textContent = 'Freezing…';
  loader.classList.remove('hidden');

  try {
    if (currentPendingId) {
      await approvePending(currentPendingId);
      isFrozen = true;
      _setFreezeZoneFrozen(currentSummary.month);
      closeFreezeModal();
      return;
    }

    const res = await authFetch(`/api/freeze/${sessionId}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        employees: window.empData,
        year:      currentSummary.year,
        month_num: currentSummary.month_num,
        month:     currentSummary.month,
        holidays:  currentSummary.holidays,
      }),
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.detail || `Server error ${res.status}`);
    }

    // If there were unsaved changes, accept them in the UI now
    if (hasUnsavedChanges) {
      lockFilledInputs();
      markSaved();
    }

    // Mark register as frozen
    isFrozen = true;
    _setFreezeZoneFrozen(currentSummary.month);

    closeFreezeModal();
    showToast(`Register frozen for ${currentSummary.month}`, 'success');

    // Refresh previous registers list
    loadPreviousRegisters(false);
  } catch (err) {
    showToast('Freeze failed: ' + err.message, 'error');
  } finally {
    confirmBtn.disabled = false;
    btnText.textContent = 'Freeze Register';
    loader.classList.add('hidden');
  }
}

/* ══════════════════════════════════════════════════════════════════════════
   DOWNLOAD
   ══════════════════════════════════════════════════════════════════════════ */
async function downloadFile() {
  if (hasUnsavedChanges) {
    showToast('Accept your changes first before downloading', 'error');
    document.getElementById('accept-btn').classList.remove('hidden');
    return;
  }
  window.location.href = downloadUrl;
}

/* ══════════════════════════════════════════════════════════════════════════
   PREVIOUS REGISTERS  (Feature 3)
   ══════════════════════════════════════════════════════════════════════════ */
async function loadPreviousRegisters(showSection = true) {
  try {
    const res  = await authFetch('/api/registers');
    const list = await res.json();
    renderPreviousRegisters(list, showSection);
  } catch (_) {
    // Silently fail on load — registers are optional
  }
}

function renderPreviousRegisters(list, showSection = true) {
  const section = document.getElementById('prev-section');
  const grid    = document.getElementById('prev-registers-list');

  if (!list.length) {
    section.classList.add('hidden');
    return;
  }

  grid.innerHTML = list.map(r => {
    const frozenDate = new Date(r.frozen_at).toLocaleDateString('en-IN', {
      day:'2-digit', month:'short', year:'numeric', hour:'2-digit', minute:'2-digit'
    });
    return `
      <div class="prev-card" id="prev-card-${r.id}">
        <div class="prev-card-header">
          <div>
            <div class="prev-card-month">${r.month}</div>
            <div class="prev-card-date">Frozen: ${frozenDate}</div>
          </div>
          <div class="prev-card-id">#${r.id}</div>
        </div>
        <div class="prev-card-stats">
          <div><span>${r.total_emp}</span> employees</div>
        </div>
        <div class="prev-card-actions">
          <button class="prev-view-btn" onclick="togglePrevDetail(${r.id})">View Details</button>
          <button class="prev-dl-btn"   onclick="downloadFrozen(${r.id})">Download</button>
        </div>
        <div id="prev-detail-${r.id}" class="hidden"></div>
      </div>`;
  }).join('');

  if (showSection) section.classList.remove('hidden');
}

async function togglePrevDetail(id) {
  const detailEl = document.getElementById(`prev-detail-${id}`);
  const card     = document.getElementById(`prev-card-${id}`);

  if (!detailEl.classList.contains('hidden')) {
    detailEl.classList.add('hidden');
    card.classList.remove('detail-open');
    return;
  }

  detailEl.innerHTML = '<div style="color:var(--text3);font-size:12px;padding:8px 0">Loading…</div>';
  detailEl.classList.remove('hidden');
  card.classList.add('detail-open');

  try {
    const res  = await authFetch(`/api/registers/${id}`);
    const data = await res.json();
    detailEl.innerHTML = buildPrevDetailTable(data.employees || []);
  } catch (err) {
    detailEl.innerHTML = `<div style="color:var(--red);font-size:12px">${err.message}</div>`;
  }
}

function buildPrevDetailTable(employees) {
  if (!employees.length) return '<p style="color:var(--text3);font-size:12px;padding:8px 0">No employee data.</p>';

  /* ── Formatters ──────────────────────────────────────────────────────── */
  const fmtN = v => (v === null || v === undefined) ? '—'
                  : typeof v === 'number' ? parseFloat(v.toFixed(2)) : v;

  // Format decimal hours → H:MM  (e.g. 2.75 → "2:45")
  const fmtHrs = h => {
    if (h === null || h === undefined) return '—';
    const sign = h < 0 ? '-' : '';
    const abs  = Math.abs(h);
    const hh   = Math.floor(abs);
    const mm   = String(Math.round((abs - hh) * 60)).padStart(2, '0');
    return `${sign}${hh}:${mm}`;
  };

  // Colour-coded cell helper
  const td = (val, colour) =>
    `<td${colour ? ` style="color:var(${colour})"` : ''}>${val}</td>`;

  /* ── Split by type ───────────────────────────────────────────────────── */
  const fullTime = employees.filter(e => e.type !== 'Labour');
  const labour   = employees.filter(e => e.type === 'Labour');
  let html = '';

  /* ── Full-time table ─────────────────────────────────────────────────── */
  if (fullTime.length) {
    const rows = fullTime.map(e => `
      <tr>
        ${td(e.name)}
        ${td(e.emp_id  || '—')}
        ${td(e.designation || '—')}
        ${td(e.department  || '—')}
        ${td(fmtN(e.present_days), '--green')}
        ${td(fmtN(e.absent_days),  '--red')}
        ${td(fmtN(e.half_days) === '0' || !e.half_days ? '—' : fmtN(e.half_days), '--purple')}
        ${td(fmtHrs(e.total_net_hours), e.total_net_hours < 0 ? '--red' : '--text2')}
        ${td(fmtN(e.el_closing), '--blue')}
        ${td(fmtN(e.cl_closing), '--green')}
        ${td(fmtN(e.sl_closing), '--blue')}
        ${td(e.total_lwp > 0 ? fmtN(e.total_lwp) : '—', e.total_lwp > 0 ? '--red' : null)}
      </tr>`).join('');

    html += `
      <div class="prev-detail-section">
        <div class="prev-detail-label">Full-time Employees · ${fullTime.length}</div>
        <div class="prev-emp-table-wrap">
          <table class="prev-emp-table">
            <thead><tr>
              <th>Name</th>
              <th>Emp ID</th>
              <th>Designation</th>
              <th>Dept</th>
              <th>Present</th>
              <th>Absent</th>
              <th>Half</th>
              <th>Net Hrs</th>
              <th>EL Closing</th>
              <th>CL Closing</th>
              <th>SL Closing</th>
              <th>LWP</th>
            </tr></thead>
            <tbody>${rows}</tbody>
          </table>
        </div>
      </div>`;
  }

  /* ── Labour table ────────────────────────────────────────────────────── */
  if (labour.length) {
    const rows = labour.map(e => {
      const otHrs  = e.total_overtime_hours || 0;
      const otDays = e.ot_days || 0;
      return `
      <tr>
        ${td(e.name)}
        ${td(e.emp_id || '—')}
        ${td(fmtN(e.present_days), '--green')}
        ${td(fmtN(e.absent_days),  '--red')}
        ${td(fmtHrs(otHrs), otHrs > 0 ? '--accent' : '--text3')}
        ${td(otDays > 0 ? fmtN(otDays) : '—', otDays > 0 ? '--accent' : null)}
      </tr>`;
    }).join('');

    html += `
      <div class="prev-detail-section">
        <div class="prev-detail-label">Labour · ${labour.length}</div>
        <div class="prev-emp-table-wrap">
          <table class="prev-emp-table">
            <thead><tr>
              <th>Name</th>
              <th>Emp ID</th>
              <th>Present</th>
              <th>Absent</th>
              <th>OT Hours</th>
              <th>OT Days</th>
            </tr></thead>
            <tbody>${rows}</tbody>
          </table>
        </div>
      </div>`;
  }

  return html;
}

function downloadFrozen(id) {
  window.location.href = `/api/registers/${id}/download`;
}

async function loadPendingRegisters(showSection = true) {
  if (!isAdmin()) return;
  try {
    const res = await authFetch('/api/pending');
    const list = await res.json();
    renderPendingRegisters(list, showSection);
  } catch (err) {
    showToast('Failed to load pending approvals: ' + err.message, 'error');
  }
}

function renderPendingRegisters(list, showSection = true) {
  const section = document.getElementById('pending-section');
  const grid = document.getElementById('pending-registers-list');
  if (!section || !grid) return;

  if (!list.length) {
    grid.innerHTML = '<p style="color:var(--text3);font-size:13px">No pending approvals.</p>';
    if (showSection) section.classList.remove('hidden');
    return;
  }

  grid.innerHTML = list.map(r => {
    const submittedDate = new Date(r.submitted_at).toLocaleDateString('en-IN', {
      day:'2-digit', month:'short', year:'numeric', hour:'2-digit', minute:'2-digit'
    });
    return `
      <div class="prev-card" id="pending-card-${r.id}">
        <div class="prev-card-header">
          <div>
            <div class="prev-card-month">${r.month}</div>
            <div class="prev-card-date">Submitted: ${submittedDate}</div>
          </div>
          <div class="prev-card-id">#${r.id}</div>
        </div>
        <div class="prev-card-stats">
          <div><span>${r.total_emp}</span> employees</div>
          <div>${r.submitted_by || 'Unknown maker'}</div>
        </div>
        <div class="pending-actions">
          <button class="prev-view-btn" onclick="loadPendingForReview(${r.id})">Review</button>
          <button class="prev-dl-btn" onclick="downloadPending(${r.id})">Download</button>
          <button class="pending-approve-btn" onclick="approvePending(${r.id})">Freeze</button>
        </div>
      </div>`;
  }).join('');
  if (showSection) section.classList.remove('hidden');
}

async function loadPendingForReview(id) {
  const res = await authFetch(`/api/pending/${id}`);
  const data = await res.json();
  if (!res.ok) throw new Error(data.detail || `Server error ${res.status}`);
  sessionId = data.session_id;
  downloadUrl = `/api/pending/${id}/download`;
  currentSummary = {
    ...data.summary,
    employees: data.employees,
  };
  renderResults(currentSummary);
  currentPendingId = id;
  showToast(`Loaded pending register #${id} for review`, 'success');
}

function downloadPending(id) {
  window.location.href = `/api/pending/${id}/download`;
}

async function approvePending(id) {
  try {
    const res = await authFetch(`/api/pending/${id}/freeze`, { method: 'POST' });
    const data = await res.json();
    if (!res.ok) throw new Error(data.detail || `Server error ${res.status}`);
    showToast(data.message || 'Register frozen', 'success');
    loadPendingRegisters(false);
    loadPreviousRegisters(false);
  } catch (err) {
    showToast('Approval failed: ' + err.message, 'error');
  }
}

/* ══════════════════════════════════════════════════════════════════════════
   PREVIOUS REGISTERS TOGGLE (header button)
   ══════════════════════════════════════════════════════════════════════════ */
function togglePrevSection() {
  const section = document.getElementById('prev-section');
  const btn     = document.querySelector('.prev-toggle-btn');
  const hidden  = section.classList.contains('hidden');

  if (hidden) {
    loadPreviousRegisters(true);
    btn.classList.add('active');
  } else {
    section.classList.add('hidden');
    btn.classList.remove('active');
  }
}

function togglePendingSection() {
  const section = document.getElementById('pending-section');
  const btn = document.getElementById('pending-toggle-btn');
  const hidden = section.classList.contains('hidden');
  if (hidden) {
    loadPendingRegisters(true);
    btn.classList.add('active');
  } else {
    section.classList.add('hidden');
    btn.classList.remove('active');
  }
}

/* ══════════════════════════════════════════════════════════════════════════
   TOAST
   ══════════════════════════════════════════════════════════════════════════ */
let _toastTimer = null;
function showToast(msg, type = 'success') {
  const toast = document.getElementById('toast');
  toast.textContent = (type === 'success' ? '✓ ' : '✕ ') + msg;
  toast.className   = `toast ${type}`;
  clearTimeout(_toastTimer);
  _toastTimer = setTimeout(() => { toast.className = 'toast hidden'; }, 3500);
}

/* ══════════════════════════════════════════════════════════════════════════
   FREEZE ZONE HELPERS
   ══════════════════════════════════════════════════════════════════════════ */
function _resetFreezeZone() {
  const zone = document.getElementById('freeze-zone');
  const btn  = document.getElementById('freeze-btn');
  if (!zone || !btn) return;
  zone.classList.remove('frozen-state');
  btn.disabled  = false;
  btn.className = 'freeze-btn';
  btn.textContent = 'Freeze Register';
  document.getElementById('freeze-zone-title').textContent = 'Freeze Register';
  document.getElementById('freeze-zone-sub').textContent =
    'Lock and store this register permanently in the local database';
}

function _setFreezeZoneFrozen(month) {
  const zone = document.getElementById('freeze-zone');
  const btn  = document.getElementById('freeze-btn');
  if (!zone || !btn) return;
  zone.classList.add('frozen-state');
  btn.disabled  = true;
  btn.className = 'freeze-btn frozen';
  btn.textContent = 'Frozen ✓';
  document.getElementById('freeze-zone-title').textContent = `Frozen — ${month}`;
  document.getElementById('freeze-zone-sub').textContent =
    'This register is stored in the local database and available in Previous Registers';
}

/* ══════════════════════════════════════════════════════════════════════════
   RESET UI
   ══════════════════════════════════════════════════════════════════════════ */
function resetUI() {
  pdfFile = null; xlsxFile = null; registerFile = null; downloadUrl = null;
  sessionId = null; currentSummary = null;
  hasUnsavedChanges = false; isFrozen = false;
  currentPendingId = null;
  window.empData = [];

  document.getElementById('pdf-filename').textContent  = '';
  document.getElementById('xlsx-filename').textContent = '';
  document.getElementById('register-filename').textContent = '';
  document.getElementById('pdf-btn-text').textContent  = 'Choose PDF';
  document.getElementById('xlsx-btn-text').textContent = 'Choose XLSX';
  document.getElementById('register-btn-text').textContent = 'Choose Register';
  document.getElementById('pdf-card').classList.remove('has-file');
  document.getElementById('xlsx-card').classList.remove('has-file');
  document.getElementById('register-card').classList.remove('has-file');
  document.getElementById('pdf-input').value  = '';
  document.getElementById('xlsx-input').value = '';
  document.getElementById('register-input').value = '';
  document.getElementById('process-btn').disabled = true;
  document.getElementById('import-register-btn').disabled = true;
  document.getElementById('error-box').classList.add('hidden');
  document.getElementById('results-section').classList.add('hidden');
  document.getElementById('upload-section').classList.remove('hidden');

  // Reset action buttons
  const acceptBtn = document.getElementById('accept-btn');
  acceptBtn.classList.add('hidden');
  acceptBtn.classList.remove('pending', 'saved');
  document.getElementById('accept-btn-text').textContent = 'Accept Changes';
  _resetFreezeZone();
  applyRoleControls();
  document.getElementById('download-sub').textContent =
    'Complete workbook · all employees · all formulas live';
}

/* ══════════════════════════════════════════════════════════════════════════
   INIT
   ══════════════════════════════════════════════════════════════════════════ */
async function initApp() {
  await loadCurrentUser();
  loadPreviousRegisters(false);
  if (isAdmin()) loadPendingRegisters(false);
  applyRoleControls();
}

initApp();
