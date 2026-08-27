(function () {
  const scriptEl = document.currentScript;
  const src = scriptEl.getAttribute('data-src');
  let DATA = null;
  let REFS = {};       // 01 메인정비표(insp.json)에서 가져온 현재 기준값들
  let FILE_SHA = null;
  let ADMIN = false;

  function escapeHtml(s) {
    if (s === null || s === undefined) return '';
    return String(s).replace(/[&<>"']/g, c => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[c]));
  }

  function hoursToStr(h) {
    if (h === null || h === undefined || isNaN(h)) return '';
    const neg = h < 0;
    h = Math.abs(h);
    const totalHours = Math.floor(h + 1e-9);
    const mins = Math.round((h - totalHours) * 60);
    return (neg ? '-' : '') + totalHours + ':' + String(mins).padStart(2, '0');
  }

  // "1234:56" 같은 hhhh:mm 문자열을 소수 시간으로 변환. 숫자만 입력해도 허용.
  function parseHoursStr(s) {
    if (s === null || s === undefined) return NaN;
    s = String(s).trim();
    if (s === '' || s === '-') return NaN;
    const neg = s.startsWith('-');
    if (neg) s = s.slice(1);
    const m = s.match(/^(\d+):([0-5]?\d)$/);
    if (m) {
      const h = parseInt(m[1], 10);
      const mm = parseInt(m[2], 10);
      return (neg ? -1 : 1) * (h + mm / 60);
    }
    const f = parseFloat(s);
    return isNaN(f) ? NaN : (neg ? -1 : 1) * Math.abs(f);
  }

  // "2026-3-19" 처럼 구분자/자릿수가 느슨해도 "YYYY-MM-DD"로 정규화.
  function parseDateStr(s) {
    if (s === null || s === undefined) return null;
    s = String(s).trim();
    if (s === '') return null;
    const m = s.match(/^(\d{4})[-./](\d{1,2})[-./](\d{1,2})$/);
    if (!m) return null;
    return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  }

  // 시간(hours)/사이클(cycles) 값 종류에 따라 표시 문자열을 다르게 만든다.
  function formatByKind(v, kind) {
    if (v === null || v === undefined || isNaN(v)) return '';
    if (kind === 'hours') return hoursToStr(v);
    const rounded = Math.round(v * 100) / 100;
    return String(rounded);
  }
  function parseByKind(s, kind) {
    if (kind === 'hours') return parseHoursStr(s);
    if (s === null || s === undefined || String(s).trim() === '') return NaN;
    const f = parseFloat(s);
    return isNaN(f) ? NaN : f;
  }

  function isoDate(d) {
    return d.toISOString().slice(0, 10);
  }
  function daysBetweenIso(iso1, iso2) {
    const d1 = new Date(iso1 + 'T00:00:00Z');
    const d2 = new Date(iso2 + 'T00:00:00Z');
    return Math.round((d2 - d1) / 86400000);
  }
  function addDaysIso(iso, days) {
    const d = new Date(iso + 'T00:00:00Z');
    d.setUTCDate(d.getUTCDate() + days);
    return isoDate(d);
  }

  // 부품/엔진 한 행의 사용시간/다음교환/잔여시간을 실시간으로 재계산한다.
  // 표준 패턴(=SUM($C$기준행-H+K또는L) 등)을 따르는 행만(row.recalc) 계산하고,
  // 그 외(row.has_interval만 true)는 "엑셀관리"로 표시하며 엑셀 계산값을 그대로 둔다.
  function recomputeRow(row) {
    if (row.installation_date_iso) {
      row.installation_date = row.installation_date_iso;
    }
    if (row.installation_time_num !== null && row.installation_time_num !== undefined && !isNaN(row.installation_time_num)) {
      row.installation_time = formatByKind(row.installation_time_num, row.value_kind);
    }
    if (row.exchange_cycle_num !== null && row.exchange_cycle_num !== undefined && !isNaN(row.exchange_cycle_num)) {
      row.exchange_cycle = formatByKind(row.exchange_cycle_num, row.value_kind);
    }

    if (!row.recalc || !row.ref_key) return;
    const refVal = REFS[row.ref_key];
    if (refVal === undefined || refVal === null) return;
    const install = row.installation_time_num;
    const cycle = row.exchange_cycle_num;
    if (install === null || install === undefined || isNaN(install)) return;
    if (cycle === null || cycle === undefined || isNaN(cycle)) return;
    const baseline = row.baseline_value || 0;

    const usage = refVal - install + baseline;
    const nextEx = install + cycle - baseline;
    const remain = nextEx - refVal;

    row.usage_time = formatByKind(usage, row.value_kind);
    row.next_exchange_time = formatByKind(nextEx, row.value_kind);
    row.remaining_time = formatByKind(remain, row.value_kind);
    row.overdue = remain < 0;
  }

  // "달력 일수" 기준(예: 5년/10년마다 교환) 항목의 다음교환일/잔여일을 계산한다.
  // 기체(AC) 시트: 다음교환일 = 장착일 + 교환일수 - 보정일수(대부분 0).
  // 엔진(ENG) 시트: 별도 "기준 시작일"이 있으면 그 날짜를 쓰고(재장착돼도 원래 정밀점검
  // 시점부터 계산 이어감), 없으면 장착일을 기준으로 쓴다.
  function recomputeDateRow(row, todayIso) {
    if (!row.date_recalc || !row.date_interval_days) return;
    let baseIso = row.installation_date_iso;
    if (row.date_mode === 'ac') {
      if (!baseIso) return;
      const nextIso = addDaysIso(baseIso, row.date_interval_days - (row.date_baseline_offset_days || 0));
      row.next_exchange_date_iso = nextIso;
    } else {
      baseIso = row.date_baseline_override_iso || row.installation_date_iso;
      if (!baseIso) return;
      row.next_exchange_date_iso = addDaysIso(baseIso, row.date_interval_days);
    }
    row.next_exchange_date = row.next_exchange_date_iso;
    if (todayIso) {
      const remainDays = daysBetweenIso(todayIso, row.next_exchange_date_iso);
      row.remaining_days = String(remainDays);
      if (remainDays < 0) row.overdue = true;
    }
  }

  // REMAIN 칸에 넣을 데이터 막대 스타일. interval을 100%로 보고 남은 비율만큼 채운다.
  function barStyle(remain, interval) {
    if (remain === null || remain === undefined || isNaN(remain)) return '';
    if (!interval || isNaN(interval) || interval <= 0) return '';
    const pct = Math.max(0, Math.min(100, (remain / interval) * 100));
    if (pct <= 0) return '';
    return ` style="background: linear-gradient(to right, var(--databar) ${pct}%, transparent ${pct}%);"`;
  }

  function recomputeAll() {
    if (!DATA) return;
    const todayIso = REFS.today;
    DATA.rows.forEach(row => {
      recomputeRow(row);
      recomputeDateRow(row, todayIso);
    });
  }

  // 인쇄 시 각 페이지 맨 위에 자동으로 반복되도록, thead 안에 제목+등록기호 정보 두 줄을
  // 넣어둔다. thead는 브라우저가 인쇄할 때 페이지가 나뉠 때마다 자동으로 다시 보여주는
  // 기본 기능이라, 픽셀 계산 없이도 몇 페이지가 되든 항상 정확히 나온다.
  function buildPrintTheadRows(todayLabel, gridItems) {
    const titleText = 'HL9667 MAINT STATUS - 기준일 ' + (todayLabel || '');
    const infoText = gridItems.map(item =>
      `<span class="info-piece">${escapeHtml(item.label)}<b>${escapeHtml(item.value)}</b></span>`
    ).join('');
    return `
      <tr class="print-only-row"><th colspan="14" class="print-thead-title">${escapeHtml(titleText)}</th></tr>
      <tr class="print-only-row"><th colspan="14" class="print-thead-info">${infoText}</th></tr>
    `;
  }

  function renderInfo() {
    const infoAll = DATA.info || [];
    const todayItem = infoAll.find(it => it.label === 'Today');
    const gridItems = infoAll.filter(it => it.label !== 'Today');

    const dateEl = document.getElementById('headerDate');
    if (dateEl) {
      dateEl.innerHTML = '기준일<strong>' + escapeHtml(todayItem ? todayItem.value : '') + '</strong>';
    }
    const infoEl = document.getElementById('info');
    infoEl.innerHTML = gridItems.map(item => `
      <div class="info-cell">
        <div class="k">${escapeHtml(item.label)}</div>
        <div class="v">${escapeHtml(item.value)}</div>
      </div>
    `).join('');
    return { todayItem, gridItems };
  }

  function renderContent() {
    const contentEl = document.getElementById('content');
    if (!DATA.rows || DATA.rows.length === 0) {
      contentEl.innerHTML = '<div class="empty">데이터가 없습니다.</div>';
      return;
    }
    const infoAll = DATA.info || [];
    const todayItem = infoAll.find(it => it.label === 'Today');
    const gridItems = infoAll.filter(it => it.label !== 'Today');

    function editableCell(value, editable, ri, field, kind) {
      if (!ADMIN || !editable) return escapeHtml(value);
      if (field === 'installation_date_iso') {
        return `<input type="text" placeholder="yyyy-mm-dd" class="cell-input" data-ri="${ri}" data-field="${field}" value="${escapeHtml(value || '')}">`;
      }
      if (field === 'exchange_cycle_num' || field === 'installation_time_num') {
        const placeholder = kind === 'hours' ? 'hhhh:mm' : '숫자';
        return `<input type="text" placeholder="${placeholder}" class="cell-input" data-ri="${ri}" data-field="${field}" value="${escapeHtml(value || '')}">`;
      }
      return `<input type="text" class="cell-input" data-ri="${ri}" data-field="${field}" value="${escapeHtml(value || '')}">`;
    }

    contentEl.innerHTML = `
      <table class="comp-table">
        <thead>
          ${buildPrintTheadRows(todayItem ? todayItem.value : '', gridItems)}
          <tr>
            <th>No</th><th>명칭</th><th>P/N</th><th>S/N</th><th>구분</th>
            <th>교환주기</th><th>장착일</th><th>장착시</th><th>위치</th>
            <th>TSN</th><th>사용시간</th><th>다음교환</th>
            <th>잔여시간</th><th class="remark-th">비고</th>
          </tr>
        </thead>
        <tbody>
          ${DATA.rows.map((r, ri) => `
            <tr class="${r.overdue ? 'overdue' : ''}">
              <td>${escapeHtml(r.no)}</td>
              <td>${ADMIN ? editableCell(r.name, r.name_editable, ri, 'name') : escapeHtml(r.name)}${r.overdue ? '<span class="badge-overdue">초과</span>' : ''}${(ADMIN && r.has_interval && !r.recalc) ? '<span class="badge-manual" title="이 항목은 자동 재계산되지 않아요. 엑셀에서 관리하세요.">엑셀관리</span>' : ''}${(ADMIN && r.date_interval_days !== null && !r.date_recalc) ? '<span class="badge-manual" title="교환일자 계산이 자동화되지 않아요. 엑셀에서 관리하세요.">엑셀관리(일자)</span>' : ''}</td>
              <td>${editableCell(r.pn, r.pn_editable, ri, 'pn')}</td>
              <td>${editableCell(r.sn, r.sn_editable, ri, 'sn')}</td>
              <td>${editableCell(r.type, r.type_editable, ri, 'type')}</td>
              <td>
                ${editableCell(r.exchange_cycle, r.exchange_cycle_editable, ri, 'exchange_cycle_num', r.value_kind)}
                ${(ADMIN && r.date_interval_days !== null) ? `<br><input type="text" placeholder="일수" class="cell-input" data-ri="${ri}" data-field="date_interval_days" value="${escapeHtml(r.date_interval_days)}">` : (r.date_interval_days !== null ? `<br><small>${escapeHtml(r.date_interval_days)}일</small>` : '')}
              </td>
              <td>${editableCell(r.installation_date_iso, r.installation_date_editable, ri, 'installation_date_iso')}</td>
              <td>${editableCell(r.installation_time, r.installation_time_editable, ri, 'installation_time_num', r.value_kind)}</td>
              <td>${editableCell(r.location, r.location_editable, ri, 'location')}</td>
              <td>${escapeHtml(r.tsn)}</td>
              <td>${escapeHtml(r.usage_time)}</td>
              <td>${escapeHtml(r.next_exchange_date_iso || r.next_exchange_date)} ${escapeHtml(r.next_exchange_time)}</td>
              <td${(() => {
                if (r.recalc && r.ref_key) return barStyle(parseByKind(r.remaining_time, r.value_kind), r.exchange_cycle_num);
                if (r.date_recalc) return barStyle(parseInt(r.remaining_days, 10), r.date_interval_days);
                return '';
              })()}>${escapeHtml(r.remaining_time)}${(r.remaining_time && r.remaining_days) ? ' / ' : ''}${escapeHtml(r.remaining_days)}${r.remaining_days ? '일' : ''}</td>
              <td class="remark">${ADMIN ? `<textarea rows="1" class="cell-input remark-input" data-ri="${ri}" data-field="note">${escapeHtml(r.note || '')}</textarea>` : escapeHtml(r.note)}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    `;

    if (ADMIN) {
      contentEl.querySelectorAll('.cell-input').forEach(inp => {
        inp.addEventListener('change', e => {
          const ri = +e.target.dataset.ri, field = e.target.dataset.field;
          const row = DATA.rows[ri];
          const raw = e.target.value;
          if (field === 'installation_date_iso') {
            row[field] = raw.trim() === '' ? null : parseDateStr(raw);
          } else if (field === 'exchange_cycle_num' || field === 'installation_time_num') {
            row[field] = raw.trim() === '' ? null : parseByKind(raw, row.value_kind);
          } else if (field === 'date_interval_days') {
            row[field] = raw.trim() === '' ? null : parseInt(raw, 10);
          } else {
            row[field] = raw;
          }
          onDataChanged();
        });
      });
    }
  }

  function onDataChanged() {
    recomputeAll();
    renderInfo();
    renderContent();
    markDirty();
    updateStickyOffsets();
  }

  function markDirty() {
    const el = document.getElementById('saveMsg');
    if (el) el.textContent = '저장되지 않은 변경사항이 있습니다';
  }

  function updateStickyOffsets() {
    const header = document.querySelector('header.top');
    const info = document.getElementById('info');
    if (header) document.documentElement.style.setProperty('--header-h', header.offsetHeight + 'px');
    if (info) document.documentElement.style.setProperty('--info-h', info.offsetHeight + 'px');
  }
  window.addEventListener('resize', updateStickyOffsets);

  function setupResizeObservers() {
    if (typeof ResizeObserver === 'undefined') return;
    const header = document.querySelector('header.top');
    const info = document.getElementById('info');
    if (header) new ResizeObserver(updateStickyOffsets).observe(header);
    if (info) new ResizeObserver(updateStickyOffsets).observe(info);
  }

  async function load() {
    const res = await fetch(src);
    DATA = await res.json();

    // 01 메인정비표(insp.json)의 현재 기준값(A/C TSN, L/D CYCLE, ENG TSN, GG/PT CSN 등)을
    // 가져와 계산에 쓴다. 이 값들은 01에서 관리자모드로 수정하면 여기서도 그대로 반영된다.
    try {
      const inspRes = await fetch('data/insp.json');
      const inspData = await inspRes.json();
      REFS = inspData.raw || {};
    } catch (e) {
      REFS = {};
    }

    recomputeAll();
    renderInfo();
    renderContent();
  }

  /* ---------- GitHub 연동 (01 메인정비표와 동일한 방식, 같은 계정 정보 재사용) ---------- */
  function utf8_to_b64(str) {
    return btoa(unescape(encodeURIComponent(str)));
  }
  function ghHeaders(token) {
    return { 'Authorization': 'Bearer ' + token, 'Accept': 'application/vnd.github+json' };
  }
  async function ghFetchFile(owner, repo, token) {
    const url = `https://api.github.com/repos/${owner}/${repo}/contents/${src}`;
    const res = await fetch(url, { headers: ghHeaders(token) });
    if (!res.ok) throw new Error('저장소/토큰을 확인하세요 (' + res.status + ')');
    const json = await res.json();
    return json.sha;
  }
  async function ghSaveFile(owner, repo, token, sha, content) {
    const url = `https://api.github.com/repos/${owner}/${repo}/contents/${src}`;
    const res = await fetch(url, {
      method: 'PUT',
      headers: ghHeaders(token),
      body: JSON.stringify({
        message: '웹 관리자 모드에서 부품현황 업데이트',
        content: utf8_to_b64(content),
        sha: sha,
        branch: 'main'
      })
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.message || ('저장 실패 (' + res.status + ')'));
    }
    const json = await res.json();
    return json.content.sha;
  }

  function setAdmin(on) {
    ADMIN = on;
    document.getElementById('adminToggleBtn').textContent = on ? '관리자 모드 끄기' : '관리자 모드';
    document.getElementById('adminToggleBtn').classList.toggle('active', on);
    document.getElementById('saveBar').style.display = on ? 'flex' : 'none';
    renderContent();
  }

  const printBtn = document.getElementById('printBtn');
  if (printBtn) printBtn.addEventListener('click', () => window.print());

  const adminBtn = document.getElementById('adminToggleBtn');
  if (adminBtn) {
    adminBtn.addEventListener('click', () => {
      if (ADMIN) { setAdmin(false); return; }
      const panel = document.getElementById('connectPanel');
      panel.style.display = panel.style.display === 'none' ? 'block' : 'none';
      const savedOwner = localStorage.getItem('gh_owner');
      const savedRepo = localStorage.getItem('gh_repo');
      const savedToken = localStorage.getItem('gh_token');
      if (savedOwner) document.getElementById('ghOwner').value = savedOwner;
      if (savedRepo) document.getElementById('ghRepo').value = savedRepo;
      if (savedToken) document.getElementById('ghToken').value = savedToken;
    });
  }

  const ghConnectBtn = document.getElementById('ghConnectBtn');
  if (ghConnectBtn) {
    ghConnectBtn.addEventListener('click', async () => {
      const owner = document.getElementById('ghOwner').value.trim();
      const repo = document.getElementById('ghRepo').value.trim();
      const token = document.getElementById('ghToken').value.trim();
      const msgEl = document.getElementById('connectMsg');
      msgEl.textContent = '확인 중...';
      try {
        FILE_SHA = await ghFetchFile(owner, repo, token);
        localStorage.setItem('gh_owner', owner);
        localStorage.setItem('gh_repo', repo);
        localStorage.setItem('gh_token', token);
        msgEl.textContent = '연결됨';
        document.getElementById('connectPanel').style.display = 'none';
        setAdmin(true);
      } catch (err) {
        msgEl.textContent = '오류: ' + err.message;
      }
    });
  }

  const saveBtn = document.getElementById('saveBtn');
  if (saveBtn) {
    saveBtn.addEventListener('click', async () => {
      const owner = localStorage.getItem('gh_owner');
      const repo = localStorage.getItem('gh_repo');
      const token = localStorage.getItem('gh_token');
      const saveMsg = document.getElementById('saveMsg');
      saveMsg.textContent = '저장 중...';
      try {
        FILE_SHA = await ghSaveFile(owner, repo, token, FILE_SHA, JSON.stringify(DATA, null, 2));
        saveMsg.textContent = '저장 완료 (1분 내 반영)';
      } catch (err) {
        saveMsg.textContent = '오류: ' + err.message;
      }
    });
  }

  load().then(updateStickyOffsets).then(setupResizeObservers).catch(err => {
    document.getElementById('content').innerHTML =
      '<div class="empty">데이터를 불러오지 못했습니다: ' + err.message + '</div>';
  });
})();
