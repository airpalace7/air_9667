(function () {
  const scriptEl = document.currentScript;
  const src = scriptEl.getAttribute('data-src');
  let DATA = null;

  function escapeHtml(s) {
    if (s === null || s === undefined) return '';
    return String(s).replace(/[&<>"']/g, c => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[c]));
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

  async function load() {
    const res = await fetch(src);
    const data = await res.json();
    DATA = data;

    const infoAll = data.info || [];
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

    const contentEl = document.getElementById('content');
    if (!data.rows || data.rows.length === 0) {
      contentEl.innerHTML = '<div class="empty">데이터가 없습니다.</div>';
      return;
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
          ${data.rows.map(r => `
            <tr class="${r.overdue ? 'overdue' : ''}">
              <td>${escapeHtml(r.no)}</td>
              <td>${escapeHtml(r.name)}${r.overdue ? '<span class="badge-overdue">초과</span>' : ''}</td>
              <td>${escapeHtml(r.pn)}</td>
              <td>${escapeHtml(r.sn)}</td>
              <td>${escapeHtml(r.type)}</td>
              <td>${escapeHtml(r.exchange_cycle)}</td>
              <td>${escapeHtml(r.installation_date)}</td>
              <td>${escapeHtml(r.installation_time)}</td>
              <td>${escapeHtml(r.location)}</td>
              <td>${escapeHtml(r.tsn)}</td>
              <td>${escapeHtml(r.usage_time)}</td>
              <td>${escapeHtml(r.next_exchange_date)} ${escapeHtml(r.next_exchange_time)}</td>
              <td>${escapeHtml(r.remaining_time)}</td>
              <td class="remark">${escapeHtml(r.note)}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    `;
  }

  function updateStickyOffsets() {
    const header = document.querySelector('header.top');
    const info = document.getElementById('info');
    if (header) document.documentElement.style.setProperty('--header-h', header.offsetHeight + 'px');
    if (info) document.documentElement.style.setProperty('--info-h', info.offsetHeight + 'px');
  }
  window.addEventListener('resize', updateStickyOffsets);

  // 창 크기 변경뿐 아니라, 등록기호 바가 좁은 화면에서 줄바뀜/글자 렌더링 차이로 높이가
  // 미세하게 바뀌는 경우까지 놓치지 않도록 계속 감시한다.
  function setupResizeObservers() {
    if (typeof ResizeObserver === 'undefined') return;
    const header = document.querySelector('header.top');
    const info = document.getElementById('info');
    if (header) new ResizeObserver(updateStickyOffsets).observe(header);
    if (info) new ResizeObserver(updateStickyOffsets).observe(info);
  }

  const printBtn = document.getElementById('printBtn');
  if (printBtn) {
    printBtn.addEventListener('click', () => window.print());
  }

  load().then(updateStickyOffsets).then(setupResizeObservers).catch(err => {
    document.getElementById('content').innerHTML =
      '<div class="empty">데이터를 불러오지 못했습니다: ' + err.message + '</div>';
  });
})();
