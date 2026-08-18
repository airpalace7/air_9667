(function () {
  const scriptEl = document.currentScript;
  const src = scriptEl.getAttribute('data-src');

  function escapeHtml(s) {
    if (s === null || s === undefined) return '';
    return String(s).replace(/[&<>"']/g, c => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[c]));
  }

  async function load() {
    const res = await fetch(src);
    const data = await res.json();

    const infoEl = document.getElementById('info');
    infoEl.innerHTML = (data.info || []).map(item => `
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
      <table>
        <thead>
          <tr>
            <th>No</th><th>명칭</th><th>P/N</th><th>S/N</th><th>구분</th>
            <th>교환주기</th><th>장착일</th><th>장착시</th><th>위치</th>
            <th>TSN</th><th>사용시간</th><th>다음교환</th>
            <th>잔여시간</th><th>비고</th>
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

  load().catch(err => {
    document.getElementById('content').innerHTML =
      '<div class="empty">데이터를 불러오지 못했습니다: ' + err.message + '</div>';
  });
})();
