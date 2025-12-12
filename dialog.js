
(function () {
  const params = new URLSearchParams(location.search);
  const data = JSON.parse(decodeURIComponent(params.get('data') || '{}'));

  const dom = document.getElementById('domains');
  dom.innerHTML = `<strong>收件人網域：</strong> ${ (data.domains || []).join(', ') || '（無）' }`;

  const subj = document.getElementById('subject');
  subj.innerHTML = `<strong>郵件主旨：</strong> ${ escapeHtml(data.subject || '（空白）') }`;

  const rec = document.getElementById('recipients');
  rec.innerHTML = `<strong>收件人：</strong><div class="list">${
    (data.recipients || []).map(r =>
      `<div class="row"><span class="label">${r.label}</span>
        <span>${escapeHtml(r.displayName)}</span>
        &lt;<span class="email">${escapeHtml(r.emailAddress)}</span>&gt;
       </div>`).join('') || '（無）'
  }</div>`;

  const att = document.getElementById('attachments');
  att.innerHTML = `<strong>附件：</strong><div class="list">${
    (data.attachments || []).map(a =>
      `<div class="row">${escapeHtml(a.name)} <span style="color:#666">(${a.size || 0} bytes)</span></div>`
    ).join('') || '（無）'
  }</div>`;

  document.getElementById('btnSend').onclick = () => {
    Office.context.ui.messageParent(JSON.stringify({ action: 'send' }));
  };
  document.getElementById('btnCancel').onclick = () => {
    Office.context.ui.messageParent(JSON.stringify({ action: 'cancel' }));
  };

  function escapeHtml(s) {
    return String(s || '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;', "'":'&#39;'}[c]));
  }
})();
