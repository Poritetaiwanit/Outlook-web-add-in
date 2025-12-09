Office.initialize = () => {};

Office.context.ui.addHandlerAsync(
  Office.EventType.DialogParentMessageReceived,
  (arg) => {
    const msg = JSON.parse(arg.message);
    if (msg.type !== "INIT") return;
    const { subject, to, cc, bcc, attachments } = msg.payload;

    document.getElementById("subject").textContent = subject || "(none)";
    document.getElementById("recipients").innerHTML = [
      renderGroup("To", to),
      renderGroup("CC", cc),
      renderGroup("BCC", bcc)
    ].join("");
    document.getElementById("attachments").innerHTML =
      attachments.length
        ? `<ul>${attachments.map(a => `<li>${escapeHTML(a.name)}${a.isInline ? " (inline)" : ""}</li>`).join("")}</ul>`
        : "(none)";
  }
);

document.getElementById("btnSend").onclick = () => {
  Office.context.ui.messageParent(JSON.stringify({ type: "DECISION", allow: true }));
};
document.getElementById("btnCancel").onclick = () => {
  Office.context.ui.messageParent(JSON.stringify({ type: "DECISION", allow: false }));
};

function renderGroup(label, list) {
  const items = list.map(r =>
    `<li><span class="info">${escapeHTML(r.emailAddress)}</span> — ${escapeHTML(r.displayName)} (domain: ${escapeHTML(r.domain)})</li>`
  ).join("");
  return `<div><strong>${label}:</strong> ${list.length ? `<ul>${items}</ul>` : "(none)"}</div>`;
}

function escapeHTML(s) {
  return String(s).replace(/[&<>"']/g, ch => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[ch]));
}