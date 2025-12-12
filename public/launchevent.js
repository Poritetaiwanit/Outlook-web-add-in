Office.onReady(async () => {
  console.log("Add-in ready");
});

async function onMessageSendHandler(event) {
  const item = Office.context.mailbox.item;

  const recipients = item.to.getAsync
    ? await item.to.getAsync()
    : item.to;

  const subject = item.subject;
  const attachments = item.attachments || [];

  window.emailInfo = { recipients, subject, attachments, event };

  window.location.href = "runtime.html";
}

window.onload = function () {
  if (!window.emailInfo) return;

  const { recipients, subject, attachments } = window.emailInfo;

  const info = document.getElementById("info");

  info.innerHTML = `
    <p><b>主旨：</b>${subject}</p>

    <p><b>收件人：</b></p>
    <ul>
      ${recipients.map(r => `<li>${r.displayName} &lt;${r.emailAddress}&gt;</li>`).join("")}
    </ul>

    <p><b>附件：</b></p>
    <ul>
      ${attachments.map(a => `<li>${a.name}</li>`).join("")}
    </ul>
  `;
};

function sendEmail() {
  window.emailInfo.event.completed({ allowEvent: true });
}

function cancelSend() {
  window.emailInfo.event.completed({ allowEvent: false });
}
