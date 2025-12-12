
/* launchevent.js */

// OnMessageSend（Smart Alerts）
function onMessageSendHandler(event) {
  const item = Office.context.mailbox.item;

  Promise.all([
    getSubjectAsync(item),
    getRecipientsAsync(item.to),
    getRecipientsAsync(item.cc),
    getRecipientsAsync(item.bcc),
    getAttachmentsAsync(item)
  ]).then(([subject, toList, ccList, bccList, attachments]) => {
    const allRecipients = [
      ...label('收件人', toList),
      ...label('副本', ccList),
      ...label('密件副本', bccList)
    ];

    const domains = Array.from(new Set(
      [...toList, ...ccList, ...bccList]
        .map(r => (r.emailAddress || '').split('@')[1])
        .filter(Boolean)
    ));

    // Markdown（支援 1.15）
    const mdLines = [];
    mdLines.push(`**寄送前檢視**`);
    mdLines.push(`- **收件人網域**：${domains.length ? domains.join(', ') : '（無）'}`);
    mdLines.push(`- **主旨**：${escapeMd(subject || '（空白）')}`);
    mdLines.push(`- **收件人**：`);
    mdLines.push(allRecipients.length
      ? allRecipients.map(r => `  - ${r.label}: ${escapeMd(r.displayName)} <${escapeMd(r.emailAddress)}>`).join('\n')
      : `  - （無）`);
    mdLines.push(`- **附件**：`);
    mdLines.push(attachments.length
      ? attachments.map(a => `  - ${escapeMd(a.name)} (${a.size || 0} bytes)`).join('\n')
      : `  - （無）`);
    mdLines.push(`\n請確認是否要寄出。選擇 **仍然寄出** 則立即投遞；選擇 **不寄出** 則返回撰寫畫面。`);

    const markdown = mdLines.join('\n');

    // 提供相容的純文字訊息（舊版用）
    const plain =
      `寄送前檢視\n` +
      `收件人網域：${domains.length ? domains.join(', ') : '（無）'}\n` +
      `主旨：${subject || '（空白）'}\n` +
      `收件人：\n${
        allRecipients.length
          ? allRecipients.map(r => `- ${r.label}: ${r.displayName} <${r.emailAddress}>`).join('\n')
          : '- （無）'
      }\n附件：\n${
        attachments.length
          ? attachments.map(a => `- ${a.name} (${a.size || 0} bytes)`).join('\n')
          : '- （無）'
      }\n請確認是否要寄出。`;

    // 顯示 Smart Alerts 對話（PromptUser）：讓使用者決定是否送出
    event.completed({
      allowEvent: false,                 // 顯示提示；由使用者決定
      errorMessage: plain,               // 舊版相容
      errorMessageMarkdown: markdown     // 新版（1.15）呈現 Markdown
      // 可選：自訂對話按鈕或開 Taskpane/執行函式（Mailbox ≥ 1.14/1.15）
      // cancelLabel: "返回編輯", commandId: "..."  // 需在 manifest 宣告相對應 commandId
    });
  }).catch(() => {
    // 若資料擷取失敗，避免阻塞寄出
    event.completed({ allowEvent: true });
  });
}

// 綁定事件名稱（需與 manifest 的 FunctionName 一致）
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

// ---------- helpers ----------

function getSubjectAsync(item) {
  return new Promise(resolve => {
    item.subject.getAsync(res => resolve(res.status === Office.AsyncResultStatus.Succeeded ? res.value : ''));
  });
}

function getRecipientsAsync(field /* Office.Recipients */) {
  return new Promise(resolve => {
    if (!field || !field.getAsync) return resolve([]);
    field.getAsync(res => {
      const list = (res.status === Office.AsyncResultStatus.Succeeded ? res.value : []).map(r => ({
        displayName: r.displayName || r.emailAddress || '',
        emailAddress: r.emailAddress || '',
        recipientType: r.recipientType || 'other'
      }));
      resolve(list);
    });
  });
}

function getAttachmentsAsync(item) {
  return new Promise(resolve => {
    if (typeof item.getAttachmentsAsync !== 'function') {
      resolve([]); // 部分環境可能不支援 Compose 取附件
      return;
    }
    item.getAttachmentsAsync(res => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) return resolve([]);
      const list = (res.value || []).map(a => ({
        id: a.id, name: a.name, size: a.size, contentType: a.contentType, attachmentType: a.attachmentType, isInline: a.isInline
      }));
      resolve(list);
    });
  });
}

function labelfunction label(label, arr) { return arr.map(r => ({ label, displayName: r.displayName, emailAddress: r.emailAddress })); }

