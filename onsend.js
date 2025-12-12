
/* onsend.js */

function onSendReview(event) {
  const item = Office.context.mailbox.item;

  // 為避免收件人尚未解析，先嘗試儲存草稿
  item.saveAsync(() => {
    Promise.all([
      getRecipientsAsync(item.to),
      getRecipientsAsync(item.cc),
      getRecipientsAsync(item.bcc),
      getSubjectAsync(item),
      getAttachmentsAsync(item)
    ])
    .then(([toList, ccList, bccList, subject, attachments]) => {
      const recipientsLabeled = [
        ...labelRecipients('收件人', toList),
        ...labelRecipients('副本', ccList),
        ...labelRecipients('密件副本', bccList)
      ];

      const domains = Array.from(new Set(
        [...toList, ...ccList, ...bccList]
          .map(r => (r.emailAddress || '').split('@')[1])
          .filter(Boolean)
      ));

      const payload = { subject, recipients: recipientsLabeled, domains, attachments };

      const dialogUrl =
        'https://ashy-smoke-03b7c5800.3.azurestaticapps.net/dialog.html?data=' +
        encodeURIComponent(JSON.stringify(payload));

      Office.context.ui.displayDialogAsync(dialogUrl,
        { height: 60, width: 60, displayInIframe: true },
        (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            // 對話無法開啟時，不要卡住使用者
            event.completed({ allowEvent: true });
            return;
          }
          const dialog = asyncResult.value;

          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
            try {
              const msg = JSON.parse(args.message);
              if (msg.action === 'send') {
                dialog.close();
                event.completed({ allowEvent: true });   // 允許寄出
              } else {
                dialog.close();
                event.completed({ allowEvent: false });  // 阻擋寄出
              }
            } catch (e) {
              dialog.close();
              event.completed({ allowEvent: false });
            }
          });
      });
    })
    .catch(() => {
      // 蒐集資料失敗時，避免阻塞寄出
      event.completed({ allowEvent: true });
    });
  });
}

// 綁定事件（名稱需與 manifest 的 FunctionName 一致）
if (Office && Office.actions) {
  Office.actions.associate('onSendReview', onSendReview);
}

// ------- helpers -------

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
      resolve([]); // 某些環境不支援
      return;
    }
    item.getAttachmentsAsync(res => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) return resolve([]);
      const list = (res.value || []).map(a => ({
        id: a.id, name: a.name, size: a.size, contentType: a.contentType, attachmentType: a.attachmentType
      }));
      resolve(list);
    });
  });
}

function labelRecipients(label, arr) {
  return arr.map(r => ({ label, displayName: r.displayName, emailAddress: r.emailAddress }));
}
