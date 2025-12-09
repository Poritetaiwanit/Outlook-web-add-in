Office.initialize = () => {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
};

async function onMessageSendHandler(event) {
  const item = Office.context.mailbox.item;

  const to = await getRecipientInfos(item.to);
  const cc = await getRecipientInfos(item.cc);
  const bcc = await getRecipientInfos(item.bcc);
  const subject = await getSubject(item);
  const attachments = await getAttachments(item);

  const payload = { to, cc, bcc, subject, attachments };

  Office.context.ui.displayDialogAsync(
    "https://poritetaiwanit.github.io/Outlook-web-add-in/confirm.html",
    { height: 60, width: 55, displayInIframe: true },
    (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        event.completed({ allowEvent: false });
        return;
      }
      const dialog = result.value;

      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        const msg = JSON.parse(arg.message);
        if (msg.type === "DECISION") {
          dialog.close();
          event.completed({ allowEvent: msg.allow });
        }
      });

      dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
        event.completed({ allowEvent: false });
      });

      dialog.messageParent(JSON.stringify({ type: "INIT", payload }));
    }
  );
}

function getRecipientInfos(field) {
  return new Promise((resolve) => {
    if (!field) return resolve([]);
    field.getAsync((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        resolve(res.value.map(r => ({
          emailAddress: r.emailAddress,
          displayName: r.displayName || "",
          domain: r.emailAddress.split("@")[1] || ""
        })));
      } else resolve([]);
    });
  });
}

function getSubject(item) {
  return new Promise((resolve) => {
    item.subject.getAsync((res) => {
      resolve(res.status === Office.AsyncResultStatus.Succeeded ? res.value : "");
    });
  });
}

function getAttachments(item) {
  return new Promise((resolve) => {
    item.getAttachmentsAsync((res) => {
      resolve(res.status === Office.AsyncResultStatus.Succeeded
        ? res.value.map(a => ({ name: a.name, type: a.attachmentType, isInline: a.isInline }))
        : []);
    });
  });
}