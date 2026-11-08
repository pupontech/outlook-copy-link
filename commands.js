/* Ribbon "Copy Link" button: copies the link to the selected email and shows a brief confirmation. */
Office.onReady(function () {});

function notify(message, type) {
  try {
    var item = Office.context.mailbox.item;
    item.notificationMessages.replaceAsync("copyLinkStatus", {
      type: type || Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: message,
      icon: "Icon.16x16",
      persistent: false
    });
  } catch (e) { /* notifications are best-effort */ }
}

function copyLink(event) {
  var link = CopyLink.currentLink();
  if (!link) {
    notify("No link available yet. Open the email, or try the Link Panel.",
      Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage);
    event.completed();
    return;
  }
  CopyLink.copyText(link).then(function () {
    notify("Link copied to clipboard.");
  }).catch(function () {
    notify("Could not copy automatically. Click \"Link Panel\" and use its Copy button.",
      Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage);
  }).then(function () {
    event.completed();
  });
}

// Required so the manifest's <FunctionName> can find the function.
Office.actions.associate("copyLink", copyLink);
