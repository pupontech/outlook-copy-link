/* Side panel: shows the link for the selected email; updates when the selection changes (if pinned).
 * Main link  = outlook:<EntryID>  -> opens the message in classic Outlook desktop.
 * Web link   = OWA deep link      -> opens in Outlook on the web (extra, tucked away). */
(function () {
  "use strict";
  var $ = function (id) { return document.getElementById(id); };
  var current = 0; // guards against out-of-order async results when switching emails quickly

  function setStatus(msg, isErr) {
    var s = $("status");
    s.textContent = msg || "";
    s.className = isErr ? "err" : "";
  }

  function render() {
    var token = ++current;
    var item = Office.context.mailbox.item;
    var web = CopyLink.currentLink();
    $("subject").textContent = item && item.subject ? item.subject : (item ? "(no subject)" : "Select an email…");
    $("webLink").value = web;
    $("copyWeb").disabled = !web;
    $("link").value = "";
    $("copy").disabled = $("copyRich").disabled = true;
    setStatus("");
    if (!item || !item.itemId) return;

    $("link").placeholder = "Working…";
    CopyLink.desktopLink(item.itemId).then(function (link) {
      if (token !== current) return;
      $("link").value = link;
      $("copy").disabled = $("copyRich").disabled = !link;
    }).catch(function (e) {
      if (token !== current) return;
      $("link").placeholder = "";
      setStatus("Could not build the desktop link (" + e.message + "). The web link below still works.", true);
    });
  }

  function doCopy(fieldId, rich) {
    var link = $(fieldId).value;
    if (!link) return;
    var item = Office.context.mailbox.item;
    var label = (item && item.subject) ? item.subject : link;
    var p = rich ? CopyLink.copyHyperlink(link, label) : CopyLink.copyText(link);
    p.then(function () { setStatus(rich ? "Hyperlink copied." : "Link copied."); })
     .catch(function () {
        $(fieldId).focus(); $(fieldId).select();
        setStatus("Automatic copy was blocked. The link is selected: press Ctrl+C.", true);
     });
  }

  Office.onReady(function () {
    $("copy").addEventListener("click", function () { doCopy("link", false); });
    $("copyRich").addEventListener("click", function () { doCopy("link", true); });
    $("copyWeb").addEventListener("click", function () { doCopy("webLink", false); });
    render();
    try {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, render);
    } catch (e) { /* older clients: panel simply won't auto-refresh */ }
  });
})();
