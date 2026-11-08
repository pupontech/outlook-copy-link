/* Side panel: shows the link for the selected email; updates when the selection changes (if pinned). */
(function () {
  "use strict";
  var $ = function (id) { return document.getElementById(id); };

  function setStatus(msg, isErr) {
    var s = $("status");
    s.textContent = msg || "";
    s.className = isErr ? "err" : "";
  }

  function render() {
    var item = Office.context.mailbox.item;
    var link = CopyLink.currentLink();
    $("subject").textContent = item && item.subject ? item.subject : (item ? "(no subject)" : "Select an email…");
    $("link").value = link;
    $("copy").disabled = $("copyRich").disabled = !link;
    setStatus("");
  }

  function doCopy(rich) {
    var link = $("link").value;
    if (!link) return;
    var item = Office.context.mailbox.item;
    var label = (item && item.subject) ? item.subject : link;
    var p = rich ? CopyLink.copyHyperlink(link, label) : CopyLink.copyText(link);
    p.then(function () { setStatus(rich ? "Hyperlink copied." : "Link copied."); })
     .catch(function () {
        $("link").focus(); $("link").select();
        setStatus("Automatic copy was blocked. The link is selected: press Ctrl+C.", true);
     });
  }

  // Try to copy as soon as the panel opens (used by the one-click "Copy Link" button).
  // Some Outlook builds only allow copying from a click inside the panel; if that happens
  // the link is pre-selected so a single Ctrl+C finishes the job.
  function autoCopy() {
    var link = $("link").value;
    if (!link) return;
    window.focus();
    CopyLink.copyText(link).then(function () {
      setStatus("Link copied to clipboard.");
    }).catch(function () {
      $("link").focus(); $("link").select();
      setStatus("Auto-copy was blocked. Press Ctrl+C (link is selected) or click Copy link.", true);
    });
  }

  Office.onReady(function () {
    $("copy").addEventListener("click", function () { doCopy(false); });
    $("copyRich").addEventListener("click", function () { doCopy(true); });
    render();
    autoCopy(); // TEST BUILD: always on; will be limited to ?copy=1 after testing
    try {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, render);
    } catch (e) { /* older clients: panel simply won't auto-refresh */ }
  });
})();
