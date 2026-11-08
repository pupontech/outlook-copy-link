/* Shared helpers: build the web link for the selected email and copy text to the clipboard. */
(function (root) {
  "use strict";

  // Change this if your mailbox is not on Microsoft 365 / Exchange Online:
  //   Microsoft 365 work/school : https://outlook.office365.com/owa/   (default)
  //   Outlook.com / Hotmail     : https://outlook.live.com/owa/
  var OWA_BASE = "https://outlook.office365.com/owa/";

  /** Build the direct link from an EWS item id (what Office.js returns for itemId). */
  function buildLink(itemId, base) {
    if (!itemId) return "";
    return (base || OWA_BASE) +
      "?ItemID=" + encodeURIComponent(itemId) +
      "&exvsurl=1&viewmodel=ReadMessageItem";
  }

  /** Link for the email currently selected in Outlook. Returns "" if it has no id yet. */
  function currentLink() {
    var item = root.Office && Office.context && Office.context.mailbox && Office.context.mailbox.item;
    return item && item.itemId ? buildLink(item.itemId) : "";
  }

  /** Copy plain text. Tries the async clipboard API first, then the legacy execCommand. */
  function copyText(text) {
    function legacy() {
      return new Promise(function (resolve, reject) {
        var ta = document.createElement("textarea");
        ta.value = text;
        ta.setAttribute("readonly", "");
        ta.style.cssText = "position:fixed;top:0;left:0;opacity:0;";
        document.body.appendChild(ta);
        ta.select();
        ta.setSelectionRange(0, text.length);
        var ok = false;
        try { ok = document.execCommand("copy"); } catch (e) { ok = false; }
        document.body.removeChild(ta);
        ok ? resolve() : reject(new Error("copy blocked"));
      });
    }
    if (navigator.clipboard && navigator.clipboard.writeText) {
      return navigator.clipboard.writeText(text).catch(legacy);
    }
    return legacy();
  }

  /** Copy a clickable hyperlink (rich text) so it pastes as a link in Word, Teams, Trello, etc. */
  function copyHyperlink(url, label) {
    return new Promise(function (resolve, reject) {
      var div = document.createElement("div");
      var a = document.createElement("a");
      a.href = url;
      a.textContent = label || url;
      div.appendChild(a);
      div.style.cssText = "position:fixed;top:0;left:0;opacity:0;";
      document.body.appendChild(div);
      var range = document.createRange();
      range.selectNodeContents(div);
      var sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
      var ok = false;
      try { ok = document.execCommand("copy"); } catch (e) { ok = false; }
      sel.removeAllRanges();
      document.body.removeChild(div);
      ok ? resolve() : reject(new Error("copy blocked"));
    });
  }

  var api = { OWA_BASE: OWA_BASE, buildLink: buildLink, currentLink: currentLink,
              copyText: copyText, copyHyperlink: copyHyperlink };
  root.CopyLink = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
