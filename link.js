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

  /**
   * Convert the EWS item id into Outlook's own EntryID (hex) using the mailbox's EWS ConvertId call.
   * The result is used for outlook:<EntryID> links, which open the message in classic Outlook for Windows.
   */
  function entryIdHex(itemId) {
    return new Promise(function (resolve, reject) {
      var mb = Office.context.mailbox;
      var email = (mb.userProfile && mb.userProfile.emailAddress) || "";
      var esc = function (v) { return String(v).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;"); };
      var xml =
        '<?xml version="1.0" encoding="utf-8"?>' +
        '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
        ' xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
        ' xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
        ' xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
        '<soap:Header><t:RequestServerVersion Version="Exchange2013"/></soap:Header>' +
        '<soap:Body>' +
        '<ConvertId xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" DestinationFormat="HexEntryId">' +
        '<SourceIds><t:AlternateId Format="EwsId" Id="' + esc(itemId) + '" Mailbox="' + esc(email) + '"/></SourceIds>' +
        '</ConvertId></soap:Body></soap:Envelope>';
      mb.makeEwsRequestAsync(xml, function (res) {
        if (res.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error((res.error && res.error.message) || "EWS request failed"));
          return;
        }
        var tag = /<(?:\w+:)?AlternateId\b[^>]*>/.exec(res.value || "");
        var id = tag && /\bId="([0-9A-Fa-f]+)"/.exec(tag[0]);
        if (!id) { reject(new Error("Outlook did not return an entry id")); return; }
        resolve(id[1].toUpperCase());
      });
    });
  }

  /** Link that opens the message in classic Outlook desktop. Resolves to "" if it cannot be built. */
  function desktopLink(itemId) {
    if (!itemId) return Promise.resolve("");
    return entryIdHex(itemId).then(function (hex) { return "outlook:" + hex; });
  }

  var api = { OWA_BASE: OWA_BASE, buildLink: buildLink, currentLink: currentLink,
              entryIdHex: entryIdHex, desktopLink: desktopLink,
              copyText: copyText, copyHyperlink: copyHyperlink };
  root.CopyLink = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
