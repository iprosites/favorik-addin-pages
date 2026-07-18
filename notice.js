"use strict";
(() => {
  // src/notice/notice.ts
  Office.onReady(() => {
    const params = new URLSearchParams(window.location.search);
    const message = params.get("message") ?? "";
    const messageEl = document.getElementById("message");
    messageEl.textContent = message;
    const okButton = document.getElementById("okButton");
    okButton.addEventListener("click", () => Office.context.ui.messageParent("ok"));
  });
})();
//# sourceMappingURL=notice.js.map
