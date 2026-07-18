"use strict";
(() => {
  // src/pairing/pairing.ts
  Office.onReady(() => {
    const codeInput = document.getElementById("code");
    const okButton = document.getElementById("okButton");
    const cancelButton = document.getElementById("cancelButton");
    const submit = () => {
      const code = codeInput.value.trim();
      if (code.length > 0) {
        Office.context.ui.messageParent(code);
      }
    };
    okButton.addEventListener("click", submit);
    cancelButton.addEventListener("click", () => Office.context.ui.messageParent(""));
    codeInput.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        submit();
      }
    });
    codeInput.focus();
  });
})();
//# sourceMappingURL=pairing.js.map
