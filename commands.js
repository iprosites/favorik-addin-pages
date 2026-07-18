"use strict";
(() => {
  // src/shared/pairing-storage.ts
  var TOKEN_KEY = "favorikCompanionToken";
  function getStoredToken() {
    try {
      const roaming = Office.context.roamingSettings?.get(TOKEN_KEY);
      if (typeof roaming === "string" && roaming.length > 0) {
        return roaming;
      }
    } catch {
    }
    try {
      return localStorage.getItem(TOKEN_KEY);
    } catch {
      return null;
    }
  }
  function storeToken(token) {
    try {
      localStorage.setItem(TOKEN_KEY, token);
    } catch {
    }
    try {
      Office.context.roamingSettings?.set(TOKEN_KEY, token);
      Office.context.roamingSettings?.saveAsync();
    } catch {
    }
  }

  // src/shared/pairing-dialog.ts
  var PAIRING_DIALOG_URL = "https://localhost:48766/pairing.html";
  function showPairingDialog() {
    return new Promise((resolve) => {
      Office.context.ui.displayDialogAsync(
        PAIRING_DIALOG_URL,
        { height: 30, width: 30, displayInIframe: false },
        (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            resolve(null);
            return;
          }
          const dialog = asyncResult.value;
          let settled = false;
          const finish = (token) => {
            if (settled) {
              return;
            }
            settled = true;
            dialog.close();
            resolve(token);
          };
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
            const message = arg.message;
            if (message) {
              storeToken(message);
              finish(message);
            } else {
              finish(null);
            }
          });
          dialog.addEventHandler(Office.EventType.DialogEventReceived, () => finish(null));
        }
      );
    });
  }

  // src/shared/api.ts
  var COMPANION_BASE_URL = "https://localhost:48766";
  async function request(path, body) {
    const token = getStoredToken();
    return fetch(`${COMPANION_BASE_URL}${path}`, {
      method: body === void 0 ? "GET" : "POST",
      headers: {
        "Content-Type": "application/json",
        ...token ? { "X-Favorik-Token": token } : {}
      },
      body: body === void 0 ? void 0 : JSON.stringify(body)
    });
  }
  async function callWithPairing(path, body) {
    if (!getStoredToken()) {
      await showPairingDialog();
    }
    let response = await request(path, body);
    if (response.status === 401) {
      await showPairingDialog();
      response = await request(path, body);
    }
    return response;
  }
  async function parseApiResult(response) {
    if (response.status === 401) {
      return { success: false, message: "Pairing with Favorik Companion was not completed." };
    }
    if (response.status === 403) {
      return { success: false, message: "Favorik Companion rejected this request's origin." };
    }
    if (!response.ok) {
      return { success: false, message: `Favorik Companion returned an unexpected error (${response.status}).` };
    }
    return await response.json();
  }
  async function saveAttachments(subject, receivedTime, attachments) {
    const response = await callWithPairing("/api/save-attachments", { subject, receivedTime, attachments });
    return parseApiResult(response);
  }
  async function saveText(subject, receivedTime) {
    const response = await callWithPairing("/api/save-text", { subject, receivedTime });
    return parseApiResult(response);
  }
  async function openOptions() {
    const response = await callWithPairing("/api/open-options");
    return parseApiResult(response);
  }

  // src/shared/notice-dialog.ts
  function showNotice(message) {
    return new Promise((resolve) => {
      const url = `https://localhost:48766/notice.html?message=${encodeURIComponent(message)}`;
      Office.context.ui.displayDialogAsync(url, { height: 20, width: 25, displayInIframe: false }, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          resolve();
          return;
        }
        const dialog = asyncResult.value;
        let settled = false;
        const finish = () => {
          if (settled) {
            return;
          }
          settled = true;
          dialog.close();
          resolve();
        };
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, finish);
        dialog.addEventHandler(Office.EventType.DialogEventReceived, finish);
      });
    });
  }

  // src/commands/commands.ts
  Office.onReady();
  function showHint(item, message) {
    try {
      item?.notificationMessages?.replaceAsync("favorik-hint", {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message,
        icon: "icon-attachment-16",
        persistent: false
      });
    } catch {
    }
  }
  function getAttachmentDetailsAsync(item) {
    return new Promise((resolve, reject) => {
      item.getAttachmentsAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(result.error);
        }
      });
    });
  }
  function getAttachmentContentAsync(item, id) {
    return new Promise((resolve, reject) => {
      item.getAttachmentContentAsync(id, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(result.error);
        }
      });
    });
  }
  async function getNonInlineAttachmentsAsBase64(item) {
    const details = await getAttachmentDetailsAsync(item);
    const nonInline = details.filter((a) => !a.isInline);
    const payloads = [];
    for (const attachment of nonInline) {
      const content = await getAttachmentContentAsync(item, attachment.id);
      if (content.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
        payloads.push({ name: attachment.name, isInline: attachment.isInline, contentBase64: content.content });
      }
    }
    return payloads;
  }
  function getReceivedTimeIso(item) {
    const created = item?.dateTimeCreated;
    return created instanceof Date ? created.toISOString() : null;
  }
  Office.actions.associate("favorikSave", async (event) => {
    try {
      const item = Office.context.mailbox.item;
      if (!item) {
        await showNotice("Please select only one email.");
        event.completed();
        return;
      }
      const attachments = await getNonInlineAttachmentsAsBase64(item);
      const result = await saveAttachments(item.subject ?? null, getReceivedTimeIso(item), attachments);
      if (!result.success) {
        showHint(item, result.message);
      }
    } catch (error) {
      showHint(Office.context.mailbox.item, `Favorik: ${String(error)}`);
    } finally {
      event.completed();
    }
  });
  Office.actions.associate("saveText", async (event) => {
    try {
      const item = Office.context.mailbox.item;
      showHint(item, "Press Ctrl+C on your selection first, then continue.");
      await saveText(item?.subject ?? null, getReceivedTimeIso(item));
    } catch (error) {
      showHint(Office.context.mailbox.item, `Favorik: ${String(error)}`);
    } finally {
      event.completed();
    }
  });
  Office.actions.associate("openOptions", async (event) => {
    try {
      await openOptions();
    } catch (error) {
      showHint(Office.context.mailbox.item, `Favorik: ${String(error)}`);
    } finally {
      event.completed();
    }
  });
})();
//# sourceMappingURL=commands.js.map
