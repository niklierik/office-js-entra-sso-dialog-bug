/* eslint-disable no-undef */
import $ from "jquery";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady(() => {
  console.log("Opened addin");
  $(".dialog-closed-text").hide();
  $("#open-dialog").removeAttr("disabled");
  $("#open-dialog").on("click", async () => {
    const dialog = await openDialog();

    console.log("Opened dialog");

    dialog.addEventHandler(Office.EventType.DialogMessageReceived, (message) => {
      console.log("Dialog message", message);
    });

    dialog.addEventHandler(Office.EventType.DialogEventReceived, (message) => {
      console.log("Dialog event", message);

      if (message["error"] === 12006) {
        $(".dialog-closed-text").show();
      }
    });

    console.log("Subscribed to dialog events.");
  });
});

export async function run() {}

async function openDialog() {
  return new Promise<Office.Dialog>((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      window.location.origin + "/test/dialog-redirect.html",
      {
        width: 40,
        height: 40,
        displayInIframe: false,
      },
      (result) => {
        if (result.status == Office.AsyncResultStatus.Failed) {
          reject(result);
          return;
        }

        resolve(result.value);
      }
    );
  });
}
