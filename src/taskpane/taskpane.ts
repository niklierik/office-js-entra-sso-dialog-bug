/* eslint-disable no-undef */
import $ from "jquery";

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady(() => {
  $("#open-dialog").on("click", () => {});
});

export async function run() {}

async function openDialog() {
  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      window.location.origin + "/dialog.html",
      {
        width: 40,
        height: 40,
        displayInIframe: false,
      },
      (result) => {
        console.log(result);
      }
    );
  });
}
