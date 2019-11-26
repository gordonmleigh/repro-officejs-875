/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Office, window */

Office.onReady(info => {
  if (info.host !== Office.HostType.PowerPoint) {
    return;
  }

  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("run").onclick = run;
  document.getElementById("stop").onclick = stop;
});

let intervalHandle;

async function run() {
  if (intervalHandle) {
    return;
  }
  intervalHandle = window.setInterval(() => {
    Office.context.document.goToByIdAsync(Office.Index.Next, Office.GoToType.Index, e =>
      console.log(`goToByIdAsync result`, e)
    );
  }, 1000);
}

async function stop() {
  window.clearInterval(intervalHandle);
  intervalHandle = undefined;
}
