/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

(() => {
  // The initialize function must be run each time a new page is loaded
  Office.initialize = () => {
    const btnClose = document.getElementById('closeBtn');
    btnClose.style.display = "block";
    btnClose.onclick = function() {
        Office.context.ui.messageParent("closeDialog")
    };

    const btnSend = document.getElementById('sendBtn');
    btnSend.style.display = "block";
    btnSend.onclick = function() {
        var messageId = Math.floor(Math.random() * 1000);
        Office.context.ui.messageParent(messageId);
    };
  };

  // Add any ui-less function here
})();