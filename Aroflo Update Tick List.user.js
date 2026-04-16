// ==UserScript==
// @name         Aroflo Update Tick List
// @namespace    https://contactgroup.com.au
// @version      2026-04-16d
// @description  Click Ctrl-Shift-Q (for NA) or Ctrl-Shit-E (for Yes) to update all tick list items to the chosen type.
// @author       James Wright
// @match        https://office.aroflo.com/ims/Site/Service/workrequest/index.cfm*
// @icon         https://www.google.com/s2/favicons?sz=64&domain=aroflo.com
// @grant        none
// @license MIT
// ==/UserScript==

(function() {
    'use strict';

    function runSetTickListSelective(target, override = false) {
        if (!target || !["Yes", "No", "NA"].includes(target)) return;

        const yesList = document.querySelectorAll('input.afRadio__input[id^="Yes_"]');
        const noList = document.querySelectorAll('input.afRadio__input[id^="No_"]');
        const naList = document.querySelectorAll('input.afRadio__input[id^="NA_"]');

        const targetMap = {
            Yes: yesList,
            No: noList,
            NA: naList
        };

        const selectedList = targetMap[target];

        for (let i = 0; i < yesList.length; i++) {
            const nothingChecked = !yesList[i].checked && !noList[i].checked && !naList[i].checked;

            if (override || nothingChecked) {
                selectedList[i].checked = true;
            }
        }
    }

    // Add event listener for the keyboard shortcut
    document.addEventListener('keydown', function(event) {
        // Check if the key combination is Ctrl+Shift+Q
        if (event.ctrlKey && event.shiftKey && event.code === 'KeyQ') {
            event.preventDefault();
            runSetTickListSelective("NA");
        }
        if (event.ctrlKey && event.shiftKey && event.code === 'KeyE') {
            event.preventDefault();
            runSetTickListSelective("Yes");
        }
    });

})();
