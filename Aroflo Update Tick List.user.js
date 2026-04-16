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

    function runSetTickList(target) {
        if (target) {

            var state = null;
            let rlist = document.querySelectorAll('input.afRadio__input[id^="' + target + '_"]');
            for (let i=0; i<rlist.length; ++i) {
                rlist[i].checked = true;
            }
            //alert("end");
        }
    }

    // Add event listener for the keyboard shortcut
    document.addEventListener('keydown', function(event) {
        // Check if the key combination is Ctrl+Shift+Q
        if (event.ctrlKey && event.shiftKey && event.code === 'KeyQ') {
            event.preventDefault();
            runSetTickList("NA");
        }
        if (event.ctrlKey && event.shiftKey && event.code === 'KeyE') {
            event.preventDefault();
            runSetTickList("Yes");
        }
    });

})();
