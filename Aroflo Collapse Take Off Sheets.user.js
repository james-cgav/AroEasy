// ==UserScript==
// @name         Aroflo Collapse Take Off Sheets
// @namespace    https://contactgroup.com.au
// @version      2026-05-01a
// @description  Ctrl-Shift-Z collapses all take off sheets and sub assemblies.
// @author       James Wright
// @match        https://office.aroflo.com/ims/Site/Service/Quotes/index.cfm*
// @icon         https://www.google.com/s2/favicons?sz=64&domain=aroflo.com
// @grant        none
// @license      MIT
// ==/UserScript==

(function() {
    'use strict';

    function collapseAllTakeoffSheets() {
        // Collapse any open folders/assemblies via Aroflo's own click handlers.
        const openFolders = document.querySelectorAll(
            '.expandFolder[data-folder-open="1"]'
        );

        openFolders.forEach(folder => {
            folder.click();
        });

        // Fallback: force-hide visible takeoff and child containers if any missed the click handler.
        document.querySelectorAll('.takeoffContainer, ul.ul-child').forEach(container => {
            container.style.display = 'none';
        });

        // Fallback: reset folder state attributes.
        document.querySelectorAll('.expandFolder').forEach(folder => {
            folder.setAttribute('data-folder-open', '0');
        });
    }

    document.addEventListener('keydown', function(event) {
        if (event.ctrlKey && event.shiftKey && event.code === 'KeyZ') {
            event.preventDefault();
            collapseAllTakeoffSheets();
        }
    });
})();
