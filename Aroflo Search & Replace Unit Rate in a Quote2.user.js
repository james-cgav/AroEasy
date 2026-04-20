
// ==UserScript==
// @name         Aroflo Search & Replace Unit Rate in a Quote
// @namespace    https://contactgroup.com.au
// @version      2026-04-21
// @description  Click Ctrl-Shift-U to replace the unit rate of matching quote items while in an Aroflo quote.
// @author       Andrew Otley & Guray Sunamak
// @match        https://office.aroflo.com/ims/Site/Service/Quotes/index.cfm*
// @icon         https://www.google.com/s2/favicons?sz=64&domain=aroflo.com
// @grant        none
// @license MIT
// ==/UserScript==

(function() {
    'use strict';

    function runQuoteUnitRateReplace() {
        const searchValue = prompt(
            "Quote Item Search",
            "Enter Part Number OR Item Description"
        );
        if (!searchValue) return;
        const replacementValue = prompt(
            "Quote Item Unit Rate Replace",
            "Enter New Unit Rate"
        );

        if (replacementValue === null || replacementValue.trim() === '') return;
        if (isNaN(replacementValue)) {
            alert("The new unit rate must be a number.");
            return;
        }

        let replaceCount = 0;
        const partList = document.querySelectorAll("input[name='partNo']");
        const descList = document.querySelectorAll("textarea[name='item']");
        for (let i = 0; i < partList.length; i++) {
            const partNo = (partList[i]?.value || '').trim().toLowerCase();
            const itemDesc = (descList[i]?.value || '').trim().toLowerCase();
            const searchText = searchValue.trim().toLowerCase();
            let unitRateId = null;
            if (partNo === searchText) {
                unitRateId = partList[i].id.replace("PartNo", "unitrate");
            } else if (itemDesc === searchText) {
                unitRateId = descList[i].id.replace("ItemText", "unitrate");
            }
            if (!unitRateId) continue;
            const unitRateInput = document.querySelector(`input[id='${unitRateId}']`);
            if (!unitRateInput) continue;
            unitRateInput.value = replacementValue;
            unitRateInput.dispatchEvent(new Event('input', { bubbles: true }));
            unitRateInput.dispatchEvent(new Event('change', { bubbles: true }));
            replaceCount++;
        }
        let outMsg = `Replaced ${replaceCount} occurrence(s).`;
        if (replaceCount > 0) outMsg += " Don't forget to click Save.";
        alert(outMsg);
    }
    document.addEventListener('keydown', function(event) {
        // Ctrl + Shift + U
        if (event.ctrlKey && event.shiftKey && event.code === 'KeyU') {
            event.preventDefault();
            runQuoteUnitRateReplace();
        }
    });
})();
