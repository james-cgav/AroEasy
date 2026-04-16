// ==UserScript==
// @name         Aroflo Search & Replace in a Quote [Price and UR]
// @namespace    https://contactgroup.com.au
// @version      2025-02-20d
// @description  Click Ctrl-Shift-Q or Ctrl-Shit-E and follow the instructions to replace quote item values while in an Aroflo quote.
// @author       Guray Sunamak & James Wright
// @match        https://office.aroflo.com/ims/Site/Service/Quotes/index.cfm*
// @icon         https://www.google.com/s2/favicons?sz=64&domain=aroflo.com
// @grant        none
// @license MIT
// ==/UserScript==

(function() {
    'use strict';

    function runQuoteSearchReplaceGeneric(target, description) {
        var searep = prompt("Takeoff Sheet Item Search", "Enter Part Number OR Item Description");
        if (searep) {
            var repval = prompt("Takeoff Sheet Item Replace", "Enter New " + description);
            if (!isNaN(repval)) {
                var partno = null;
                var targetid = null;
                var targetinput = null;
                var replacecnt = 0;
                let plist = document.querySelectorAll("input[name='partNo']");
                let dlist = document.querySelectorAll("textarea[name='item']");
                for (let i=0; i<plist.length; ++i) {
                    if((plist[i].value).toLowerCase() == searep.toLowerCase()) {
                        targetid = plist[i].id.replace("PartNo", target);
                        let targetinput = document.querySelector("input[id='" + targetid + "']");
                        targetinput.value = repval;
                        targetinput.dispatchEvent(new Event('change', { 'bubbles': true }));
                        replacecnt++;
                    } else if((dlist[i].value).toLowerCase() == searep.toLowerCase()) {
                        targetid = dlist[i].id.replace("ItemText", target);
                        let targetinput = document.querySelector("input[id='" + targetid + "']");
                        targetinput.value = repval;
                        targetinput.dispatchEvent(new Event('change', { 'bubbles': true }));
                        replacecnt++;
                    }
                }
                var outmsg = "Replaced " + replacecnt + " occurences.";
                if (replacecnt>0) outmsg += " Don't forget to click Save";
                alert(outmsg);
            }
        }
    }

    // Add event listener for the keyboard shortcut
    document.addEventListener('keydown', function(event) {
        // Check if the key combination is Ctrl+Shift+Q
        if (event.ctrlKey && event.shiftKey && event.code === 'KeyQ') {
            event.preventDefault();
            runQuoteSearchReplaceGeneric("costex", "Price")
        }
        if (event.ctrlKey && event.shiftKey && event.code === 'KeyE') {
            event.preventDefault();
            runQuoteSearchReplaceGeneric("unitrate", "UR")
        }
    });

})();
