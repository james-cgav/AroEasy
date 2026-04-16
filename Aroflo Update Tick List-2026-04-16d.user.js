// ==UserScript==
// @name         Aroflo Update Tick List
// @namespace    https://contactgroup.com.au
// @version      2026-04-16d
// @description  Click Ctrl-Shift-Q or Ctrl-Shit-E and follow the instructions to replace quote item values while in an Aroflo quote.
// @author       Guray Sunamak & James Wright
// @match        https://office.aroflo.com/ims/Site/Service/workrequest/index.cfm*
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


    function runSetTickList(target) {
        if (target) {
            //var partno = null;
            //var targetid = null;
            //var targetinput = null;
            //var replacecnt = 0;
            var state = null;
            //let rlist = document.querySelectorAll('label.afRadio__label[for^="NA_"]');
            let rlist = document.querySelectorAll('input.afRadio__input[id^="' + target + '_"]');
            //let dlist = document.querySelectorAll("textarea[name='item']");
            for (let i=0; i<rlist.length; ++i) {
/*                if((plist[i].value).toLowerCase() == searep.toLowerCase()) {
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
*/
                /*if (rlist[i].checked) {
                    state = "Checked";
                }
                else
                {
                    state = "Un-checked";
                }*/
                //var outmsg = "Found " + rlist[i].for + " with value of " + state;
                //if (replacecnt>0) outmsg += " Don't forget to click Save";
                //alert(outmsg);
                rlist[i].checked = true;
            }
            //var outmsg = "Replaced " + replacecnt + " occurences.";
            //if (replacecnt>0) outmsg += " Don't forget to click Save";
            //alert(outmsg);
            alert("end");
        }
    }
/*
    function runQuoteSearchReplace() {
        var searep = prompt("Takeoff Sheet Item Search", "Enter Part Number OR Item Description");
        if (searep) {
            var repval = prompt("Takeoff Sheet Item Replace", "Enter New Price");
            if (!isNaN(repval)) {
                var partno = null;
                var costexid = null;
                var costexinput = null;
                var replacecnt = 0;
                let plist = document.querySelectorAll("input[name='partNo']");
                let dlist = document.querySelectorAll("textarea[name='item']");
                for (let i=0; i<plist.length; ++i) {
                    if((plist[i].value).toLowerCase() == searep.toLowerCase()) {
                        costexid = plist[i].id.replace("PartNo", "costex");
                        let costexinput = document.querySelector("input[id='" + costexid + "']");
                        costexinput.value = repval;
                        costexinput.dispatchEvent(new Event('change', { 'bubbles': true }));
                        replacecnt++;
                    } else if((dlist[i].value).toLowerCase() == searep.toLowerCase()) {
                        costexid = dlist[i].id.replace("ItemText", "costex");
                        let costexinput = document.querySelector("input[id='" + costexid + "']");
                        costexinput.value = repval;
                        costexinput.dispatchEvent(new Event('change', { 'bubbles': true }));
                        replacecnt++;
                    }
                }
                var outmsg = "Replaced " + replacecnt + " occurences.";
                if (replacecnt>0) outmsg += " Don't forget to click Save";
                alert(outmsg);
            }
        }
    }

    function runQuoteSearchReplace_unitrate() {
        var searep = prompt("Takeoff Sheet Item Search", "Enter Part Number OR Item Description");
        if (searep) {
            var repval = prompt("Takeoff Sheet Item Replace", "Enter New Hours per Item");
            if (!isNaN(repval)) {
                var partno = null;
                var costexid = null;
                var costexinput = null;
                var replacecnt = 0;
                let plist = document.querySelectorAll("input[name='partNo']");
                let dlist = document.querySelectorAll("textarea[name='item']");
                for (let i=0; i<plist.length; ++i) {
                    if((plist[i].value).toLowerCase() == searep.toLowerCase()) {
                        costexid = plist[i].id.replace("PartNo", "unitrate");
                        let costexinput = document.querySelector("input[id='" + costexid + "']");
                        costexinput.value = repval;
                        costexinput.dispatchEvent(new Event('change', { 'bubbles': true }));
                        replacecnt++;
                    } else if((dlist[i].value).toLowerCase() == searep.toLowerCase()) {
                        costexid = dlist[i].id.replace("ItemText", "unitrate");
                        let costexinput = document.querySelector("input[id='" + costexid + "']");
                        costexinput.value = repval;
                        costexinput.dispatchEvent(new Event('change', { 'bubbles': true }));
                        replacecnt++;
                    }
                }
                var outmsg = "Replaced " + replacecnt + " occurences.";
                if (replacecnt>0) outmsg += " Don't forget to click Save";
                alert(outmsg);
            }
        }
    }

*/
    // Add event listener for the keyboard shortcut
    document.addEventListener('keydown', function(event) {
        // Check if the key combination is Ctrl+Shift+Q
        if (event.ctrlKey && event.shiftKey && event.code === 'KeyQ') {
            event.preventDefault();
            //runQuoteSearchReplace();
            //runQuoteSearchReplaceGeneric("costex", "Price")
            runSetTickList("NA");
        }
        if (event.ctrlKey && event.shiftKey && event.code === 'KeyE') {
            event.preventDefault();
            //runQuoteSearchReplace_unitrate();
            //runQuoteSearchReplaceGeneric("unitrate", "UR")
            runSetTickList("Yes");
        }
    });

})();
