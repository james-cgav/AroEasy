// ==UserScript==
// @name         Aroflo Search & Replace in a Quote [Flexible Field Update]
// @namespace    https://contactgroup.com.au
// @version      2026-04-22e
// @description  Ctrl-Shift-Q updates quote item fields within a selected takeoff sheet or first level assembly.
// @author       James Wright, Andrew Otley & Guray Sunamak
// @match        https://office.aroflo.com/ims/Site/Service/Quotes/index.cfm*
// @icon         https://www.google.com/s2/favicons?sz=64&domain=aroflo.com
// @grant        none
// @license MIT
// ==/UserScript==

(function () {
    'use strict';

    const UPDATE_FIELDS = {
        costex: {
            label: 'Price',
            selector: "input[name='costex']",
            numeric: true
        },
        unitLab: {
            label: 'Unit Rate',
            selector: "input[name='unitLab']",
            numeric: true
        },
        markup: {
            label: 'Markup',
            selector: "input[name='markup']",
            numeric: true
        },
        item: {
            label: 'Description',
            selector: "textarea[name='item']",
            numeric: false
        },
        partNo: {
            label: 'Part No',
            selector: "input[name='partNo']",
            numeric: false
        }
    };

    function escapeHtml(str) {
        return String(str ?? '').replace(/[&<>"']/g, function (m) {
            return ({
                '&': '&amp;',
                '<': '&lt;',
                '>': '&gt;',
                '"': '&quot;',
                "'": '&#39;'
            })[m];
        });
    }

    function getTakeoffName(takeoffLi) {
        const input = takeoffLi.querySelector('textarea[name="to_takeoff_name"]');
        return input ? input.value.trim() : `Takeoff ${takeoffLi.dataset.takeoffId || ''}`;
    }

    function getLevel1Assemblies(takeoffLi) {
        const container = takeoffLi.querySelector('.takeoffContainer');
        if (!container) return [];

        return Array.from(container.querySelectorAll('li.qteLnItem[data-level-no="1"]')).map(li => {
            const nameEl = li.querySelector('textarea[name="item"]');
            const name = nameEl ? nameEl.value.trim() : `Assembly ${li.dataset.lineItemId || ''}`;

            return {
                type: 'assembly',
                label: `${getTakeoffName(takeoffLi)} > ${name}`,
                element: li
            };
        });
    }

    function getSelectableScopes() {
        const takeoffLis = Array.from(document.querySelectorAll('#totopLevel > li[data-takeoff-id]'));
        const scopes = [];

        takeoffLis.forEach(takeoffLi => {
            const takeoffName = getTakeoffName(takeoffLi);

            scopes.push({
                type: 'takeoff',
                label: `[Takeoff] ${takeoffName}`,
                element: takeoffLi.querySelector('.takeoffContainer') || takeoffLi
            });

            scopes.push(...getLevel1Assemblies(takeoffLi));
        });

        return scopes;
    }

    function getLineItemsInScope(scopeElement) {
        if (!scopeElement) return [];

        if (scopeElement.matches('li.qteLnItem')) {
            return [scopeElement, ...scopeElement.querySelectorAll('li.qteLnItem')];
        }

        return Array.from(scopeElement.querySelectorAll('li.qteLnItem'));
    }

    function getFieldValueFromLineItem(lineItem, fieldKey) {
        const cfg = UPDATE_FIELDS[fieldKey];
        if (!cfg) return '';

        const el = lineItem.querySelector(cfg.selector);
        return el ? (el.value || '').trim() : '';
    }

    function getSearchableItemsForScope(scopeElement, fieldKey) {
        const lineItems = getLineItemsInScope(scopeElement);
        const map = new Map();

        lineItems.forEach(lineItem => {
            const partInput = lineItem.querySelector("input[name='partNo']");
            const descInput = lineItem.querySelector("textarea[name='item']");

            const partNo = (partInput?.value || '').trim();
            const desc = (descInput?.value || '').trim();

            if (!partNo && !desc) return;

            const matchKey = `${partNo.toLowerCase()}|||${desc.toLowerCase()}`;

            if (!map.has(matchKey)) {
                map.set(matchKey, {
                    key: matchKey,
                    partNo,
                    desc,
                    fieldValue: getFieldValueFromLineItem(lineItem, fieldKey),
                    count: 1
                });
            } else {
                map.get(matchKey).count++;
            }
        });

        return Array.from(map.values()).sort((a, b) => {
            const aText = `${a.partNo} ${a.desc}`.trim().toLowerCase();
            const bText = `${b.partNo} ${b.desc}`.trim().toLowerCase();
            return aText.localeCompare(bText);
        });
    }

    function renderItemTable(listEl, items, query, selectedKey, onSelect) {
        const q = query.trim().toLowerCase();

        const filtered = !q
            ? items
            : items.filter(item => {
                const haystack = `${item.partNo} ${item.desc} ${item.fieldValue} ${item.count}`.toLowerCase();
                return haystack.includes(q);
            });

        listEl.innerHTML = `
            <div style="display:grid;grid-template-columns:180px 1fr 180px 90px;font-weight:bold;border-bottom:1px solid #ccc;background:#f7f7f7;">
                <div style="padding:8px;">Part No</div>
                <div style="padding:8px;">Description</div>
                <div style="padding:8px;">Current Value</div>
                <div style="padding:8px;text-align:right;">Matches</div>
            </div>
            <div>
                ${filtered.map(item => `
                    <div
                        class="tm-item-row"
                        data-key="${escapeHtml(item.key)}"
                        style="
                            display:grid;
                            grid-template-columns:180px 1fr 180px 90px;
                            border-bottom:1px solid #eee;
                            cursor:pointer;
                            background:${item.key === selectedKey ? '#eef5ff' : '#fff'};
                        "
                    >
                        <div style="padding:8px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${escapeHtml(item.partNo)}</div>
                        <div style="padding:8px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${escapeHtml(item.desc)}</div>
                        <div style="padding:8px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${escapeHtml(item.fieldValue)}</div>
                        <div style="padding:8px;text-align:right;">${escapeHtml(item.count)}</div>
                    </div>
                `).join('')}
            </div>
        `;

        Array.from(listEl.querySelectorAll('.tm-item-row')).forEach(el => {
            el.addEventListener('click', () => onSelect(el.dataset.key));
        });
    }

    function showMessageModal(title, message, isSuccess = true) {
        return new Promise((resolve) => {
            const overlay = document.createElement('div');
            overlay.style.cssText = `
                position: fixed;
                inset: 0;
                background: rgba(0,0,0,0.45);
                z-index: 999999;
                display: flex;
                align-items: center;
                justify-content: center;
            `;

            const modal = document.createElement('div');
            modal.style.cssText = `
                background: #fff;
                width: min(560px, 92vw);
                border-radius: 10px;
                box-shadow: 0 10px 30px rgba(0,0,0,0.25);
                padding: 18px;
                font-family: Arial, sans-serif;
            `;

            modal.innerHTML = `
                <div style="font-size:18px;font-weight:bold;margin-bottom:12px;color:${isSuccess ? '#1b5e20' : '#8a1c1c'};">
                    ${escapeHtml(title)}
                </div>
                <div style="white-space:pre-line;line-height:1.5;margin-bottom:18px;">
                    ${escapeHtml(message)}
                </div>
                <div style="display:flex;justify-content:flex-end;">
                    <button id="tm_msg_ok" type="button" style="padding:8px 14px;">OK</button>
                </div>
            `;

            overlay.appendChild(modal);
            document.body.appendChild(overlay);

            function close() {
                overlay.remove();
                resolve();
            }

            modal.querySelector('#tm_msg_ok').addEventListener('click', close);
            overlay.addEventListener('click', (e) => {
                if (e.target === overlay) close();
            });
        });
    }

    function showReplaceModal(scopes) {
        return new Promise((resolve) => {
            if (!scopes.length) {
                showMessageModal('No scopes found', 'No takeoff sheets or level 1 assemblies were found on this quote.', false);
                resolve(null);
                return;
            }

            let selectedScope = scopes[0];
            let scopeQuery = '';
            let itemQuery = '';
            let selectedItemKey = '';
            let selectedField = 'costex';
            let currentItems = getSearchableItemsForScope(selectedScope.element, selectedField);

            const overlay = document.createElement('div');
            overlay.style.cssText = `
                position: fixed;
                inset: 0;
                background: rgba(0,0,0,0.45);
                z-index: 999999;
                display: flex;
                align-items: center;
                justify-content: center;
            `;

            const modal = document.createElement('div');
            modal.style.cssText = `
                background: #fff;
                width: min(1100px, 97vw);
                max-height: 92vh;
                border-radius: 10px;
                box-shadow: 0 10px 30px rgba(0,0,0,0.25);
                padding: 18px;
                font-family: Arial, sans-serif;
                display: flex;
                flex-direction: column;
            `;

            modal.innerHTML = `
                <div style="font-size:18px;font-weight:bold;margin-bottom:14px;">
                    Update Quote Item
                </div>

                <div style="display:flex;gap:18px;align-items:flex-end;flex-wrap:wrap;margin-bottom:16px;">
                    <div>
                        <label style="font-weight:bold;display:block;margin-bottom:6px;">1. What to update</label>
                        <select id="tm_field_select" style="padding:8px;min-width:180px;">
                            ${Object.entries(UPDATE_FIELDS).map(([key, cfg]) => `
                                <option value="${escapeHtml(key)}">${escapeHtml(cfg.label)}</option>
                            `).join('')}
                        </select>
                    </div>

                    <div>
                        <label id="tm_new_value_label" style="font-weight:bold;display:block;margin-bottom:6px;">2. New Value</label>
                        <input id="tm_new_value" type="text" placeholder="Enter new value" style="padding:8px;min-width:260px;">
                    </div>
                </div>

                <div style="display:grid;grid-template-columns:320px 1fr;gap:16px;min-height:420px;">
                    <div style="display:flex;flex-direction:column;min-height:0;">
                        <label style="font-weight:bold;margin-bottom:6px;">3. Select Takeoff Sheet / Level 1 Assembly</label>
                        <input id="tm_scope_search" type="text" placeholder="Search scope..." style="padding:8px;margin-bottom:8px;">
                        <div id="tm_scope_list" style="border:1px solid #ccc;border-radius:6px;overflow:auto;min-height:300px;flex:1;"></div>
                    </div>

                    <div style="display:flex;flex-direction:column;min-height:0;">
                        <label style="font-weight:bold;margin-bottom:6px;">4. Select Part Number / Description</label>
                        <input id="tm_item_search" type="text" placeholder="Search part number, description, current value, or match count..." style="padding:8px;margin-bottom:8px;">
                        <div id="tm_item_list" style="border:1px solid #ccc;border-radius:6px;overflow:auto;min-height:300px;flex:1;"></div>
                    </div>
                </div>

                <div id="tm_summary" style="margin-top:12px;color:#444;font-size:13px;"></div>

                <div style="display:flex;gap:10px;justify-content:flex-end;margin-top:18px;">
                    <button id="tm_cancel" type="button" style="padding:8px 14px;">Cancel</button>
                    <button id="tm_ok" type="button" style="padding:8px 14px;">Apply</button>
                </div>
            `;

            overlay.appendChild(modal);
            document.body.appendChild(overlay);

            const scopeSearch = modal.querySelector('#tm_scope_search');
            const scopeList = modal.querySelector('#tm_scope_list');
            const itemSearch = modal.querySelector('#tm_item_search');
            const itemList = modal.querySelector('#tm_item_list');
            const fieldSelect = modal.querySelector('#tm_field_select');
            const newValueInput = modal.querySelector('#tm_new_value');
            const newValueLabel = modal.querySelector('#tm_new_value_label');
            const summary = modal.querySelector('#tm_summary');

            function updateNewValueLabel() {
                const fieldLabel = UPDATE_FIELDS[selectedField].label;
                newValueLabel.textContent = `2. New ${fieldLabel}`;
                newValueInput.placeholder = `Enter new ${fieldLabel}`;
            }

            function updateSummary() {
                const selectedItem = currentItems.find(i => i.key === selectedItemKey);
                summary.innerHTML = `
                    <div><strong>Scope:</strong> ${escapeHtml(selectedScope?.label || 'None')}</div>
                    <div><strong>Item:</strong> ${escapeHtml(selectedItem ? `${selectedItem.partNo} — ${selectedItem.desc}` : 'None')}</div>
                    <div><strong>Field:</strong> ${escapeHtml(UPDATE_FIELDS[selectedField]?.label || 'None')}</div>
                    <div><strong>Current value:</strong> ${escapeHtml(selectedItem?.fieldValue || '')}</div>
                    <div><strong>Matching rows in scope:</strong> ${escapeHtml(selectedItem?.count || '')}</div>
                `;
            }

            function renderScopes() {
                const q = scopeQuery.trim().toLowerCase();
                const filtered = !q
                    ? scopes
                    : scopes.filter(s => s.label.toLowerCase().includes(q));

                scopeList.innerHTML = filtered.map(scope => `
                    <div
                        class="tm-scope-option"
                        data-scope-index="${scopes.indexOf(scope)}"
                        style="
                            padding: 8px 10px;
                            cursor: pointer;
                            border-bottom: 1px solid #eee;
                            background: ${scope.label === selectedScope.label ? '#eef5ff' : '#fff'};
                        "
                    >
                        ${escapeHtml(scope.label)}
                    </div>
                `).join('');

                Array.from(scopeList.querySelectorAll('.tm-scope-option')).forEach(el => {
                    el.addEventListener('click', () => {
                        selectedScope = scopes[Number(el.dataset.scopeIndex)];
                        currentItems = getSearchableItemsForScope(selectedScope.element, selectedField);
                        selectedItemKey = '';
                        itemQuery = '';
                        itemSearch.value = '';
                        renderScopes();
                        renderItems();
                        updateSummary();
                    });
                });
            }

            function renderItems() {
                renderItemTable(itemList, currentItems, itemQuery, selectedItemKey, (key) => {
                    selectedItemKey = key;
                    renderItems();
                    updateSummary();
                });
            }

            scopeSearch.addEventListener('input', () => {
                scopeQuery = scopeSearch.value;
                renderScopes();
            });

            itemSearch.addEventListener('input', () => {
                itemQuery = itemSearch.value;
                renderItems();
            });

            fieldSelect.addEventListener('change', () => {
                selectedField = fieldSelect.value;
                currentItems = getSearchableItemsForScope(selectedScope.element, selectedField);
                selectedItemKey = '';
                updateNewValueLabel();
                renderItems();
                updateSummary();
            });

            modal.querySelector('#tm_cancel').addEventListener('click', () => {
                overlay.remove();
                resolve(null);
            });

            modal.querySelector('#tm_ok').addEventListener('click', () => {
                const selectedItem = currentItems.find(i => i.key === selectedItemKey);
                const newValue = newValueInput.value.trim();
                const fieldCfg = UPDATE_FIELDS[selectedField];

                if (!selectedScope) {
                    showMessageModal('Missing selection', 'Please select a Takeoff Sheet or Level 1 Assembly.', false);
                    return;
                }

                if (!selectedItem) {
                    showMessageModal('Missing selection', 'Please select a part number / description.', false);
                    return;
                }

                if (!selectedField || !fieldCfg) {
                    showMessageModal('Missing selection', 'Please select a field to update.', false);
                    return;
                }

                if (newValue === '') {
                    showMessageModal('Missing value', 'Please enter a new value.', false);
                    return;
                }

                if (fieldCfg.numeric && isNaN(newValue)) {
                    showMessageModal('Invalid value', `${fieldCfg.label} must be numeric.`, false);
                    return;
                }

                overlay.remove();
                resolve({
                    scope: selectedScope,
                    item: selectedItem,
                    fieldKey: selectedField,
                    newValue
                });
            });

            overlay.addEventListener('click', (e) => {
                if (e.target === overlay) {
                    overlay.remove();
                    resolve(null);
                }
            });

            renderScopes();
            renderItems();
            updateNewValueLabel();
            updateSummary();
            newValueInput.focus();
        });
    }

    async function runQuoteFieldUpdate() {
        const scopes = getSelectableScopes();
        const result = await showReplaceModal(scopes);
        if (!result) return;

        const { scope, item, fieldKey, newValue } = result;
        const fieldCfg = UPDATE_FIELDS[fieldKey];
        const lineItems = getLineItemsInScope(scope.element);

        let updateCount = 0;
        const oldValues = new Set();

        lineItems.forEach(lineItem => {
            const partInput = lineItem.querySelector("input[name='partNo']");
            const descInput = lineItem.querySelector("textarea[name='item']");

            const partNo = (partInput?.value || '').trim();
            const desc = (descInput?.value || '').trim();

            const isMatch =
                partNo.toLowerCase() === item.partNo.toLowerCase() &&
                desc.toLowerCase() === item.desc.toLowerCase();

            if (!isMatch) return;

            const targetEl = lineItem.querySelector(fieldCfg.selector);
            if (!targetEl) return;

            oldValues.add((targetEl.value || '').trim());

            targetEl.value = newValue;
            targetEl.dispatchEvent(new Event('input', { bubbles: true }));
            targetEl.dispatchEvent(new Event('change', { bubbles: true }));

            updateCount++;
        });

        if (updateCount > 0) {
            await showMessageModal(
                'Update complete',
                `Updated ${updateCount} occurrence${updateCount === 1 ? '' : 's'}.\n\n` +
                `Scope: ${scope.label}\n` +
                `Field: ${fieldCfg.label}\n` +
                `Old value${oldValues.size === 1 ? '' : 's'}: ${Array.from(oldValues).join(', ')}\n` +
                `New value: ${newValue}\n\n` +
                `Don't forget to click Save.`,
                true
            );
        } else {
            await showMessageModal(
                'Update failed',
                `No matching editable fields were found for:\n\n` +
                `Scope: ${scope.label}\n` +
                `Field: ${fieldCfg.label}`,
                false
            );
        }
    }

    document.addEventListener('keydown', function (event) {
        if (event.ctrlKey && event.shiftKey && event.code === 'KeyQ') {
            event.preventDefault();
            runQuoteFieldUpdate();
        }
    });

})();
