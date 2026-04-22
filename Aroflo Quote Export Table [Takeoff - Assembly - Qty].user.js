// ==UserScript==
// @name         Aroflo Quote Export Table [Takeoff / Assembly / Qty]
// @namespace    https://contactgroup.com.au
// @version      2026-04-22j
// @description  Ctrl-Shift-E exports quote line data into a modal table and clipboard-ready format.
// @author       James Wright
// @match        https://office.aroflo.com/ims/Site/Service/Quotes/index.cfm*
// @icon         https://www.google.com/s2/favicons?sz=64&domain=aroflo.com
// @grant        none
// @license MIT
// ==/UserScript==

(function () {
    'use strict';

    const OPTIONAL_COLUMNS_DEFAULT = {
        sortIndex: true,
        isAssembly: true,
        assemblyName: true
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

    function parseQty(value) {
        const n = parseFloat(String(value ?? '').replace(/,/g, '').trim());
        return Number.isFinite(n) ? n : 0;
    }

    function formatQty(value) {
        if (!Number.isFinite(value)) return '';
        if (Math.abs(value - Math.round(value)) < 0.0000001) return String(Math.round(value));
        return String(value);
    }

    function getTakeoffName(takeoffLi) {
        const input = takeoffLi.querySelector('textarea[name="to_takeoff_name"]');
        return input ? input.value.trim() : `Takeoff ${takeoffLi.dataset.takeoffId || ''}`;
    }

    function getLineData(lineItem) {
        const levelNo = parseInt(lineItem.dataset.levelNo || '0', 10);
        const lineItemId = lineItem.dataset.lineItemId || '';
        const parentId = lineItem.dataset.parentId || '';
        const partNo = (lineItem.querySelector("input[name='partNo']")?.value || '').trim();
        const description = (lineItem.querySelector("textarea[name='item']")?.value || '').trim();
        const qty = parseQty(lineItem.querySelector("input[name='qty']")?.value || '0');

        return {
            element: lineItem,
            lineItemId,
            parentId,
            levelNo,
            assemblyDepth: Math.max(levelNo - 1, 0),
            partNo,
            description,
            qty
        };
    }

    function buildParentItemPath(item, lineMap, takeoffName) {
        if (!item.parentId) {
            return `[${takeoffName}]`;
        }

        const parents = [];
        let current = lineMap.get(item.parentId);

        while (current) {
            parents.unshift(current.description || '');
            if (!current.parentId) break;
            current = lineMap.get(current.parentId);
        }

        const cleanParents = parents.filter(Boolean);
        if (!cleanParents.length) {
            return `[${takeoffName}]`;
        }

        return `[${takeoffName}] ${cleanParents.join(' > ')}`;
    }

    function buildAssemblyName(item, lineMap) {
        if (!item.parentId) return '';

        const parent = lineMap.get(item.parentId);
        return parent ? (parent.description || '') : '';
    }

    function collectTakeoffData() {
        const takeoffLis = Array.from(document.querySelectorAll('#totopLevel > li[data-takeoff-id]'));
        const rows = [];

        takeoffLis.forEach((takeoffLi, takeoffIdx) => {
            const takeoffId = takeoffLi.dataset.takeoffId || '';
            const takeoffName = getTakeoffName(takeoffLi);
            const lineItems = Array.from(takeoffLi.querySelectorAll('li.qteLnItem'));
            const lineMap = new Map();

            const parsedItems = lineItems.map(getLineData);
            parsedItems.forEach(item => lineMap.set(item.lineItemId, item));

            rows.push({
                rowKey: `takeoff_${takeoffId}`,
                rowType: 'takeoffHeader',
                takeoffId,
                groupKey: `takeoff_${takeoffId}`,
                parentGroupKey: null,
                parentItem: `[${takeoffName}]`,
                assemblyName: '',
                lineItemId: '',
                takeoffSheet: takeoffName,
                assemblyDepth: '',
                partNo: '',
                description: '',
                qty: '',
                extendedQty: '',
                isAssembly: 'No',
                sortIndex: `${String(takeoffIdx + 1).padStart(3, '0')}.0000`,
                isControllable: true
            });

            parsedItems.forEach((item, rowIdx) => {
                let multiplier = 1;
                let current = item;

                while (current && current.parentId) {
                    const parent = lineMap.get(current.parentId);
                    if (!parent) break;
                    multiplier *= parent.qty || 0;
                    current = parent;
                }

                const hasChildren = parsedItems.some(child => child.parentId === item.lineItemId);
                const rowType = hasChildren ? 'assemblyHeader' : 'item';

                let parentGroupKey = `takeoff_${takeoffId}`;
                let p = item.parentId ? lineMap.get(item.parentId) : null;

                while (p) {
                    const pHasChildren = parsedItems.some(child => child.parentId === p.lineItemId);
                    if (pHasChildren) {
                        parentGroupKey = `assembly_${takeoffId}_${p.lineItemId}`;
                        break;
                    }
                    p = p.parentId ? lineMap.get(p.parentId) : null;
                }

                const parentItem = buildParentItemPath(item, lineMap, takeoffName);
                const assemblyName = buildAssemblyName(item, lineMap);

                rows.push({
                    rowKey: `${rowType}_${takeoffId}_${item.lineItemId}`,
                    rowType,
                    takeoffId,
                    groupKey: rowType === 'assemblyHeader' ? `assembly_${takeoffId}_${item.lineItemId}` : null,
                    parentGroupKey,
                    parentItem,
                    assemblyName,
                    lineItemId: item.lineItemId,
                    takeoffSheet: takeoffName,
                    assemblyDepth: item.assemblyDepth,
                    partNo: item.partNo,
                    description: item.description,
                    qty: formatQty(item.qty),
                    extendedQty: formatQty(item.qty * multiplier),
                    isAssembly: rowType === 'assemblyHeader' ? 'Yes' : 'No',
                    sortIndex: `${String(takeoffIdx + 1).padStart(3, '0')}.${String(rowIdx + 1).padStart(4, '0')}`,
                    isControllable: rowType === 'assemblyHeader'
                });
            });
        });

        return rows;
    }

    function getEnabledOptionalHeaders(optionalColumns) {
        return [
            ...(optionalColumns.sortIndex ? ['Sort Index'] : []),
            ...(optionalColumns.isAssembly ? ['Is Assembly'] : []),
            ...(optionalColumns.assemblyName ? ['Assembly Name'] : [])
        ];
    }

    function getOptionalRowValues(row, optionalColumns) {
        return [
            ...(optionalColumns.sortIndex ? [row.sortIndex] : []),
            ...(optionalColumns.isAssembly ? [row.isAssembly] : []),
            ...(optionalColumns.assemblyName ? [row.assemblyName] : [])
        ];
    }

    function csvCell(value) {
        const str = String(value ?? '');
        return `"${str.replace(/"/g, '""')}"`;
    }

    function toCsv(rows, optionalColumns) {
        const headers = [
            ...getEnabledOptionalHeaders(optionalColumns),
            'Takeoff Sheet',
            'Assembly Depth',
            'Parent Item',
            'Part Number',
            'Description',
            'Qty',
            'Extended Qty'
        ];

        const lines = [headers.map(csvCell).join(',')];

        rows.forEach(row => {
            lines.push([
                ...getOptionalRowValues(row, optionalColumns),
                row.takeoffSheet,
                row.assemblyDepth,
                row.parentItem,
                row.partNo,
                row.description,
                row.qty,
                row.extendedQty
            ].map(csvCell).join(','));
        });

        return lines.join('\n');
    }

    function toPlainTsv(rows, optionalColumns) {
        const headers = [
            ...getEnabledOptionalHeaders(optionalColumns),
            'Takeoff Sheet',
            'Assembly Depth',
            'Parent Item',
            'Part Number',
            'Description',
            'Qty',
            'Extended Qty'
        ];

        const lines = [headers.join('\t')];

        rows.forEach(row => {
            lines.push([
                ...getOptionalRowValues(row, optionalColumns),
                row.takeoffSheet,
                row.assemblyDepth,
                row.parentItem,
                row.partNo,
                row.description,
                row.qty,
                row.extendedQty
            ].map(v => String(v ?? '').replace(/\t/g, ' ').replace(/\r?\n/g, ' ')).join('\t'));
        });

        return lines.join('\n');
    }

    function buildHtmlTable(rows, optionalColumns) {
        const headerCells = [
            ...getEnabledOptionalHeaders(optionalColumns),
            'Takeoff Sheet',
            'Assembly Depth',
            'Parent Item',
            'Part Number',
            'Description',
            'Qty',
            'Extended Qty'
        ].map(text => `<th>${escapeHtml(text)}</th>`).join('');

        const bodyRows = rows.map(row => {
            const isTakeoffHeader = row.rowType === 'takeoffHeader';
            const isAssemblyHeader = row.rowType === 'assemblyHeader';

            let bg = '#ffffff';
            let weight = 'normal';

            if (isTakeoffHeader) {
                bg = '#dfeaf7';
                weight = 'bold';
            } else if (isAssemblyHeader) {
                bg = '#f5f5f5';
                weight = 'bold';
            }

            return `
                <tr style="background:${bg};font-weight:${weight};">
                    ${optionalColumns.sortIndex ? `<td>${escapeHtml(row.sortIndex)}</td>` : ''}
                    ${optionalColumns.isAssembly ? `<td>${escapeHtml(row.isAssembly)}</td>` : ''}
                    ${optionalColumns.assemblyName ? `<td>${escapeHtml(row.assemblyName)}</td>` : ''}
                    <td>${escapeHtml(row.takeoffSheet)}</td>
                    <td>${escapeHtml(row.assemblyDepth)}</td>
                    <td>${escapeHtml(row.parentItem)}</td>
                    <td>${escapeHtml(row.partNo)}</td>
                    <td>${escapeHtml(row.description)}</td>
                    <td>${escapeHtml(row.qty)}</td>
                    <td>${escapeHtml(row.extendedQty)}</td>
                </tr>
            `;
        }).join('');

        return `
            <table border="1" style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:12px;width:100%;">
                <thead>
                    <tr>${headerCells}</tr>
                </thead>
                <tbody>
                    ${bodyRows}
                </tbody>
            </table>
        `;
    }

    function copyTableToClipboard(rows, optionalColumns) {
        const html = buildHtmlTable(rows, optionalColumns);
        const text = toPlainTsv(rows, optionalColumns);

        if (navigator.clipboard && window.ClipboardItem) {
            return navigator.clipboard.write([
                new ClipboardItem({
                    'text/html': new Blob([html], { type: 'text/html' }),
                    'text/plain': new Blob([text], { type: 'text/plain' })
                })
            ]).then(() => true).catch(() => false);
        }

        const wrapper = document.createElement('div');
        wrapper.contentEditable = 'true';
        wrapper.style.position = 'fixed';
        wrapper.style.left = '-9999px';
        wrapper.style.top = '0';
        wrapper.innerHTML = html;
        document.body.appendChild(wrapper);

        const range = document.createRange();
        range.selectNodeContents(wrapper);

        const sel = window.getSelection();
        sel.removeAllRanges();
        sel.addRange(range);

        let ok = false;
        try {
            ok = document.execCommand('copy');
        } catch (e) {
            ok = false;
        }

        sel.removeAllRanges();
        wrapper.remove();

        return Promise.resolve(ok);
    }

    function printHtmlTable(rows, optionalColumns) {
        const tableHtml = buildHtmlTable(rows, optionalColumns);

        const printWindow = window.open('', '_blank', 'width=1200,height=800');
        if (!printWindow) return false;

        printWindow.document.open();
        printWindow.document.write(`
            <!DOCTYPE html>
            <html>
            <head>
                <title>Quote Export</title>
                <style>
                    body {
                        font-family: Arial, sans-serif;
                        padding: 16px;
                    }
                    table {
                        border-collapse: collapse;
                        width: 100%;
                        font-size: 12px;
                    }
                    th, td {
                        border: 1px solid #999;
                        padding: 6px 8px;
                        text-align: left;
                        vertical-align: top;
                    }
                    th {
                        background: #f0f0f0;
                    }
                    @media print {
                        body {
                            padding: 0;
                        }
                    }
                </style>
            </head>
            <body>
                ${tableHtml}
            </body>
            </html>
        `);
        printWindow.document.close();

        printWindow.focus();
        setTimeout(() => {
            printWindow.print();
            printWindow.close();
        }, 250);

        return true;
    }

    function downloadFile(filename, content, mimeType) {
        const blob = new Blob([content], { type: mimeType });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
    }

    function getVisibleRows(allRows, includeTakeoffHeaders, includeAssemblyHeaders) {
        return allRows.filter(row => {
            if (row.rowType === 'takeoffHeader') return includeTakeoffHeaders;
            if (row.rowType === 'assemblyHeader') return includeAssemblyHeaders;
            return true;
        });
    }

    function buildChildrenMap(rows) {
        const childrenMap = {};
        rows.forEach(row => {
            if (row.parentGroupKey) {
                if (!childrenMap[row.parentGroupKey]) childrenMap[row.parentGroupKey] = [];
                if (row.groupKey) {
                    childrenMap[row.parentGroupKey].push(row.groupKey);
                }
            }
        });
        return childrenMap;
    }

    function setGroupStateRecursive(groupKey, value, childrenMap, groupState) {
        groupState[groupKey] = value;
        const children = childrenMap[groupKey] || [];
        children.forEach(childKey => setGroupStateRecursive(childKey, value, childrenMap, groupState));
    }

    function getExportRows(visibleRows, groupState) {
        const groupRowMap = new Map();
        visibleRows.forEach(r => {
            if (r.groupKey) groupRowMap.set(r.groupKey, r);
        });

        function areAncestorsEnabled(row) {
            let parent = row.parentGroupKey;
            while (parent) {
                if (!groupState[parent]) return false;
                const parentRow = groupRowMap.get(parent);
                parent = parentRow ? parentRow.parentGroupKey : null;
            }
            return true;
        }

        return visibleRows.filter(row => {
            if (row.rowType === 'takeoffHeader') {
                return !!groupState[row.groupKey];
            }

            if (row.rowType === 'assemblyHeader') {
                return !!groupState[row.groupKey] && areAncestorsEnabled(row);
            }

            return areAncestorsEnabled(row);
        });
    }

    function renderTableRows(rows, groupState, optionalColumns) {
        return rows.map(row => {
            const isTakeoffHeader = row.rowType === 'takeoffHeader';
            const isAssemblyHeader = row.rowType === 'assemblyHeader';

            let bg = '#fff';
            let weight = 'normal';

            if (isTakeoffHeader) {
                bg = '#dfeaf7';
                weight = 'bold';
            } else if (isAssemblyHeader) {
                bg = '#f5f5f5';
                weight = 'bold';
            }

            const checkboxHtml = row.isControllable
                ? `<input type="checkbox" class="tm-group-toggle" data-group-key="${escapeHtml(row.groupKey)}" ${groupState[row.groupKey] ? 'checked' : ''}>`
                : '';

            return `
                <tr style="background:${bg};font-weight:${weight};">
                    <td style="padding:6px 8px;border:1px solid #ddd;text-align:center;">${checkboxHtml}</td>
                    ${optionalColumns.sortIndex ? `<td style="padding:6px 8px;border:1px solid #ddd;">${escapeHtml(row.sortIndex)}</td>` : ''}
                    ${optionalColumns.isAssembly ? `<td style="padding:6px 8px;border:1px solid #ddd;">${escapeHtml(row.isAssembly)}</td>` : ''}
                    ${optionalColumns.assemblyName ? `<td style="padding:6px 8px;border:1px solid #ddd;">${escapeHtml(row.assemblyName)}</td>` : ''}
                    <td style="padding:6px 8px;border:1px solid #ddd;">${escapeHtml(row.takeoffSheet)}</td>
                    <td style="padding:6px 8px;border:1px solid #ddd;text-align:center;">${escapeHtml(row.assemblyDepth)}</td>
                    <td style="padding:6px 8px;border:1px solid #ddd;">${escapeHtml(row.parentItem)}</td>
                    <td style="padding:6px 8px;border:1px solid #ddd;">${escapeHtml(row.partNo)}</td>
                    <td style="padding:6px 8px;border:1px solid #ddd;">${escapeHtml(row.description)}</td>
                    <td style="padding:6px 8px;border:1px solid #ddd;text-align:right;">${escapeHtml(row.qty)}</td>
                    <td style="padding:6px 8px;border:1px solid #ddd;text-align:right;">${escapeHtml(row.extendedQty)}</td>
                </tr>
            `;
        }).join('');
    }

    function showExportModal(allRows) {
        return new Promise((resolve) => {
            let includeTakeoffHeaders = true;
            let includeAssemblyHeaders = true;
            const optionalColumns = { ...OPTIONAL_COLUMNS_DEFAULT };
            const groupState = {};

            allRows.forEach(row => {
                if (row.isControllable && row.groupKey) {
                    groupState[row.groupKey] = true;
                }
            });

            const childrenMap = buildChildrenMap(allRows);

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
                width: min(1700px, 99vw);
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
                    Export Quote Data
                </div>

                <div style="display:flex;gap:20px;align-items:center;flex-wrap:wrap;margin-bottom:12px;">
                    <label style="display:flex;gap:8px;align-items:center;">
                        <input type="checkbox" id="tm_takeoff_headers" checked>
                        Show Takeoff Sheet headers
                    </label>
                    <label style="display:flex;gap:8px;align-items:center;">
                        <input type="checkbox" id="tm_assembly_headers" checked>
                        Show Assembly headers
                    </label>

                    <label style="display:flex;gap:8px;align-items:center;">
                        <input type="checkbox" id="tm_opt_sortindex" checked>
                        Sort Index
                    </label>
                    <label style="display:flex;gap:8px;align-items:center;">
                        <input type="checkbox" id="tm_opt_isassembly" checked>
                        Is Assembly
                    </label>
                    <label style="display:flex;gap:8px;align-items:center;">
                        <input type="checkbox" id="tm_opt_assemblyname" checked>
                        Assembly Name
                    </label>

                    <button id="tm_select_all" type="button" style="padding:6px 10px;">Enable All Groups</button>
                    <button id="tm_select_none" type="button" style="padding:6px 10px;">Disable All Groups</button>
                    <span id="tm_row_count" style="color:#555;"></span>
                </div>

                <div style="overflow:auto;border:1px solid #ccc;border-radius:6px;flex:1;min-height:320px;">
                    <table style="width:100%;border-collapse:collapse;font-size:13px;">
                        <thead id="tm_table_head" style="position:sticky;top:0;background:#f0f0f0;z-index:1;"></thead>
                        <tbody id="tm_table_body"></tbody>
                    </table>
                </div>

                <div style="display:flex;gap:10px;justify-content:flex-end;margin-top:16px;">
                    <button id="tm_copy_table" type="button" style="padding:8px 14px;">Copy to Clipboard</button>
                    <button id="tm_download_csv" type="button" style="padding:8px 14px;">Download CSV</button>
                    <button id="tm_print_table" type="button" style="padding:8px 14px;">Print</button>
                    <button id="tm_close" type="button" style="padding:8px 14px;">Close</button>
                </div>
            `;

            overlay.appendChild(modal);
            document.body.appendChild(overlay);

            const takeoffHeadersCheckbox = modal.querySelector('#tm_takeoff_headers');
            const assemblyHeadersCheckbox = modal.querySelector('#tm_assembly_headers');
            const optSortIndexCheckbox = modal.querySelector('#tm_opt_sortindex');
            const optIsAssemblyCheckbox = modal.querySelector('#tm_opt_isassembly');
            const optAssemblyNameCheckbox = modal.querySelector('#tm_opt_assemblyname');
            const tableHead = modal.querySelector('#tm_table_head');
            const tableBody = modal.querySelector('#tm_table_body');
            const rowCount = modal.querySelector('#tm_row_count');
            const selectAllButton = modal.querySelector('#tm_select_all');
            const selectNoneButton = modal.querySelector('#tm_select_none');
            const copyTableButton = modal.querySelector('#tm_copy_table');
            const downloadCsvButton = modal.querySelector('#tm_download_csv');
            const printTableButton = modal.querySelector('#tm_print_table');
            const closeButton = modal.querySelector('#tm_close');

            function renderHeader() {
                tableHead.innerHTML = `
                    <tr>
                        <th style="padding:8px;border:1px solid #ddd;text-align:center;">Include</th>
                        ${optionalColumns.sortIndex ? '<th style="padding:8px;border:1px solid #ddd;text-align:left;">Sort Index</th>' : ''}
                        ${optionalColumns.isAssembly ? '<th style="padding:8px;border:1px solid #ddd;text-align:left;">Is Assembly</th>' : ''}
                        ${optionalColumns.assemblyName ? '<th style="padding:8px;border:1px solid #ddd;text-align:left;">Assembly Name</th>' : ''}
                        <th style="padding:8px;border:1px solid #ddd;text-align:left;">Takeoff Sheet</th>
                        <th style="padding:8px;border:1px solid #ddd;text-align:center;">Assembly Depth</th>
                        <th style="padding:8px;border:1px solid #ddd;text-align:left;">Parent Item</th>
                        <th style="padding:8px;border:1px solid #ddd;text-align:left;">Part Number</th>
                        <th style="padding:8px;border:1px solid #ddd;text-align:left;">Description</th>
                        <th style="padding:8px;border:1px solid #ddd;text-align:right;">Qty</th>
                        <th style="padding:8px;border:1px solid #ddd;text-align:right;">Extended Qty</th>
                    </tr>
                `;
            }

            function getVisible() {
                return getVisibleRows(allRows, includeTakeoffHeaders, includeAssemblyHeaders);
            }

            function getExportable() {
                return getExportRows(getVisible(), groupState);
            }

            function bindGroupCheckboxes() {
                Array.from(tableBody.querySelectorAll('.tm-group-toggle')).forEach(cb => {
                    cb.addEventListener('change', () => {
                        setGroupStateRecursive(cb.dataset.groupKey, cb.checked, childrenMap, groupState);
                        render();
                    });
                });
            }

            function updateCounts() {
                const visible = getVisible();
                const exportable = getExportable();
                rowCount.textContent = `${exportable.length} export rows / ${visible.length} visible rows`;
            }

            function render() {
                renderHeader();
                const visible = getVisible();
                tableBody.innerHTML = renderTableRows(visible, groupState, optionalColumns);
                bindGroupCheckboxes();
                updateCounts();
            }

            takeoffHeadersCheckbox.addEventListener('change', () => {
                includeTakeoffHeaders = takeoffHeadersCheckbox.checked;
                render();
            });

            assemblyHeadersCheckbox.addEventListener('change', () => {
                includeAssemblyHeaders = assemblyHeadersCheckbox.checked;
                render();
            });

            optSortIndexCheckbox.addEventListener('change', () => {
                optionalColumns.sortIndex = optSortIndexCheckbox.checked;
                render();
            });

            optIsAssemblyCheckbox.addEventListener('change', () => {
                optionalColumns.isAssembly = optIsAssemblyCheckbox.checked;
                render();
            });

            optAssemblyNameCheckbox.addEventListener('change', () => {
                optionalColumns.assemblyName = optAssemblyNameCheckbox.checked;
                render();
            });

            selectAllButton.addEventListener('click', () => {
                Object.keys(groupState).forEach(k => groupState[k] = true);
                render();
            });

            selectNoneButton.addEventListener('click', () => {
                Object.keys(groupState).forEach(k => groupState[k] = false);
                render();
            });

            copyTableButton.addEventListener('click', async () => {
                const rows = getExportable();
                const ok = rows.length ? await copyTableToClipboard(rows, optionalColumns) : false;

                const oldText = copyTableButton.textContent;
                copyTableButton.textContent = ok ? 'Copied' : 'Nothing Selected';
                setTimeout(() => {
                    copyTableButton.textContent = oldText;
                }, 1500);
            });

            downloadCsvButton.addEventListener('click', () => {
                const rows = getExportable();
                if (!rows.length) return;
                downloadFile('quote_export.csv', toCsv(rows, optionalColumns), 'text/csv;charset=utf-8;');
            });

            printTableButton.addEventListener('click', () => {
                const rows = getExportable();
                if (!rows.length) return;
                printHtmlTable(rows, optionalColumns);
            });

            function close() {
                overlay.remove();
                resolve();
            }

            closeButton.addEventListener('click', close);
            overlay.addEventListener('click', (e) => {
                if (e.target === overlay) close();
            });

            render();
        });
    }

    function showSimpleMessage(title, message) {
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
                width: min(500px, 92vw);
                border-radius: 10px;
                box-shadow: 0 10px 30px rgba(0,0,0,0.25);
                padding: 18px;
                font-family: Arial, sans-serif;
            `;

            modal.innerHTML = `
                <div style="font-size:18px;font-weight:bold;margin-bottom:12px;">
                    ${escapeHtml(title)}
                </div>
                <div style="margin-bottom:16px;white-space:pre-line;">
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

    async function runQuoteExportTable() {
        const allRows = collectTakeoffData();

        if (!allRows.length) {
            await showSimpleMessage('No quote data found', 'No quote line items were found on the page.');
            return;
        }

        await showExportModal(allRows);
    }

    document.addEventListener('keydown', function (event) {
        if (event.ctrlKey && event.shiftKey && event.code === 'KeyE') {
            event.preventDefault();
            runQuoteExportTable();
        }
    });
})();
