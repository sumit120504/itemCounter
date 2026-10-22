(function () {
    "use strict";

    var STORAGE_KEY = "item-counter:items";

    // ---------- State ----------
    // items: [{ id: number, name: string, count: number }]
    var items = load();

    // ---------- DOM references ----------
    var form = document.getElementById("add-form");
    var input = document.getElementById("item-name");
    var list = document.getElementById("item-list");
    var emptyState = document.getElementById("empty-state");
    var exportBtn = document.getElementById("btn-export");
    var statItemsCount = document.getElementById("stat-items-count");
    var statTotalCount = document.getElementById("stat-total-count");
    var toast = document.getElementById("toast");
    var toastMessage = document.getElementById("toast-message");
    var toastTimeout = null;

    // ---------- Toast feedback ----------
    function showToast(message) {
        if (!toast || !toastMessage) return;
        toastMessage.textContent = message;
        toast.hidden = false;
        toast.classList.remove("toast-hidden");
        clearTimeout(toastTimeout);
        toastTimeout = setTimeout(function () {
            toast.classList.add("toast-hidden");
            setTimeout(function () {
                toast.hidden = true;
            }, 300);
        }, 3000);
    }

    // ---------- Storage (optional persistence) ----------
    function load() {
        try {
            var raw = localStorage.getItem(STORAGE_KEY);
            return raw ? JSON.parse(raw) : [];
        } catch (e) {
            return [];
        }
    }

    function save() {
        try {
            localStorage.setItem(STORAGE_KEY, JSON.stringify(items));
        } catch (e) {
            /* storage unavailable: app still works for this session */
        }
    }

    // ---------- Actions ----------
    function addItem(name) {
        items.push({ id: Date.now(), name: name, count: 0 });
        commit();
    }

    function changeCount(id, delta) {
        items = items.map(function (item) {
            if (item.id !== id) return item;
            return { id: item.id, name: item.name, count: Math.max(0, item.count + delta) };
        });
        commit();
    }

    function removeItem(id) {
        items = items.filter(function (item) {
            return item.id !== id;
        });
        commit();
    }

    function commit() {
        save();
        render();
    }

    // ---------- Export ----------
    function downloadBlob(blob, filename) {
        var url = URL.createObjectURL(blob);
        var a = document.createElement("a");
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        setTimeout(function () {
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }, 150);
    }

    function exportSpreadsheetXml(filename) {
        var xml = '<?xml version="1.0" encoding="UTF-8"?>\r\n' +
            '<?mso-application progid="Excel.Sheet"?>\r\n' +
            '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"\r\n' +
            ' xmlns:o="urn:schemas-microsoft-com:office:office"\r\n' +
            ' xmlns:x="urn:schemas-microsoft-com:office:excel"\r\n' +
            ' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"\r\n' +
            ' xmlns:html="http://www.w3.org/TR/REC-html40">\r\n' +
            ' <Styles>\r\n' +
            '  <Style ss:ID="Header">\r\n' +
            '   <Font ss:Bold="1" ss:Color="#FFFFFF" ss:FontName="Segoe UI"/>\r\n' +
            '   <Interior ss:Color="#107C41" ss:Pattern="Solid"/>\r\n' +
            '   <Alignment ss:Horizontal="Center"/>\r\n' +
            '  </Style>\r\n' +
            '  <Style ss:ID="DataCell">\r\n' +
            '   <Font ss:FontName="Segoe UI"/>\r\n' +
            '  </Style>\r\n' +
            '  <Style ss:ID="NumberCell">\r\n' +
            '   <Font ss:FontName="Segoe UI"/>\r\n' +
            '   <Alignment ss:Horizontal="Right"/>\r\n' +
            '  </Style>\r\n' +
            ' </Styles>\r\n' +
            ' <Worksheet ss:Name="Item Counts">\r\n' +
            '  <Table>\r\n' +
            '   <Column ss:Width="200"/>\r\n' +
            '   <Column ss:Width="80"/>\r\n' +
            '   <Row ss:StyleID="Header">\r\n' +
            '    <Cell><Data ss:Type="String">Item Name</Data></Cell>\r\n' +
            '    <Cell><Data ss:Type="String">Count</Data></Cell>\r\n' +
            '   </Row>\r\n';

        items.forEach(function (item) {
            var escapedName = String(item.name)
                .replace(/&/g, "&amp;")
                .replace(/</g, "&lt;")
                .replace(/>/g, "&gt;")
                .replace(/"/g, "&quot;")
                .replace(/'/g, "&apos;");
            xml += '   <Row>\r\n' +
                '    <Cell ss:StyleID="DataCell"><Data ss:Type="String">' + escapedName + '</Data></Cell>\r\n' +
                '    <Cell ss:StyleID="NumberCell"><Data ss:Type="Number">' + item.count + '</Data></Cell>\r\n' +
                '   </Row>\r\n';
        });

        xml += '  </Table>\r\n' +
            ' </Worksheet>\r\n' +
            '</Workbook>';

        var blob = new Blob([xml], { type: "application/vnd.ms-excel;charset=utf-8" });
        downloadBlob(blob, filename);
    }

    function exportToExcel() {
        if (!items || items.length === 0) return;

        var dateStr = new Date().toISOString().slice(0, 10);
        var filename = "item-counts_" + dateStr + ".xlsx";

        try {
            if (typeof XLSX !== "undefined") {
                var rows = [["Item Name", "Count"]];
                items.forEach(function (item) {
                    rows.push([item.name, item.count]);
                });

                var wb = XLSX.utils.book_new();
                var ws = XLSX.utils.aoa_to_sheet(rows);

                var maxNameLen = "Item Name".length;
                items.forEach(function (item) {
                    if (item.name && item.name.length > maxNameLen) {
                        maxNameLen = item.name.length;
                    }
                });

                ws["!cols"] = [
                    { wch: Math.min(Math.max(maxNameLen + 4, 15), 50) },
                    { wch: 12 }
                ];

                XLSX.utils.book_append_sheet(wb, ws, "Item Counts");
                XLSX.writeFile(wb, filename);
                showToast("Exported " + items.length + " item" + (items.length === 1 ? "" : "s") + " to Excel");
            } else {
                var xlsName = "item-counts_" + dateStr + ".xls";
                exportSpreadsheetXml(xlsName);
                showToast("Exported " + items.length + " item" + (items.length === 1 ? "" : "s") + " to Excel (.xls)");
            }
        } catch (err) {
            console.error("Export error, using XML fallback:", err);
            var fallbackXls = "item-counts_" + dateStr + ".xls";
            exportSpreadsheetXml(fallbackXls);
            showToast("Exported " + items.length + " item" + (items.length === 1 ? "" : "s") + " to Excel (.xls)");
        }
    }

    // ---------- Rendering ----------
    var PLUS_ICON =
        '<svg width="20" height="20" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" aria-hidden="true"><path d="M4 10h12M10 4v12"/></svg>';
    var MINUS_ICON =
        '<svg width="20" height="20" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" aria-hidden="true"><path d="M4 10h12"/></svg>';
    var TRASH_ICON =
        '<svg width="20" height="20" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3.5 5.5h13"/><path d="M8 5.5V3.5a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1v2"/><path d="M5 5.5l1 11a1.5 1.5 0 0 0 1.5 1.5h5a1.5 1.5 0 0 0 1.5-1.5l1-11"/><path d="M8.5 9v5"/><path d="M11.5 9v5"/></svg>';

    function stepButton(action, label, icon, disabled, extraClass) {
        var btn = document.createElement("button");
        btn.type = "button";
        btn.className = "btn-step" + (extraClass ? " " + extraClass : "");
        btn.dataset.action = action;
        btn.setAttribute("aria-label", label);
        btn.title = label;
        btn.disabled = !!disabled;
        btn.innerHTML = icon; // static, trusted SVG only
        return btn;
    }

    function renderItem(item) {
        var li = document.createElement("li");
        li.className = "item";
        li.dataset.id = String(item.id);

        var name = document.createElement("div");
        name.className = "item-name";
        name.textContent = item.name; // textContent: user text is never parsed as HTML

        var counter = document.createElement("div");
        counter.className = "counter";

        var count = document.createElement("div");
        count.className = "count";
        count.setAttribute("aria-live", "polite");
        count.textContent = String(item.count);

        counter.appendChild(stepButton("dec", "Decrease " + item.name, MINUS_ICON, item.count === 0));
        counter.appendChild(count);
        counter.appendChild(stepButton("inc", "Increase " + item.name, PLUS_ICON, false));
        counter.appendChild(stepButton("remove", "Remove " + item.name, TRASH_ICON, false, "btn-remove"));

        li.appendChild(name);
        li.appendChild(counter);
        return li;
    }

    function render() {
        list.innerHTML = "";
        var totalCount = 0;
        items.forEach(function (item) {
            list.appendChild(renderItem(item));
            totalCount += item.count;
        });

        var hasItems = items.length > 0;
        emptyState.hidden = hasItems;

        if (statItemsCount) {
            statItemsCount.textContent = items.length + (items.length === 1 ? " item" : " items");
        }
        if (statTotalCount) {
            statTotalCount.textContent = totalCount + " total count";
        }
        if (exportBtn) {
            exportBtn.disabled = !hasItems;
            exportBtn.setAttribute("aria-disabled", String(!hasItems));
        }
    }

    // ---------- Events ----------
    form.addEventListener("submit", function (e) {
        e.preventDefault();
        var name = input.value.trim();
        if (!name) return;
        addItem(name);
        input.value = "";
        input.focus();
    });

    // One listener on the list handles every +, -, and remove button
    list.addEventListener("click", function (e) {
        var btn = e.target.closest(".btn-step");
        if (!btn) return;
        var row = btn.closest(".item");
        var id = Number(row.dataset.id);
        var action = btn.dataset.action;
        if (action === "inc") {
            changeCount(id, 1);
        } else if (action === "dec") {
            changeCount(id, -1);
        } else if (action === "remove") {
            removeItem(id);
        }
    });

    if (exportBtn) {
        exportBtn.addEventListener("click", exportToExcel);
    }

    // ---------- Init ----------
    render();
})();