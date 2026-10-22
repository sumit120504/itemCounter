(function () {
    "use strict";

    var STORAGE_KEY = "item-counter:items";
    var TITLE_STORAGE_KEY = "item-counter:title";

    // ---------- State ----------
    // items: [{ id: number, name: string, count: number }]
    var items = load();

    // ---------- DOM references ----------
    var titleInput = document.getElementById("list-title");
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

    // ---------- Storage ----------
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

    function loadTitle() {
        try {
            return localStorage.getItem(TITLE_STORAGE_KEY) || "";
        } catch (e) {
            return "";
        }
    }

    function saveTitle(val) {
        try {
            localStorage.setItem(TITLE_STORAGE_KEY, val);
        } catch (e) {}
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

    function moveItem(index, delta) {
        var targetIndex = index + delta;
        if (targetIndex < 0 || targetIndex >= items.length) return;
        var temp = items[index];
        items[index] = items[targetIndex];
        items[targetIndex] = temp;
        commit();
    }

    function reorderItems(fromIndex, toIndex) {
        if (fromIndex === toIndex) return;
        if (fromIndex < 0 || fromIndex >= items.length) return;
        if (toIndex < 0 || toIndex >= items.length) return;
        var moved = items.splice(fromIndex, 1)[0];
        items.splice(toIndex, 0, moved);
        commit();
    }

    function commit() {
        save();
        render();
    }

    // ---------- Export File Helpers ----------
    function getExportInfo() {
        var raw = (titleInput && titleInput.value ? titleInput.value.trim() : "");
        var dateStr = new Date().toISOString().slice(0, 10);
        if (!raw) {
            return {
                filename: "item-counts_" + dateStr + ".xlsx",
                sheetName: "Item Counts"
            };
        }
        var sanitized = raw.replace(/[\\/:*?"<>|]+/g, "_").trim();
        if (!sanitized) {
            sanitized = "item-counts_" + dateStr;
        }
        var sheet = raw.replace(/[\\/:*?\[\]]+/g, "_").trim().substring(0, 31) || "Item Counts";
        return {
            filename: sanitized + ".xlsx",
            sheetName: sheet
        };
    }

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

    function exportSpreadsheetXml(filename, sheetName) {
        sheetName = sheetName || "Item Counts";
        var safeSheet = sheetName
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&apos;");

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
            ' <Worksheet ss:Name="' + safeSheet + '">\r\n' +
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

        var info = getExportInfo();
        var filename = info.filename;
        var sheetName = info.sheetName;

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

                XLSX.utils.book_append_sheet(wb, ws, sheetName);
                XLSX.writeFile(wb, filename);
                showToast("Exported to " + filename);
            } else {
                var xlsName = filename.replace(/\.xlsx$/i, ".xls");
                exportSpreadsheetXml(xlsName, sheetName);
                showToast("Exported to " + xlsName);
            }
        } catch (err) {
            console.error("Export error, using XML fallback:", err);
            var fallbackXls = filename.replace(/\.xlsx$/i, ".xls");
            exportSpreadsheetXml(fallbackXls, sheetName);
            showToast("Exported to " + fallbackXls);
        }
    }

    // ---------- Rendering ----------
    var PLUS_ICON =
        '<svg width="20" height="20" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" aria-hidden="true"><path d="M4 10h12M10 4v12"/></svg>';
    var MINUS_ICON =
        '<svg width="20" height="20" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" aria-hidden="true"><path d="M4 10h12"/></svg>';
    var TRASH_ICON =
        '<svg width="20" height="20" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3.5 5.5h13"/><path d="M8 5.5V3.5a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1v2"/><path d="M5 5.5l1 11a1.5 1.5 0 0 0 1.5 1.5h5a1.5 1.5 0 0 0 1.5-1.5l1-11"/><path d="M8.5 9v5"/><path d="M11.5 9v5"/></svg>';
    var UP_ICON =
        '<svg width="14" height="14" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M15 12l-5-5-5 5"/></svg>';
    var DOWN_ICON =
        '<svg width="14" height="14" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M5 8l5 5 5-5"/></svg>';
    var GRIP_ICON =
        '<svg width="14" height="14" viewBox="0 0 20 20" fill="currentColor" aria-hidden="true"><circle cx="7" cy="5" r="1.5"/><circle cx="13" cy="5" r="1.5"/><circle cx="7" cy="10" r="1.5"/><circle cx="13" cy="10" r="1.5"/><circle cx="7" cy="15" r="1.5"/><circle cx="13" cy="15" r="1.5"/></svg>';

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

    var draggedIndex = null;

    function clearDragIndicators() {
        var rows = list.querySelectorAll(".item");
        rows.forEach(function (r) {
            r.classList.remove("drop-above", "drop-below");
        });
    }

    function attachDragEvents(li, grip) {
        // Desktop HTML5 drag & drop
        li.addEventListener("dragstart", function (e) {
            if (e.target.closest("button")) {
                e.preventDefault();
                return;
            }
            draggedIndex = Number(li.dataset.index);
            e.dataTransfer.effectAllowed = "move";
            e.dataTransfer.setData("text/plain", String(draggedIndex));
            setTimeout(function () {
                li.classList.add("is-dragging");
            }, 0);
        });

        li.addEventListener("dragend", function () {
            li.classList.remove("is-dragging");
            clearDragIndicators();
            draggedIndex = null;
        });

        li.addEventListener("dragover", function (e) {
            e.preventDefault();
            if (draggedIndex === null) return;
            var targetIdx = Number(li.dataset.index);
            if (targetIdx === draggedIndex) {
                clearDragIndicators();
                return;
            }
            e.dataTransfer.dropEffect = "move";
            var rect = li.getBoundingClientRect();
            var isAbove = e.clientY < rect.top + rect.height / 2;
            clearDragIndicators();
            if (isAbove) {
                li.classList.add("drop-above");
            } else {
                li.classList.add("drop-below");
            }
        });

        li.addEventListener("dragleave", function () {
            li.classList.remove("drop-above", "drop-below");
        });

        li.addEventListener("drop", function (e) {
            e.preventDefault();
            if (draggedIndex === null) return;
            var targetIdx = Number(li.dataset.index);
            var rect = li.getBoundingClientRect();
            var isAbove = e.clientY < rect.top + rect.height / 2;
            var toIndex = isAbove ? targetIdx : targetIdx + 1;
            if (toIndex > draggedIndex) {
                toIndex--;
            }
            clearDragIndicators();
            reorderItems(draggedIndex, toIndex);
            draggedIndex = null;
        });

        // Mobile touch slide support on grip
        grip.addEventListener("touchstart", function (e) {
            draggedIndex = Number(li.dataset.index);
            li.classList.add("is-touch-dragging");
        }, { passive: true });

        grip.addEventListener("touchmove", function (e) {
            if (draggedIndex === null) return;
            var touch = e.touches[0];
            var elem = document.elementFromPoint(touch.clientX, touch.clientY);
            var targetRow = elem ? elem.closest(".item") : null;
            clearDragIndicators();
            if (targetRow && targetRow !== li) {
                var rect = targetRow.getBoundingClientRect();
                if (touch.clientY < rect.top + rect.height / 2) {
                    targetRow.classList.add("drop-above");
                } else {
                    targetRow.classList.add("drop-below");
                }
            }
        }, { passive: true });

        grip.addEventListener("touchend", function (e) {
            if (draggedIndex === null) return;
            var changedTouch = e.changedTouches[0];
            var elem = document.elementFromPoint(changedTouch.clientX, changedTouch.clientY);
            var targetRow = elem ? elem.closest(".item") : null;
            li.classList.remove("is-touch-dragging");
            if (targetRow && targetRow !== li) {
                var targetIdx = Number(targetRow.dataset.index);
                var rect = targetRow.getBoundingClientRect();
                var isAbove = changedTouch.clientY < rect.top + rect.height / 2;
                var toIndex = isAbove ? targetIdx : targetIdx + 1;
                if (toIndex > draggedIndex) {
                    toIndex--;
                }
                clearDragIndicators();
                reorderItems(draggedIndex, toIndex);
            }
            clearDragIndicators();
            draggedIndex = null;
        });
    }

    function renderItem(item, index) {
        var li = document.createElement("li");
        li.className = "item";
        li.dataset.id = String(item.id);
        li.dataset.index = String(index);
        li.draggable = true;

        // Reorder controls (Move Up / Grip / Move Down)
        var reorder = document.createElement("div");
        reorder.className = "item-reorder";

        var moveUp = document.createElement("button");
        moveUp.type = "button";
        moveUp.className = "btn-move";
        moveUp.dataset.action = "move-up";
        moveUp.setAttribute("aria-label", "Move " + item.name + " up");
        moveUp.title = "Move up";
        moveUp.disabled = index === 0;
        moveUp.innerHTML = UP_ICON;

        var grip = document.createElement("div");
        grip.className = "drag-grip";
        grip.setAttribute("title", "Drag to slide up or down");
        grip.setAttribute("aria-label", "Drag to reorder");
        grip.innerHTML = GRIP_ICON;

        var moveDown = document.createElement("button");
        moveDown.type = "button";
        moveDown.className = "btn-move";
        moveDown.dataset.action = "move-down";
        moveDown.setAttribute("aria-label", "Move " + item.name + " down");
        moveDown.title = "Move down";
        moveDown.disabled = index === items.length - 1;
        moveDown.innerHTML = DOWN_ICON;

        reorder.appendChild(moveUp);
        reorder.appendChild(grip);
        reorder.appendChild(moveDown);

        var name = document.createElement("div");
        name.className = "item-name";
        name.textContent = item.name;

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

        li.appendChild(reorder);
        li.appendChild(name);
        li.appendChild(counter);

        attachDragEvents(li, grip);

        return li;
    }

    function render() {
        list.innerHTML = "";
        var totalCount = 0;
        items.forEach(function (item, index) {
            list.appendChild(renderItem(item, index));
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
    if (titleInput) {
        titleInput.value = loadTitle();
        titleInput.addEventListener("input", function () {
            saveTitle(titleInput.value);
        });
        titleInput.addEventListener("keydown", function (e) {
            if (e.key === "Enter") {
                e.preventDefault();
                input.focus();
            }
        });
    }

    form.addEventListener("submit", function (e) {
        e.preventDefault();
        var name = input.value.trim();
        if (!name) return;
        addItem(name);
        input.value = "";
        input.focus();
    });

    // List handles count changes, removal, and moving up/down
    list.addEventListener("click", function (e) {
        var btn = e.target.closest("button");
        if (!btn) return;
        var row = btn.closest(".item");
        if (!row) return;
        var id = Number(row.dataset.id);
        var idx = Number(row.dataset.index);
        var action = btn.dataset.action;

        if (action === "inc") {
            changeCount(id, 1);
        } else if (action === "dec") {
            changeCount(id, -1);
        } else if (action === "remove") {
            removeItem(id);
        } else if (action === "move-up") {
            moveItem(idx, -1);
        } else if (action === "move-down") {
            moveItem(idx, 1);
        }
    });

    if (exportBtn) {
        exportBtn.addEventListener("click", exportToExcel);
    }

    // ---------- Init ----------
    render();
})();