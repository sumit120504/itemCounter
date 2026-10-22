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

    function commit() {
        save();
        render();
    }

    // ---------- Rendering ----------
    var PLUS_ICON =
        '<svg width="20" height="20" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" aria-hidden="true"><path d="M4 10h12M10 4v12"/></svg>';
    var MINUS_ICON =
        '<svg width="20" height="20" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" aria-hidden="true"><path d="M4 10h12"/></svg>';

    function stepButton(action, label, icon, disabled) {
        var btn = document.createElement("button");
        btn.type = "button";
        btn.className = "btn-step";
        btn.dataset.action = action;
        btn.setAttribute("aria-label", label);
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

        li.appendChild(name);
        li.appendChild(counter);
        return li;
    }

    function render() {
        list.innerHTML = "";
        items.forEach(function (item) {
            list.appendChild(renderItem(item));
        });
        emptyState.hidden = items.length > 0;
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

    // One listener on the list handles every + and - button
    list.addEventListener("click", function (e) {
        var btn = e.target.closest(".btn-step");
        if (!btn) return;
        var row = btn.closest(".item");
        var id = Number(row.dataset.id);
        changeCount(id, btn.dataset.action === "inc" ? 1 : -1);
    });

    // ---------- Init ----------
    render();
})();