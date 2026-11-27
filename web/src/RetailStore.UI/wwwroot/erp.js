// RetailStore web: the few browser calls MudBlazor has no API for.
window.erp = {
    // the light / dark choice: kept a year in the cookie "theme" (read back by App.razor on the next request) and set
    // on <html data-theme> at once, which the page styles of retailstore.css follow
    setTheme: function (value) {
        document.cookie = "theme=" + encodeURIComponent(value) + "; path=/; max-age=31536000; samesite=lax";
        document.documentElement.setAttribute("data-theme", value);
    },

    // a file made by the page (the CSV export of EntityGrid): a Blob, never a data: link holding the whole file
    download: function (fileName, text, mime) {
        const blob = new Blob([text], { type: mime || "text/csv;charset=utf-8" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
    },

    // bring a section of the page into view (the item card of /stock after its row is chosen)
    scrollToId: function (id) {
        const el = document.getElementById(id);
        if (el) {
            el.scrollIntoView({ behavior: "smooth", block: "start" });
        }
    },

    // the keyboard shortcuts of the screens (Services/ShortcutService.cs): F2 save, F3 search, F4 new, Esc close,
    // Ctrl+P print. Only the keys that have a handler in .NET are taken from the browser; F3 without one goes to the
    // search box of the page, Ctrl+P without one stays the print of the browser.
    shortcuts: {
        owner: null,
        keys: [],
        attach: function (owner, keys) {
            const s = window.erp.shortcuts;
            s.owner = owner;
            s.keys = keys || [];
            if (!s.listening) {
                document.addEventListener("keydown", s.onKey, true);
                s.listening = true;
            }
        },
        detach: function () {
            const s = window.erp.shortcuts;
            s.owner = null;
            s.keys = [];
        },
        name: function (e) {
            if (e.key === "F2" || e.key === "F3" || e.key === "F4" || e.key === "Escape") {
                return e.ctrlKey || e.altKey || e.metaKey ? null : e.key;
            }
            if ((e.ctrlKey || e.metaKey) && !e.altKey && (e.code === "KeyP" || e.key === "p" || e.key === "P")) {
                return "Ctrl+P";
            }
            return null;
        },
        dialogOpen: function () {
            return document.querySelector(".mud-dialog-container") !== null;
        },
        onKey: function (e) {
            const s = window.erp.shortcuts;
            const name = s.name(e);
            if (!name || e.repeat) {
                return;
            }
            // an open dialog closes itself on Esc (MudDialogProvider CloseOnEscapeKey)
            if (name === "Escape" && s.dialogOpen()) {
                return;
            }
            if (s.owner && s.keys.indexOf(name) >= 0) {
                e.preventDefault();
                e.stopPropagation();
                // the value being typed is kept by its field before the handler runs
                if (document.activeElement && document.activeElement.dispatchEvent) {
                    document.activeElement.dispatchEvent(new Event("change", { bubbles: true }));
                }
                s.owner.invokeMethodAsync("OnKey", name);
                return;
            }
            if (name === "F3") {
                const box = document.querySelector("[data-erp-search] input, input[data-erp-search]");
                if (box) {
                    e.preventDefault();
                    box.focus();
                    box.select();
                }
            }
        }
    },

    // the lines of a document (Components/Shared/InvoiceLinesEditor.razor): Enter goes to the next cell, after the
    // last cell a new line; Delete on an empty cell (or Ctrl+Delete) removes the line
    lines: {
        attach: function (table, owner) {
            if (!table || table.erpLines) {
                return;
            }
            table.erpLines = true;
            table.addEventListener("keydown", function (e) {
                const input = e.target;
                if (!input || input.tagName !== "INPUT") {
                    return;
                }
                const cell = input.closest("[data-cell]");
                const row = input.closest("tr[data-row]");
                if (!cell || !row) {
                    return;
                }
                if (e.key === "Enter" && !e.shiftKey && !e.ctrlKey && !e.altKey) {
                    // the list of a lookup field is open: Enter chooses its item
                    if (cell.querySelector(".mud-autocomplete") && document.querySelector(".mud-popover-open .mud-list-item")) {
                        return;
                    }
                    e.preventDefault();
                    input.dispatchEvent(new Event("change", { bubbles: true }));
                    const cells = Array.from(table.querySelectorAll("tr[data-row] [data-cell]"))
                        .filter(function (c) { return c.querySelector("input:not([disabled]):not([readonly])"); });
                    const i = cells.indexOf(cell);
                    if (i >= 0 && i < cells.length - 1) {
                        window.erp.lines.focusIn(cells[i + 1]);
                    } else {
                        owner.invokeMethodAsync("AddLineFromKeyboard");
                    }
                } else if (e.key === "Delete" && (e.ctrlKey || input.value === "")) {
                    e.preventDefault();
                    owner.invokeMethodAsync("RemoveLineFromKeyboard", parseInt(row.dataset.row, 10));
                }
            });
        },
        focusIn: function (cell) {
            const input = cell && cell.querySelector("input:not([disabled]):not([readonly])");
            if (input) {
                input.focus();
                if (input.select) {
                    input.select();
                }
            }
        },
        focusCell: function (table, rowIndex, cellIndex) {
            if (!table) {
                return;
            }
            const row = table.querySelector("tr[data-row='" + rowIndex + "']");
            if (row) {
                const cells = row.querySelectorAll("[data-cell]");
                window.erp.lines.focusIn(cells[Math.min(cellIndex, cells.length - 1)]);
            }
        }
    }
};
