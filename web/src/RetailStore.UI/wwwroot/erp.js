// RetailStore web: the few browser calls MudBlazor has no API for.
window.erp = {
    // the light / dark choice: kept a year in the cookie "theme" (read back by App.razor on the next request) and set
    // on <html data-theme> at once, which the page styles of retailstore.css follow
    setTheme: function (value) {
        document.cookie = "theme=" + encodeURIComponent(value) + "; path=/; max-age=31536000; samesite=lax";
        document.documentElement.setAttribute("data-theme", value);
    }
};
