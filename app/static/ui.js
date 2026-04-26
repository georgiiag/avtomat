(function () {
  function getStoredTheme() {
    try {
      return String(localStorage.getItem("theme") || "").trim();
    } catch {
      return "";
    }
  }

  function setStoredTheme(theme) {
    try {
      localStorage.setItem("theme", theme);
    } catch {}
  }

  function getSystemTheme() {
    try {
      return window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
    } catch {
      return "light";
    }
  }

  function getCurrentTheme() {
    const t = document.documentElement.getAttribute("data-theme");
    return t === "dark" ? "dark" : "light";
  }

  function applyTheme(theme) {
    const t = theme === "dark" ? "dark" : "light";
    document.documentElement.setAttribute("data-theme", t);
    document.documentElement.style.colorScheme = t;
    document.querySelectorAll("[data-theme-toggle]").forEach((btn) => {
      const pressed = t === "dark";
      btn.setAttribute("aria-pressed", pressed ? "true" : "false");
      const icon = pressed ? "fa-sun" : "fa-moon";
      const label = pressed ? "Светлая тема" : "Тёмная тема";
      btn.setAttribute("title", label);
      btn.setAttribute("aria-label", label);
      btn.innerHTML = `<i class="fa-solid ${icon}"></i>`;
    });
  }

  function toggleTheme() {
    const next = getCurrentTheme() === "dark" ? "light" : "dark";
    setStoredTheme(next);
    applyTheme(next);
  }

  function initToggle() {
    document.querySelectorAll("[data-theme-toggle]").forEach((btn) => {
      btn.addEventListener("click", toggleTheme);
      btn.addEventListener("keydown", (e) => {
        if (e.key === "Enter" || e.key === " ") {
          e.preventDefault();
          toggleTheme();
        }
      });
    });
  }

  function initThemeSync() {
    const mql = window.matchMedia ? window.matchMedia("(prefers-color-scheme: dark)") : null;
    if (!mql || !mql.addEventListener) return;
    mql.addEventListener("change", () => {
      const stored = getStoredTheme();
      if (stored) return;
      applyTheme(getSystemTheme());
    });
  }

  const stored = getStoredTheme();
  applyTheme(stored || getSystemTheme());

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", () => {
      initToggle();
      initThemeSync();
    });
  } else {
    initToggle();
    initThemeSync();
  }
})();

