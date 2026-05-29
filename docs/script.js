(function () {
  const STORAGE_KEY = "theme";
  const ORDER = ["system", "light", "dark"];
  const ICON = { system: "◐", light: "☀", dark: "☾" };

  function apply(theme) {
    const root = document.documentElement;
    if (theme === "light" || theme === "dark") {
      root.dataset.theme = theme;
    } else {
      delete root.dataset.theme;
    }
  }

  function read() {
    try {
      const v = localStorage.getItem(STORAGE_KEY);
      return ORDER.includes(v) ? v : "system";
    } catch {
      return "system";
    }
  }

  function write(theme) {
    try {
      if (theme === "system") {
        localStorage.removeItem(STORAGE_KEY);
      } else {
        localStorage.setItem(STORAGE_KEY, theme);
      }
    } catch {
      // ignore
    }
  }

  function updateButton(btn, theme) {
    btn.textContent = ICON[theme];
    btn.setAttribute("aria-label", `Theme: ${theme}`);
    btn.setAttribute("title", `Theme: ${theme}`);
  }

  // Initialize
  const initial = read();
  apply(initial);

  document.addEventListener("DOMContentLoaded", () => {
    const btn = document.getElementById("theme-toggle");
    if (!btn) return;
    updateButton(btn, initial);

    btn.addEventListener("click", () => {
      const current = read();
      const next = ORDER[(ORDER.indexOf(current) + 1) % ORDER.length];
      write(next);
      apply(next);
      updateButton(btn, next);
    });
  });

  // Copy-to-clipboard for [data-copy] buttons
  document.addEventListener("click", (e) => {
    const btn = e.target.closest("[data-copy]");
    if (!btn) return;
    const target = document.getElementById(btn.dataset.copy);
    if (!target) return;
    const text = target.textContent;
    if (!navigator.clipboard || !navigator.clipboard.writeText) return;
    navigator.clipboard
      .writeText(text)
      .then(() => {
        const original = btn.textContent;
        btn.textContent = "Copied";
        setTimeout(() => {
          btn.textContent = original;
        }, 1500);
      })
      .catch(() => {
        // Silent no-op on permission denial or other failure
      });
  });
})();
