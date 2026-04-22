// Simple interactivity for the ASSP Migration dashboard mockup.
(function () {
    const root = document.documentElement;

    // ----- Theme toggle (persists to localStorage) -----
    const THEME_KEY = 'assp-theme';
    const saved = localStorage.getItem(THEME_KEY);
    if (saved === 'dark' || (!saved && window.matchMedia('(prefers-color-scheme: dark)').matches)) {
        root.setAttribute('data-theme', 'dark');
    }

    const toggle = document.getElementById('themeToggle');
    if (toggle) {
        toggle.addEventListener('click', () => {
            const next = root.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
            if (next === 'dark') root.setAttribute('data-theme', 'dark');
            else root.removeAttribute('data-theme');
            localStorage.setItem(THEME_KEY, next);
        });
    }

    // ----- Reset button: clears search inputs & filter selection -----
    const resetBtn = document.getElementById('resetBtn');
    if (resetBtn) {
        resetBtn.addEventListener('click', () => {
            document.querySelectorAll('input[type="text"]').forEach((el) => { el.value = ''; });
            document.querySelectorAll('.toolbar select').forEach((el) => { el.selectedIndex = 0; });
        });
    }
})();
