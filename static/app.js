function attachAutocomplete(inputEl, listEl, opts) {
  opts = opts || {};
  let debounceTimer = null;
  let items = [];
  let activeIndex = -1;

  function render() {
    listEl.innerHTML = "";
    items.forEach(function (s, i) {
      const div = document.createElement("div");
      div.className = "autocomplete-item" + (i === activeIndex ? " active" : "");
      div.innerHTML = s.name + ' <span class="meta">Class ' + (s.class || "?") + (s.section ? "-" + s.section : "") + "</span>";
      div.addEventListener("mousedown", function (e) {
        e.preventDefault();
        select(s);
      });
      listEl.appendChild(div);
    });
    listEl.classList.toggle("open", items.length > 0);
  }

  function select(s) {
    inputEl.value = s.name;
    listEl.classList.remove("open");
    if (opts.onSelect) opts.onSelect(s);
  }

  function search(q) {
    fetch("/api/students?q=" + encodeURIComponent(q))
      .then(function (r) { return r.json(); })
      .then(function (data) {
        items = data;
        activeIndex = -1;
        render();
      });
  }

  inputEl.addEventListener("input", function () {
    clearTimeout(debounceTimer);
    const q = inputEl.value.trim();
    debounceTimer = setTimeout(function () { search(q); }, 150);
  });

  inputEl.addEventListener("focus", function () {
    search(inputEl.value.trim());
  });

  inputEl.addEventListener("keydown", function (e) {
    if (!listEl.classList.contains("open")) return;
    if (e.key === "ArrowDown") {
      e.preventDefault();
      activeIndex = Math.min(activeIndex + 1, items.length - 1);
      render();
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      activeIndex = Math.max(activeIndex - 1, 0);
      render();
    } else if (e.key === "Enter") {
      if (activeIndex >= 0 && items[activeIndex]) {
        e.preventDefault();
        select(items[activeIndex]);
      }
    } else if (e.key === "Escape") {
      listEl.classList.remove("open");
    }
  });

  document.addEventListener("click", function (e) {
    if (e.target !== inputEl) listEl.classList.remove("open");
  });
}
