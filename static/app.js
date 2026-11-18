function lookupStudent(q, cb) {
  q = (q || "").trim();
  if (!q) { cb(null); return; }
  fetch("/api/students/lookup?q=" + encodeURIComponent(q))
    .then(function (r) { return r.json(); })
    .then(cb)
    .catch(function () { cb(null); });
}

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
      const meta = document.createElement("span");
      meta.className = "meta";
      meta.textContent = " " + s.id + " · Class " + (s.class || "?") + (s.section ? "-" + s.section : "");
      div.appendChild(document.createTextNode(s.name));
      div.appendChild(meta);
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

  // Typed a full name or a student ID without clicking a suggestion: look it
  // up anyway so the rest of the row still fills in.
  inputEl.addEventListener("change", function () {
    lookupStudent(inputEl.value, function (s) {
      if (s) select(s);
      else if (opts.onNoMatch) opts.onNoMatch(inputEl.value.trim());
    });
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

// Manual entry / edit form: picking or typing a name or ID fills in ID,
// class, section and house from the master list.
function attachStudentForm() {
  const nameInput = document.getElementById("name-input");
  if (!nameInput) return;
  const idInput = document.getElementById("student-id-input");
  const classInput = document.getElementById("class-input");
  const sectionInput = document.getElementById("section-input");
  const houseInput = document.getElementById("house-input");
  const status = document.getElementById("student-status");

  function fill(s) {
    nameInput.value = s.name;
    idInput.value = s.id;
    classInput.value = s.class || "";
    sectionInput.value = s.section || "";
    houseInput.value = s.house || "";
    warn("");  // the filled-in fields say it all
  }

  // The status line only ever shows a problem: a name or ID that isn't in the master list.
  function warn(message) {
    status.textContent = message;
    status.hidden = !message;
    status.classList.toggle("is-missing", !!message);
  }

  function notFound(typed) {
    idInput.value = "";
    warn(typed ? "“" + typed + "” isn't in the master list. Add the student first, or fill in class, section and house." : "");
  }

  attachAutocomplete(nameInput, document.getElementById("name-list"), { onSelect: fill, onNoMatch: notFound });

  idInput.addEventListener("change", function () {
    const typed = idInput.value.trim();
    if (!typed) return;
    lookupStudent(typed, function (s) {
      if (s) fill(s);
      else warn("No student with ID “" + typed + "” in the master list.");
    });
  });
}

// Client-side filter for a table: hides rows whose text doesn't contain the query.
function attachTableFilter(input, table, emptyNote) {
  if (!input || !table) return;
  input.addEventListener("input", function () {
    const q = input.value.trim().toLowerCase();
    let shown = 0;
    table.querySelectorAll("tbody tr").forEach(function (tr) {
      const text = (tr.dataset.search || tr.textContent).toLowerCase();
      const match = !q || text.indexOf(q) !== -1;
      tr.hidden = !match;
      if (match) shown++;
    });
    if (emptyNote) emptyNote.hidden = shown > 0;
  });
}

// Any form with data-confirm asks before submitting (used for deletes).
document.addEventListener("submit", function (e) {
  const msg = e.target.dataset && e.target.dataset.confirm;
  if (msg && !window.confirm(msg)) e.preventDefault();
});
