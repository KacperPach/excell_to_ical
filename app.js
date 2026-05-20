/* app.js — CalSync: Outlook-Export Excel → .ics converter (Client-Side Only)
   Optimised for German Outlook exports:
     Col A  "Betreff"     → event title
     Col D  "Beginnt am"  → start  (e.g. "Montag, 18. Mai 2026 09:00")
     Col E  "Endet am"    → end    (same format)
   All events are written as ALL-DAY events (time part is ignored).
   Pure client-side - no server needed!
*/

/* ─── German month map ──────────────────────────────────────────────────────── */
const DE_MONTHS = {
  januar: 1,
  februar: 2,
  märz: 3,
  april: 4,
  mai: 5,
  juni: 6,
  juli: 7,
  august: 8,
  september: 9,
  oktober: 10,
  november: 11,
  dezember: 12,
};

/* ─── State ─────────────────────────────────────────────────────────────────── */
let workbook = null;
let sheetData = [];
let headers = [];
let fileName = "";
let currentIcsContent = "";

/* ─── DOM refs ──────────────────────────────────────────────────────────────── */
const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("fileInput");
const mappingSection = document.getElementById("mappingSection");
const fileNameLabel = document.getElementById("fileNameLabel");
const rowCount = document.getElementById("rowCount");
const convertBtn = document.getElementById("convertBtn");
const resetBtn = document.getElementById("resetBtn");
const toast = document.getElementById("toast");
const previewHead = document.getElementById("previewHead");
const previewBody = document.getElementById("previewBody");

const selTitle = document.getElementById("colTitle");
const selStartDate = document.getElementById("colStartDate");
const selEndDate = document.getElementById("colEndDate");

/* ─── Drag & Drop ───────────────────────────────────────────────────────────── */
dropzone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropzone.classList.add("drag-over");
});
dropzone.addEventListener("dragleave", () =>
  dropzone.classList.remove("drag-over"),
);
dropzone.addEventListener("drop", (e) => {
  e.preventDefault();
  dropzone.classList.remove("drag-over");
  const file = e.dataTransfer.files[0];
  if (file) handleFile(file);
});
dropzone.addEventListener("click", () => fileInput.click());
dropzone.addEventListener("keydown", (e) => {
  if (e.key === "Enter" || e.key === " ") {
    e.preventDefault();
    fileInput.click();
  }
});
fileInput.addEventListener("change", () => {
  if (fileInput.files[0]) handleFile(fileInput.files[0]);
});

/* ─── File Handler ──────────────────────────────────────────────────────────── */
function handleFile(file) {
  const ext = file.name.toLowerCase().split(".").pop();
  if (!["xls", "xlsx"].includes(ext)) {
    showToast("Only .xls or .xlsx Excel files are supported.", "error");
    return;
  }
  fileName = file.name;

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      workbook = XLSX.read(data, { type: "array", cellDates: true });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      sheetData = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
      headers = sheetData.length ? Object.keys(sheetData[0]) : [];
      populateMappingUI();
      renderPreview();
      mappingSection.classList.remove("hidden");
      mappingSection.scrollIntoView({ behavior: "smooth", block: "start" });
      showToast(
        `File loaded: ${headers.length} columns, ${sheetData.length} rows`,
        "",
      );
    } catch (err) {
      console.error("Error reading Excel file:", err);
      showToast(
        "Failed to read the Excel file. Make sure it's a valid .xls or .xlsx file.",
        "error",
      );
    }
  };
  reader.readAsArrayBuffer(file);
}

/* ─── Populate Mapping UI ───────────────────────────────────────────────────── */
function populateMappingUI() {
  [selTitle, selStartDate, selEndDate].forEach((sel) => (sel.innerHTML = ""));

  // Build options for all dropdowns
  const optAll = headers
    .map((h, i) => `<option value="${i}">${escHtml(h)}</option>`)
    .join("");
  const optReq = '<option value="">— Select a column —</option>' + optAll;

  selTitle.innerHTML = optReq;
  selStartDate.innerHTML = optReq;
  selEndDate.innerHTML = '<option value="">— Optional —</option>' + optAll;

  // Auto-map known German column names
  autoSelect(selTitle, ["betreff", "titel", "subject", "title"]);
  autoSelect(selStartDate, ["beginnt am", "start", "startdatum", "start date"]);
  autoSelect(selEndDate, ["endet am", "end", "enddatum", "end date"]);

  fileNameLabel.textContent = fileName;

  // Add listeners to refresh preview
  [selTitle, selStartDate, selEndDate].forEach((sel) => {
    sel.addEventListener("change", renderPreview);
  });
}

/* ─── Auto-select columns ───────────────────────────────────────────────────── */
function autoSelect(selectElem, keywords) {
  for (let i = 0; i < headers.length; i++) {
    const lc = headers[i].toLowerCase().trim();
    for (const kw of keywords) {
      const idx = headers.indexOf(headers[i]);
      if (lc.includes(kw.toLowerCase())) {
        selectElem.value = idx;
        return;
      }
    }
  }
}

/* ─── Preview Table ─────────────────────────────────────────────────────────── */
function renderPreview() {
  const activeCols = [
    { label: "Event Title", key: selTitle.value },
    { label: "Start Date", key: selStartDate.value },
    { label: "End Date", key: selEndDate.value },
  ].filter((c) => c.key !== "");

  previewHead.innerHTML = `<tr>${activeCols.map((c) => `<th>${c.label}</th>`).join("")}</tr>`;

  const rows = sheetData.slice(0, 5).map((row) => {
    const cells = activeCols.map((c) => {
      const val = row[headers[c.key]] || "";
      return `<td>${escHtml(truncate(String(val), 40))}</td>`;
    });
    return `<tr>${cells.join("")}</tr>`;
  });
  previewBody.innerHTML = rows.join("");
  rowCount.textContent = sheetData.length;
}

/* ─── Sanitization function to strip HTML/JS ───────────────────────────────── */
function stripHtml(str) {
  if (!str) return "";

  // Decode HTML entities first
  let s = str
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&amp;/gi, "&");

  // Remove script tags
  let prev;
  do {
    prev = s;
    s = s.replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, "");
  } while (s !== prev);

  // Remove all HTML tags
  do {
    prev = s;
    s = s.replace(/<[^>]*>/g, "");
  } while (s !== prev);

  // Strip dangerous URI schemes
  s = s.replace(
    /(?:javascript|data|vbscript|onload|onerror|onmouseover|onfocus|onblur):\s*/gi,
    "",
  );

  // Strip HTML entities
  s = s.replace(/&#\d+;/g, "").replace(/&#x[0-9a-fA-F]+;/g, "");

  return s.trim();
}

/* ─── Conversion & Download ─────────────────────────────────────────────────── */
convertBtn.addEventListener("click", () => {
  if (!selTitle.value || !selStartDate.value) {
    showToast('Please map at least "Event Title" and "Start Date".', "error");
    return;
  }

  const events = [];
  const skipped = [];

  sheetData.forEach((row, i) => {
    // Sanitize all input data
    const title = stripHtml(String(row[headers[selTitle.value]] ?? "").trim());
    const startRaw = stripHtml(
      String(row[headers[selStartDate.value]] ?? "").trim(),
    );
    const endRaw = selEndDate.value
      ? stripHtml(String(row[headers[selEndDate.value]] ?? "").trim())
      : "";

    if (!title || !startRaw) {
      skipped.push(i + 2);
      return;
    }

    const startDate = parseGermanDate(startRaw);
    if (!startDate) {
      skipped.push(i + 2);
      return;
    }

    let endDate = endRaw ? parseGermanDate(endRaw) : null;
    if (!endDate) {
      endDate = new Date(startDate);
    }
    const endExclusive = new Date(endDate);
    endExclusive.setDate(endExclusive.getDate() + 1);

    events.push({ title, startDate, endDate: endExclusive });
  });

  if (!events.length) {
    showToast(
      "No valid events found. Check your column mapping and date format.",
      "error",
    );
    return;
  }

  const ics = buildICS(events);
  currentIcsContent = ics;

  // Download the .ics file directly
  const downloadName = fileName.replace(/\.(xls|xlsx)$/i, "") + ".ics";
  downloadFile(ics, downloadName);

  if (skipped.length) {
    showToast(
      `Converted ${events.length} events (skipped ${skipped.length} row(s)). File downloaded!`,
      "",
    );
  } else {
    showToast(
      `Success! Converted ${events.length} events. File downloaded!`,
      "",
    );
  }
});

/* ─── Reset Handler ─────────────────────────────────────────────────────────── */
resetBtn.addEventListener("click", () => {
  workbook = null;
  sheetData = [];
  headers = [];
  fileName = "";
  currentIcsContent = "";
  fileInput.value = "";
  mappingSection.classList.add("hidden");
  window.scrollTo({ top: 0, behavior: "smooth" });
  showToast("Reset complete. Upload a new file to start.", "");
});

/* ─── Parse German Date ─────────────────────────────────────────────────────── */
function parseGermanDate(raw) {
  if (!raw) return null;

  // Clean the input
  let s = raw.toLowerCase().trim();
  const m = s.match(/(\d{1,2})\.\s*(\w+)\s*(\d{4})/);
  if (!m) return null;

  const day = parseInt(m[1], 10);
  const month = DE_MONTHS[m[2].toLowerCase()];
  const year = parseInt(m[3], 10);

  if (!month || day < 1 || day > 31 || year < 1900 || year > 2100) return null;

  const d = new Date(year, month - 1, day);
  return d;
}

/* ─── Build ICS Content ─────────────────────────────────────────────────────── */
function buildICS(events) {
  const stamp = formatDT(new Date());
  const lines = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//CalSync//Excel to ICS Converter//EN",
    "CALSCALE:GREGORIAN",
    "METHOD:PUBLISH",
  ];

  events.forEach((evt, i) => {
    const uid = `event-${Date.now()}-${i}@calsync.local`;
    lines.push("BEGIN:VEVENT");
    lines.push(`UID:${uid}`);
    lines.push(`DTSTAMP:${stamp}`);
    lines.push(`DTSTART;VALUE=DATE:${formatDate(evt.startDate)}`);
    lines.push(`DTEND;VALUE=DATE:${formatDate(evt.endDate)}`);
    lines.push(`SUMMARY:${icsEscape(evt.title)}`);
    lines.push("END:VEVENT");
  });

  lines.push("END:VCALENDAR");
  return foldLines(lines);
}

/* ─── Fold Lines (RFC 5545 compliance) ──────────────────────────────────────── */
function foldLines(lines) {
  const out = [];
  for (const line of lines) {
    if (line.length <= 75) {
      out.push(line);
    } else {
      out.push(line.slice(0, 75));
      let pos = 75;
      while (pos < line.length) {
        out.push(" " + line.slice(pos, pos + 74));
        pos += 74;
      }
    }
  }
  return out.join("\r\n");
}

/* ─── Date Formatters ───────────────────────────────────────────────────────── */
function formatDate(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}${m}${day}`;
}

function formatDT(d) {
  return d.toISOString().replace(/[-:]/g, "").split(".")[0] + "Z";
}

/* ─── ICS Escape ────────────────────────────────────────────────────────────── */
function icsEscape(str) {
  return str
    .replace(/\\/g, "\\\\")
    .replace(/;/g, "\\;")
    .replace(/,/g, "\\,")
    .replace(/\n/g, "\\n");
}

/* ─── Download File Helper ──────────────────────────────────────────────────── */
function downloadFile(content, filename) {
  const blob = new Blob([content], { type: "text/calendar;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = Object.assign(document.createElement("a"), {
    href: url,
    download: filename,
  });
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/* ─── Toast Notifications ───────────────────────────────────────────────────── */
let toastTimer = null;
function showToast(msg, type = "") {
  toast.textContent = msg;
  toast.className = "toast";
  if (type === "error") toast.classList.add("toast-error");
  toast.classList.remove("hidden");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.add("hidden"), 4000);
}

/* ─── Utility Functions ─────────────────────────────────────────────────────── */
function escHtml(str) {
  const div = document.createElement("div");
  div.textContent = str;
  return div.innerHTML;
}
function truncate(str, len) {
  return str.length > len ? str.substring(0, len) + "..." : str;
}
