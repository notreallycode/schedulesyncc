let events = [];

const DAY_ORDER = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
const statusEl = document.getElementById("status");
const outputEl = document.getElementById("output");
const summaryEl = document.getElementById("summaryText");
const reviewPanel = document.getElementById("reviewPanel");
const uploadPanel = document.getElementById("uploadPanel");
const stepUpload = document.getElementById("stepUpload");
const stepReview = document.getElementById("stepReview");
const detectBtn = document.getElementById("detectBtn");
const fileInput = document.getElementById("image");
const fileInfoEl = document.getElementById("fileInfo");
const uploadStateEl = document.getElementById("uploadState");
const uploadStateTextEl = document.getElementById("uploadStateText");
const ocrSkeletonEl = document.getElementById("ocrSkeleton");
const classModalBackdrop = document.getElementById("classModalBackdrop");
const classForm = document.getElementById("classForm");
const modalTitle = document.getElementById("modalTitle");
const modalEditIndex = document.getElementById("modalEditIndex");
const modalDay = document.getElementById("modalDay");
const modalSubject = document.getElementById("modalSubject");
const modalStart = document.getElementById("modalStart");
const modalEnd = document.getElementById("modalEnd");
const reminderMinutesEl = document.getElementById("reminderMinutes");
const startDateEl = document.getElementById("startDate");
const endDateEl = document.getElementById("endDate");
const undoBtn = document.getElementById("undoBtn");
const uploadPreview = document.getElementById("uploadPreview");
const previewImg = document.getElementById("previewImg");
const scanOverlay = document.getElementById("scanOverlay");
const scanTitle = document.getElementById("scanTitle");
const scanSubtitle = document.getElementById("scanSubtitle");
const uploadSuccessBadge = document.getElementById("uploadSuccessBadge");
const dropContent = document.getElementById("dropContent");
const enhanceAutoEl = document.getElementById("enhanceAuto");
const enhanceInvertEl = document.getElementById("enhanceInvert");
const enhancePixelEl = document.getElementById("enhancePixel");
const rotateLeftBtn = document.getElementById("rotateLeftBtn");
const rotateDegreeEl = document.getElementById("rotateDegree");
const modalErrorEl = document.getElementById("modalError");
const exportModalBackdrop = document.getElementById("exportModalBackdrop");
const exportSummaryBody = document.getElementById("exportSummaryBody");

let previewObjectUrl = null;
let scanStepsTimer = null;
let rotationDeg = 0;
let pendingIcs = "";
let undoStack = [];
const DRAFT_STORAGE_KEY = "autoscan-draft-v1";

function setRotationDegreeLabel() {
  if (rotateDegreeEl) rotateDegreeEl.textContent = `${rotationDeg}°`;
}

function toDateInputValue(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function ensureDateDefaults() {
  if (!startDateEl || !endDateEl) return;
  if (!startDateEl.value) {
    const today = new Date();
    startDateEl.value = toDateInputValue(today);
  }
  if (!endDateEl.value) {
    const end = new Date(startDateEl.value || Date.now());
    end.setMonth(end.getMonth() + 4);
    endDateEl.value = toDateInputValue(end);
  }
}

const SCAN_STEPS = [
  { title: "Uploading image...", subtitle: "Preparing your timetable photo" },
  { title: "Enhancing image...", subtitle: "Improving clarity for OCR" },
  { title: "Reading day rows...", subtitle: "Detecting Monday to Friday blocks" },
  { title: "Mapping time slots...", subtitle: "Aligning classes with timetable columns" },
  { title: "Finalizing schedule...", subtitle: "Building events for review and export" },
];

function startScanSteps() {
  let idx = 0;
  function applyStep() {
    const step = SCAN_STEPS[idx % SCAN_STEPS.length];
    uploadStateTextEl.classList.add("scan-text-swap");
    scanTitle.classList.add("scan-text-swap");
    scanSubtitle.classList.add("scan-text-swap");
    setTimeout(() => {
      uploadStateTextEl.textContent = step.title;
      scanTitle.textContent = step.title;
      scanSubtitle.textContent = step.subtitle;
      uploadStateTextEl.classList.remove("scan-text-swap");
      scanTitle.classList.remove("scan-text-swap");
      scanSubtitle.classList.remove("scan-text-swap");
    }, 130);
    idx += 1;
  }
  applyStep();
  scanStepsTimer = setInterval(applyStep, 1300);
}

function stopScanSteps() {
  if (scanStepsTimer) {
    clearInterval(scanStepsTimer);
    scanStepsTimer = null;
  }
  uploadStateTextEl.textContent = "Uploading...";
  scanTitle.textContent = "Scanning timetable...";
  scanSubtitle.textContent = "AI is reading your schedule";
}

function showUploadSuccessBadge() {
  uploadSuccessBadge.classList.remove("hidden");
  uploadSuccessBadge.classList.remove("show");
  requestAnimationFrame(() => uploadSuccessBadge.classList.add("show"));
  setTimeout(() => {
    uploadSuccessBadge.classList.remove("show");
    uploadSuccessBadge.classList.add("hidden");
  }, 1500);
}

function animateStepTransition(showReview) {
  const showPanel = showReview ? reviewPanel : uploadPanel;
  const hidePanel = showReview ? uploadPanel : reviewPanel;

  hidePanel.classList.remove("hidden");
  hidePanel.classList.remove("panel-enter", "panel-enter-active");
  hidePanel.classList.add("panel-exit");
  requestAnimationFrame(() => {
    hidePanel.classList.add("panel-exit-active");
  });

  setTimeout(() => {
    hidePanel.classList.remove("panel-exit", "panel-exit-active");
    hidePanel.classList.add("hidden");
  }, 220);

  showPanel.classList.remove("hidden", "panel-exit", "panel-exit-active");
  showPanel.classList.add("panel-enter");
  requestAnimationFrame(() => {
    showPanel.classList.add("panel-enter-active");
  });
  setTimeout(() => {
    showPanel.classList.remove("panel-enter", "panel-enter-active");
  }, 320);
}

function setStatus(message, type = "") {
  statusEl.textContent = message || "";
  statusEl.classList.remove("success", "error");
  if (type) statusEl.classList.add(type);
}

function showOcrSkeleton() {
  if (!ocrSkeletonEl) return;
  ocrSkeletonEl.innerHTML = Array.from({ length: 4 })
    .map(
      () => `
      <div class="skeleton-card">
        <div class="skeleton-line long"></div>
        <div class="skeleton-line mid"></div>
        <div class="skeleton-line short"></div>
      </div>
    `
    )
    .join("");
  ocrSkeletonEl.classList.remove("hidden");
}

function hideOcrSkeleton() {
  if (!ocrSkeletonEl) return;
  ocrSkeletonEl.classList.add("hidden");
  ocrSkeletonEl.innerHTML = "";
}

function renderEmptyReviewState(message, tip) {
  outputEl.innerHTML = `
    <div class="empty-day">
      ${escapeHtml(message || "No classes detected yet.")}
      <div class="empty-tip">${escapeHtml(tip || "Tip: try rotating the image, enable Auto Enhance, or use a clearer timetable photo.")}</div>
    </div>
  `;
}

function normalizeDay(day) {
  const s = String(day || "").trim().toLowerCase();
  const map = {
    mon: "Monday",
    monday: "Monday",
    tue: "Tuesday",
    tues: "Tuesday",
    tuesday: "Tuesday",
    wed: "Wednesday",
    wednesday: "Wednesday",
    thu: "Thursday",
    thur: "Thursday",
    thurs: "Thursday",
    thursday: "Thursday",
    fri: "Friday",
    friday: "Friday",
  };
  return map[s] || "Monday";
}

function toMinutes(hm) {
  const m = String(hm || "").match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  const h = Number(m[1]);
  const mm = Number(m[2]);
  if (Number.isNaN(h) || Number.isNaN(mm)) return null;
  return h * 60 + mm;
}

function findConflictIndices(items) {
  const conflicts = new Set();
  const byDay = {};
  items.forEach((e, i) => {
    const d = normalizeDay(e.day);
    if (!byDay[d]) byDay[d] = [];
    byDay[d].push({ ...e, _idx: i });
  });
  Object.values(byDay).forEach((arr) => {
    arr.sort((a, b) => (toMinutes(a.start_time) ?? 0) - (toMinutes(b.start_time) ?? 0));
    for (let i = 1; i < arr.length; i += 1) {
      const prev = arr[i - 1];
      const cur = arr[i];
      const prevEnd = toMinutes(prev.end_time);
      const curStart = toMinutes(cur.start_time);
      if (prevEnd !== null && curStart !== null && curStart < prevEnd) {
        conflicts.add(prev._idx);
        conflicts.add(cur._idx);
      }
    }
  });
  return conflicts;
}


function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function normalizeSubject(raw) {
  const original = String(raw || "").trim();
  if (!original) return "Untitled";

  // Keep an upper-case token for reliable alias matching.
  const compact = original
    .replace(/[().,_\-]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();

  const aliasMap = {
    CHEM: "Chemistry",
    CHEMISTRY: "Chemistry",
    PPS: "PPS",
    BES: "BES",
    BESL: "BESL",
    AC: "A.C.",
    "A C": "A.C.",
    ENG: "English",
    ENGLISH: "English",
    "M II": "Mathematics II",
    "M-II": "Mathematics II",
    "M 2": "Mathematics II",
    "MATH II": "Mathematics II",
    "MATHEMATICS II": "Mathematics II",
    "CHEM L": "Chemistry Lab",
    "CHEM LAB": "Chemistry Lab",
    "C P L": "C.P.L.",
    CPL: "C.P.L.",
  };

  if (aliasMap[compact]) return aliasMap[compact];

  // Title case fallback while preserving short all-caps codes.
  if (/^[A-Z0-9]{2,6}$/.test(compact.replace(/\s/g, ""))) {
    return compact.replace(/\s+/g, " ");
  }
  return original
    .toLowerCase()
    .split(/\s+/)
    .map((w) => w.charAt(0).toUpperCase() + w.slice(1))
    .join(" ");
}

function dedupeEvents(items) {
  const seen = new Map();
  items.forEach((e) => {
    const key = [
      normalizeDay(e.day),
      String(e.start_time || "").trim(),
      String(e.end_time || "").trim(),
      normalizeSubject(e.subject),
    ].join("|");
    if (!seen.has(key)) {
      seen.set(key, {
        day: normalizeDay(e.day),
        subject: normalizeSubject(e.subject),
        start_time: String(e.start_time || "09:00"),
        end_time: String(e.end_time || "09:50"),
        confidence: String(e.confidence || "high"),
        verified: !!e.verified,
        confidence_score:
          typeof e.confidence_score === "number"
            ? e.confidence_score
            : String(e.confidence || "high") === "low"
              ? 0.35
              : String(e.confidence || "high") === "medium"
                ? 0.6
                : 0.9,
      });
    }
  });
  return Array.from(seen.values());
}

function pushUndoSnapshot() {
  undoStack.push(JSON.parse(JSON.stringify(events)));
  if (undoStack.length > 50) undoStack.shift();
  updateUndoButton();
}

function updateUndoButton() {
  if (undoBtn) undoBtn.disabled = undoStack.length === 0;
}

function saveDraft() {
  const payload = {
    events,
    reminderMinutes: reminderMinutesEl?.value || "10",
    startDate: startDateEl?.value || "",
    endDate: endDateEl?.value || "",
    rotationDeg,
    enhance: {
      auto: !!enhanceAutoEl?.checked,
      invert: !!enhanceInvertEl?.checked,
      pixel: !!enhancePixelEl?.checked,
    },
  };
  localStorage.setItem(DRAFT_STORAGE_KEY, JSON.stringify(payload));
}

function restoreDraft() {
  const raw = localStorage.getItem(DRAFT_STORAGE_KEY);
  if (!raw) return;
  try {
    const draft = JSON.parse(raw);
    if (Array.isArray(draft.events) && draft.events.length) {
      events = draft.events.map((e) => ({
        day: normalizeDay(e.day),
        subject: normalizeSubject(e.subject),
        start_time: String(e.start_time || "09:00"),
        end_time: String(e.end_time || "09:50"),
        confidence: String(e.confidence || "medium"),
        verified: !!e.verified,
        confidence_score:
          typeof e.confidence_score === "number"
            ? e.confidence_score
            : String(e.confidence || "medium") === "low"
              ? 0.35
              : String(e.confidence || "medium") === "medium"
                ? 0.6
                : 0.9,
      }));
    }
    if (reminderMinutesEl && draft.reminderMinutes) reminderMinutesEl.value = String(draft.reminderMinutes);
    if (startDateEl && draft.startDate) startDateEl.value = draft.startDate;
    if (endDateEl && draft.endDate) endDateEl.value = draft.endDate;
    if (typeof draft.rotationDeg === "number") {
      rotationDeg = draft.rotationDeg;
      setRotationDegreeLabel();
    }
    if (enhanceAutoEl && draft.enhance) {
      enhanceAutoEl.checked = !!draft.enhance.auto;
      enhanceInvertEl.checked = !!draft.enhance.invert;
      enhancePixelEl.checked = !!draft.enhance.pixel;
    }
    if (events.length) {
      display();
      setStatus("Draft restored from previous session.", "success");
    }
  } catch {
    // ignore invalid draft
  }
}

async function upload() {
  const file = document.getElementById("image").files[0];
  if (!file) {
    setStatus("Please select a timetable image first.", "error");
    renderEmptyReviewState(
      "No image selected.",
      "Tip: upload a timetable photo (PNG/JPG), then click Detect Timetable."
    );
    return;
  }

  detectBtn.disabled = true;
  detectBtn.classList.add("loading");
  uploadStateEl.classList.remove("hidden");
  startScanSteps();
  showOcrSkeleton();
  setStatus("Detecting classes from your image...");
  scanOverlay.classList.remove("hidden");

  let uploadFile = file;
  try {
    uploadFile = await preprocessImage(file, {
      autoEnhance: !!enhanceAutoEl?.checked,
      invert: !!enhanceInvertEl?.checked,
      pixelImprove: !!enhancePixelEl?.checked,
      rotationDeg,
    });
  } catch (e) {
    // Fallback to raw file if enhancement fails.
    uploadFile = file;
  }

  const formData = new FormData();
  formData.append("image", uploadFile, uploadFile.name || file.name);

  try {
    const res = await fetch("/detect", {
      method: "POST",
      body: formData,
    });

    let data;
    try {
      data = await res.json();
    } catch (e) {
      setStatus("Detect failed: server returned non-JSON response.", "error");
      return;
    }

    if (!res.ok) {
      const msg = data && data.error ? data.error : "Request failed";
      setStatus(`Detect failed: ${msg}`, "error");
      return;
    }

    const normalized = (Array.isArray(data) ? data : []).map((e) => ({
      day: normalizeDay(e.day),
      subject: normalizeSubject(e.subject),
      start_time: String(e.start_time || "09:00"),
      end_time: String(e.end_time || "09:50"),
      confidence: String(e.confidence || "medium"),
      verified: false,
      confidence_score:
        typeof e.confidence_score === "number"
          ? e.confidence_score
          : String(e.confidence || "medium") === "low"
            ? 0.35
            : String(e.confidence || "medium") === "medium"
              ? 0.6
              : 0.9,
    }));
    if (events.length) pushUndoSnapshot();
    events = dedupeEvents(normalized);

    display();
    saveDraft();
    setStatus(`Uploaded successfully. Detected ${events.length} classes.`, "success");
    showUploadSuccessBadge();
  } catch (e) {
    setStatus(`Detect failed: ${String(e?.message || e)}`, "error");
  } finally {
    stopScanSteps();
    detectBtn.disabled = false;
    detectBtn.classList.remove("loading");
    uploadStateEl.classList.add("hidden");
    hideOcrSkeleton();
    scanOverlay.classList.add("hidden");
  }
}

function display() {
  if (!events.length) {
    renderEmptyReviewState();
    return;
  }

  const grouped = {};
  const conflicts = findConflictIndices(events);
  for (const day of DAY_ORDER) grouped[day] = [];
  events.forEach((e, idx) => {
    const d = normalizeDay(e.day);
    if (!grouped[d]) grouped[d] = [];
    grouped[d].push({ ...e, _index: idx });
  });

  const sections = [];
  for (const day of DAY_ORDER) {
    const visibleForDay = grouped[day] || [];
    let html = `<section class="day-block day-${day.toLowerCase()}">
      <div class="day-header">
        <h3>${day.toUpperCase()}</h3>
        <button class="icon-btn" onclick="addClassForDay('${day}')">+ Add Class</button>
      </div>
      <div class="day-cards">`;
    if (!visibleForDay.length) {
      html += `<div class="empty-day">No classes yet for ${day}. Use Add Class to set subject and timing.</div>`;
    }
    visibleForDay.forEach((e) => {
      const hasConflict = conflicts.has(e._index);
      const confidence = String(e.confidence || "high");
      const confidenceScore =
        typeof e.confidence_score === "number"
          ? e.confidence_score
          : confidence === "low"
            ? 0.35
            : confidence === "medium"
              ? 0.6
              : 0.9;
      const needsVerify = confidenceScore < 0.65;
      const isVerified = !!e.verified;
      html += `
        <article class="class-card ${hasConflict ? "conflict" : ""}">
          <div class="class-top">
            <div>
              <h4 class="class-title">${escapeHtml(e.subject)}</h4>
              <p class="class-time">${escapeHtml(e.start_time)} - ${escapeHtml(e.end_time)}</p>
              ${needsVerify && !isVerified ? `<button type="button" class="verify-btn" data-verify-index="${e._index}">Reverify it</button>` : ""}
              ${needsVerify && isVerified ? `<span class="confidence-badge verified">Verified</span>` : ""}
              ${hasConflict ? `<p class="class-conflict-note">Time conflict with another class</p>` : ""}
            </div>
            <div class="card-actions">
              <button class="icon-btn" onclick="editClass(${e._index})">Edit</button>
              <button class="icon-btn delete" onclick="deleteClass(${e._index})">Delete</button>
            </div>
          </div>
        </article>
      `;
    });
    html += "</div></section>";
    sections.push(html);
  }

  outputEl.innerHTML = "";
  const staggerBase = events.length >= 18 ? 46 : 14;
  sections.forEach((sectionHtml, idx) => {
    const delay = idx * staggerBase;
    setTimeout(() => {
      outputEl.insertAdjacentHTML("beforeend", sectionHtml);
      const node = outputEl.lastElementChild;
      if (node) node.style.animationDelay = `${delay}ms`;
    }, delay);
  });
  summaryEl.textContent = `${events.length} classes detected${conflicts.size ? ` · ${conflicts.size} conflict(s) found` : ""} · Edit before exporting`;
  animateStepTransition(true);
  stepUpload.classList.remove("is-active");
  stepReview.classList.add("is-active");
}

function editClass(index) {
  const e = events[index];
  if (!e) return;
  openClassModal({
    mode: "edit",
    index,
    day: normalizeDay(e.day),
    subject: e.subject,
    start_time: e.start_time,
    end_time: e.end_time,
  });
}

function deleteClass(index) {
  pushUndoSnapshot();
  events.splice(index, 1);
  if (!events.length) {
    startOver();
    return;
  }
  display();
  saveDraft();
}

function addClass() {
  openClassModal({ mode: "add", day: "Monday", subject: "", start_time: "09:00", end_time: "09:50" });
}

function addClassForDay(day) {
  openClassModal({ mode: "add", day: normalizeDay(day), subject: "", start_time: "09:00", end_time: "09:50" });
}

function markVerified(btn, index) {
  if (!events[index]) return;
  if (btn) btn.classList.add("verified-out");
  setTimeout(() => {
    pushUndoSnapshot();
    events[index].verified = true;
    if (btn && btn.parentElement) {
      const badge = document.createElement("span");
      badge.className = "confidence-badge verified verified-in";
      badge.textContent = "Verified";
      btn.replaceWith(badge);
      requestAnimationFrame(() => badge.classList.remove("verified-in"));
    }
    saveDraft();
  }, 140);
}

outputEl?.addEventListener("click", (ev) => {
  const btn = ev.target.closest(".verify-btn[data-verify-index]");
  if (!btn) return;
  ev.preventDefault();
  ev.stopPropagation();
  const index = Number(btn.getAttribute("data-verify-index"));
  if (Number.isNaN(index)) return;
  markVerified(btn, index);
});

function startOver() {
  events = [];
  renderEmptyReviewState("No classes detected yet.", "Tip: upload a fresh image and run detect.");
  summaryEl.textContent = "0 classes detected";
  animateStepTransition(false);
  stepUpload.classList.add("is-active");
  stepReview.classList.remove("is-active");
  uploadStateEl.classList.add("hidden");
  undoStack = [];
  updateUndoButton();
  localStorage.removeItem(DRAFT_STORAGE_KEY);
  setStatus("");
}

function openClassModal(config) {
  const mode = config.mode || "add";
  modalTitle.textContent = mode === "edit" ? "Edit Class" : "Add Class";
  modalEditIndex.value = mode === "edit" ? String(config.index) : "";
  modalDay.value = normalizeDay(config.day || "Monday");
  modalSubject.value = config.subject || "";
  modalStart.value = config.start_time || "09:00";
  modalEnd.value = config.end_time || "09:50";
  modalErrorEl.textContent = "";
  modalErrorEl.classList.add("hidden");
  classModalBackdrop.classList.remove("hidden");
  modalSubject.focus();
}

function closeClassModal() {
  classModalBackdrop.classList.add("hidden");
  modalErrorEl.textContent = "";
  modalErrorEl.classList.add("hidden");
}

function closeExportModal() {
  exportModalBackdrop.classList.add("hidden");
}

function confirmExport() {
  if (!pendingIcs) return;
  download(pendingIcs);
  pendingIcs = "";
  closeExportModal();
}

function exportICS() {
  if (!events.length) {
    alert("No classes to export.");
    return;
  }

  let ics = `BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//AutoScan//EN
CALSCALE:GREGORIAN
METHOD:PUBLISH
`;

  const IST_OFFSET_MIN = 330;
  const OFFSET_MS = IST_OFFSET_MIN * 60 * 1000;
  const WEEKDAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  const ICS_DAY = {
    Monday: "MO",
    Tuesday: "TU",
    Wednesday: "WE",
    Thursday: "TH",
    Friday: "FR",
  };

  function pad2(n) {
    return String(n).padStart(2, "0");
  }

  function parseHM(t) {
    const s = String(t || "").trim();
    const m = s.match(/(\d{1,2})\s*[:.]\s*(\d{1,2})/);
    if (!m) return null;
    const h = Number(m[1]);
    const mm = Number(m[2]);
    if (Number.isNaN(h) || Number.isNaN(mm)) return null;
    return { h, mm };
  }

  function toIcsUtc(ms) {
    const d = new Date(ms);
    return (
      d.getUTCFullYear() +
      pad2(d.getUTCMonth() + 1) +
      pad2(d.getUTCDate()) +
      "T" +
      pad2(d.getUTCHours()) +
      pad2(d.getUTCMinutes()) +
      pad2(d.getUTCSeconds())
    );
  }

  function icsEscape(s) {
    return String(s)
      .replace(/\\/g, "\\\\")
      .replace(/\r?\n/g, "\\n")
      .replace(/,/g, "\\,")
      .replace(/;/g, "\\;");
  }

  function uidSafe(s) {
    return String(s).replace(/[^a-zA-Z0-9_-]/g, "_");
  }

  ensureDateDefaults();
  const start = startDateEl?.value ? new Date(`${startDateEl.value}T00:00:00`) : new Date();
  let end = endDateEl?.value ? new Date(`${endDateEl.value}T00:00:00`) : new Date(start.getTime());
  if (!endDateEl?.value) {
    end.setMonth(end.getMonth() + 4);
  }
  if (end <= start) {
    alert("End date must be after start date.");
    return;
  }

  const dtstamp = toIcsUtc(Date.now()) + "Z";
  const reminderMinutes = Number(reminderMinutesEl?.value || 10);
  let recurringCount = 0;
  const untilUtc = (() => {
    const endOfDayUtcMs =
      Date.UTC(end.getFullYear(), end.getMonth(), end.getDate(), 23, 59, 59) - OFFSET_MS;
    return `${toIcsUtc(endOfDayUtcMs)}Z`;
  })();

  const firstDateForDay = (targetDay) => {
    const startDate = new Date(start.getTime());
    const targetIndex = WEEKDAYS.indexOf(targetDay);
    const diff = (targetIndex - startDate.getDay() + 7) % 7;
    startDate.setDate(startDate.getDate() + diff);
    return startDate;
  };

  events.forEach((slot) => {
    const dayName = normalizeDay(slot.day);
    if (!DAY_ORDER.includes(dayName)) return;
    const st = parseHM(slot.start_time);
    const et = parseHM(slot.end_time);
    if (!st || !et) return;

    const first = firstDateForDay(dayName);
    const startUtcMs =
      Date.UTC(first.getFullYear(), first.getMonth(), first.getDate(), st.h, st.mm, 0) - OFFSET_MS;
    const endUtcMs =
      Date.UTC(first.getFullYear(), first.getMonth(), first.getDate(), et.h, et.mm, 0) - OFFSET_MS;
    const uid = `${uidSafe(slot.subject)}-${dayName}-${toIcsUtc(startUtcMs)}@autoscan`;
    const alarmBlock =
      reminderMinutes > 0
        ? `BEGIN:VALARM
TRIGGER:-PT${reminderMinutes}M
ACTION:DISPLAY
DESCRIPTION:Reminder
END:VALARM`
        : "";

    ics += `
BEGIN:VEVENT
UID:${uid}
DTSTAMP:${dtstamp}
SUMMARY:${icsEscape(slot.subject)}
DTSTART:${toIcsUtc(startUtcMs)}Z
DTEND:${toIcsUtc(endUtcMs)}Z
RRULE:FREQ=WEEKLY;BYDAY=${ICS_DAY[dayName]};UNTIL=${untilUtc}
${alarmBlock}
END:VEVENT
`;
    recurringCount += 1;
  });

  ics += "END:VCALENDAR";
  pendingIcs = ics;

  const startLabel = startDateEl?.value || toDateInputValue(start);
  const endLabel = endDateEl?.value || toDateInputValue(end);
  const remindersCount = reminderMinutes > 0 ? recurringCount : 0;
  exportSummaryBody.innerHTML = `
    <p><strong>${recurringCount}</strong> recurring events</p>
    <p><strong>${remindersCount}</strong> reminders (${reminderMinutes > 0 ? `${reminderMinutes} min before` : "none"})</p>
    <p>Date range: <strong>${startLabel}</strong> to <strong>${endLabel}</strong></p>
  `;
  exportModalBackdrop.classList.remove("hidden");
}

function download(data) {
  const blob = new Blob([data], { type: "text/calendar;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "timetable_next4months.ics";
  a.click();
  URL.revokeObjectURL(a.href);
}

function clamp(v, min, max) {
  return Math.max(min, Math.min(max, v));
}

async function preprocessImage(file, options) {
  const img = await createImageBitmap(file);
  const scale = options.pixelImprove ? 1.4 : 1;
  const srcW = Math.max(1, Math.round(img.width * scale));
  const srcH = Math.max(1, Math.round(img.height * scale));
  const rot = ((options.rotationDeg || 0) % 360 + 360) % 360;
  const swap = rot === 90 || rot === 270;
  const w = swap ? srcH : srcW;
  const h = swap ? srcW : srcH;
  const canvas = document.createElement("canvas");
  canvas.width = w;
  canvas.height = h;
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = "high";
  ctx.save();
  ctx.translate(w / 2, h / 2);
  ctx.rotate((rot * Math.PI) / 180);
  ctx.drawImage(img, -srcW / 2, -srcH / 2, srcW, srcH);
  ctx.restore();

  let imageData = ctx.getImageData(0, 0, w, h);
  const data = imageData.data;

  const contrast = options.autoEnhance ? 1.17 : 1;
  const brightness = options.autoEnhance ? 6 : 0;
  const gamma = options.autoEnhance ? 0.96 : 1;

  for (let i = 0; i < data.length; i += 4) {
    let r = data[i];
    let g = data[i + 1];
    let b = data[i + 2];

    if (options.invert) {
      r = 255 - r;
      g = 255 - g;
      b = 255 - b;
    }

    r = clamp((r - 128) * contrast + 128 + brightness, 0, 255);
    g = clamp((g - 128) * contrast + 128 + brightness, 0, 255);
    b = clamp((b - 128) * contrast + 128 + brightness, 0, 255);

    r = clamp(255 * Math.pow(r / 255, gamma), 0, 255);
    g = clamp(255 * Math.pow(g / 255, gamma), 0, 255);
    b = clamp(255 * Math.pow(b / 255, gamma), 0, 255);

    data[i] = r;
    data[i + 1] = g;
    data[i + 2] = b;
  }

  if (options.pixelImprove) {
    imageData = sharpenImageData(imageData, w, h, 0.3);
  }

  ctx.putImageData(imageData, 0, 0);

  const blob = await new Promise((resolve) =>
    canvas.toBlob(resolve, "image/jpeg", 0.95)
  );
  if (!blob) return file;
  return new File([blob], file.name.replace(/\.\w+$/, "") + "-enhanced.jpg", {
    type: "image/jpeg",
  });
}

function sharpenImageData(imageData, w, h, amount) {
  const src = imageData.data;
  const out = new Uint8ClampedArray(src.length);
  const k = [0, -1, 0, -1, 5, -1, 0, -1, 0];

  for (let y = 0; y < h; y += 1) {
    for (let x = 0; x < w; x += 1) {
      const idx = (y * w + x) * 4;
      for (let c = 0; c < 3; c += 1) {
        let acc = 0;
        let ki = 0;
        for (let ky = -1; ky <= 1; ky += 1) {
          for (let kx = -1; kx <= 1; kx += 1) {
            const px = clamp(x + kx, 0, w - 1);
            const py = clamp(y + ky, 0, h - 1);
            const pidx = (py * w + px) * 4 + c;
            acc += src[pidx] * k[ki++];
          }
        }
        out[idx + c] = clamp(src[idx + c] * (1 - amount) + acc * amount, 0, 255);
      }
      out[idx + 3] = src[idx + 3];
    }
  }
  return new ImageData(out, w, h);
}

fileInput.addEventListener("change", () => {
  const file = fileInput.files && fileInput.files[0];
  if (!file) {
    fileInfoEl.textContent = "No file selected";
    uploadPreview.classList.add("hidden");
    scanOverlay.classList.add("hidden");
    uploadSuccessBadge.classList.add("hidden");
    fileInput.closest(".dropzone").classList.remove("has-preview");
    dropContent.classList.remove("hidden");
    if (previewObjectUrl) {
      URL.revokeObjectURL(previewObjectUrl);
      previewObjectUrl = null;
    }
    setStatus("");
    renderEmptyReviewState(
      "No image selected.",
      "Tip: choose a timetable image, then click Detect Timetable."
    );
    return;
  }
  if (previewObjectUrl) {
    URL.revokeObjectURL(previewObjectUrl);
  }
  previewObjectUrl = URL.createObjectURL(file);
  previewImg.src = previewObjectUrl;
  rotationDeg = 0;
  previewImg.style.transform = "scale(1.02) rotate(0deg)";
  setRotationDegreeLabel();
  uploadPreview.classList.remove("hidden");
  fileInput.closest(".dropzone").classList.add("has-preview");
  dropContent.classList.add("hidden");
  fileInfoEl.textContent = `Attached: ${file.name}`;
  setStatus("Image attached. Click Detect Timetable to upload.", "success");
});

ensureDateDefaults();
startDateEl?.addEventListener("change", () => {
  if (!endDateEl?.value) {
    const e = new Date(`${startDateEl.value}T00:00:00`);
    e.setMonth(e.getMonth() + 4);
    endDateEl.value = toDateInputValue(e);
  }
});

function attachDatePickerOpen(input) {
  if (!input) return;
  input.addEventListener("click", () => {
    if (typeof input.showPicker === "function") {
      input.showPicker();
    }
  });
  input.addEventListener("focus", () => {
    if (typeof input.showPicker === "function") {
      input.showPicker();
    }
  });
}

attachDatePickerOpen(startDateEl);
attachDatePickerOpen(endDateEl);

classForm.addEventListener("submit", (e) => {
  e.preventDefault();
  const startMin = toMinutes(modalStart.value);
  const endMin = toMinutes(modalEnd.value);
  if (startMin === null || endMin === null) {
    modalErrorEl.textContent = "Please enter valid start and end time.";
    modalErrorEl.classList.remove("hidden");
    return;
  }
  if (endMin <= startMin) {
    modalErrorEl.textContent = "End time must be after start time.";
    modalErrorEl.classList.remove("hidden");
    return;
  }

  modalErrorEl.textContent = "";
  modalErrorEl.classList.add("hidden");
  const payload = {
    day: normalizeDay(modalDay.value),
    subject: modalSubject.value.trim() || "Untitled",
    start_time: modalStart.value,
    end_time: modalEnd.value,
    confidence: "high",
    verified: true,
    confidence_score: 0.99,
  };
  const indexRaw = modalEditIndex.value;
  pushUndoSnapshot();
  if (indexRaw === "") {
    events.push(payload);
  } else {
    const idx = Number(indexRaw);
    if (!Number.isNaN(idx) && events[idx]) {
      events[idx] = payload;
    }
  }
  closeClassModal();
  display();
  saveDraft();
});

classModalBackdrop.addEventListener("click", (e) => {
  if (e.target === classModalBackdrop) {
    closeClassModal();
  }
});

rotateLeftBtn.addEventListener("click", () => {
  rotationDeg = (rotationDeg + 270) % 360;
  previewImg.style.transform = `scale(1.02) rotate(${rotationDeg}deg)`;
  setRotationDegreeLabel();
});
setRotationDegreeLabel();
updateUndoButton();

function undoLastChange() {
  if (!undoStack.length) return;
  events = undoStack.pop();
  updateUndoButton();
  display();
  saveDraft();
  setStatus("Undid last change.", "success");
}

reminderMinutesEl?.addEventListener("change", saveDraft);
startDateEl?.addEventListener("change", saveDraft);
endDateEl?.addEventListener("change", saveDraft);
enhanceAutoEl?.addEventListener("change", saveDraft);
enhanceInvertEl?.addEventListener("change", saveDraft);
enhancePixelEl?.addEventListener("change", saveDraft);

restoreDraft();
if (!events.length) {
  renderEmptyReviewState(
    "No classes detected yet.",
    "Tip: upload your timetable image and click Detect Timetable."
  );
}
