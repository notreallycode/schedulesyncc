import express from "express";
import multer from "multer";
import cors from "cors";

const app = express();

app.use(cors());
app.use(express.static("public"));

// keep uploads in memory to send as base64
const upload = multer({ storage: multer.memoryStorage() });

const SLOT_TIMES = [
  ["09:00", "09:50"],
  ["09:50", "10:40"],
  ["10:50", "11:40"],
  ["11:40", "12:30"],
  ["13:30", "14:20"],
  ["14:20", "15:10"],
  ["15:10", "16:00"],
  ["16:00", "16:50"],
];

function cleanSubject(s) {
  return String(s)
    .replace(/[|\\]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function isLikelyFreeSlot(token) {
  const t = token.toLowerCase().trim();
  return (
    !t ||
    t === "-" ||
    t === "--" ||
    t === "—" ||
    t === "_" ||
    t === "free" ||
    t === "break"
  );
}

function scoreSubjectConfidence(subject) {
  const s = cleanSubject(subject);
  if (!s) return { score: 0.2, label: "low" };
  let score = 0.85;
  if (s.length < 3) score -= 0.35;
  if (/[^a-zA-Z0-9 .,&()/-]/.test(s)) score -= 0.15;
  if (/^[A-Z]{1,2}$/.test(s.replace(/\s+/g, ""))) score -= 0.15;
  if (score < 0.5) return { score: Math.max(score, 0.2), label: "low" };
  if (score < 0.78) return { score, label: "medium" };
  return { score, label: "high" };
}

function parseTimetableFromText(rawText) {
  const dayMap = {
    monday: "Monday",
    mon: "Monday",
    tuesday: "Tuesday",
    tue: "Tuesday",
    tues: "Tuesday",
    wednesday: "Wednesday",
    wed: "Wednesday",
    thursday: "Thursday",
    thu: "Thursday",
    thurs: "Thursday",
    friday: "Friday",
    fri: "Friday",
  };

  const lines = String(rawText || "")
    .split(/\r?\n/)
    .map((l) => l.replace(/\t/g, " ").trim())
    .filter(Boolean);

  const events = [];

  for (const line of lines) {
    const m = line.match(
      /^(monday|mon|tuesday|tue|tues|wednesday|wed|thursday|thu|thurs|friday|fri)\b[\s:.-]*(.*)$/i
    );
    if (!m) continue;

    const day = dayMap[m[1].toLowerCase()];
    const rest = m[2].trim();
    if (!rest) continue;

    // OCR often preserves larger gaps between cells; use those first.
    let cells = rest
      .split(/\s{2,}|(?=\s[-_—]\s)/)
      .map(cleanSubject)
      .filter(Boolean);

    // Fallback if OCR collapsed spacing: split by single spaces.
    if (cells.length < 2) {
      cells = rest.split(/\s+/).map(cleanSubject).filter(Boolean);
    }

    for (let i = 0; i < cells.length && i < SLOT_TIMES.length; i += 1) {
      const subject = cleanSubject(cells[i]);
      if (isLikelyFreeSlot(subject)) continue;
      let confidence = scoreSubjectConfidence(subject);

      events.push({
        day,
        subject,
        start_time: SLOT_TIMES[i][0],
        end_time: SLOT_TIMES[i][1],
        confidence: confidence.label,
        confidence_score: Number(confidence.score.toFixed(2)),
      });
    }
  }

  return events;
}

function parseTimetableFromOverlay(ocrJson) {
  const dayMap = {
    monday: "Monday",
    mon: "Monday",
    tuesday: "Tuesday",
    tue: "Tuesday",
    tues: "Tuesday",
    wednesday: "Wednesday",
    wed: "Wednesday",
    thursday: "Thursday",
    thu: "Thursday",
    thurs: "Thursday",
    friday: "Friday",
    fri: "Friday",
  };

  const parsed = Array.isArray(ocrJson?.ParsedResults) ? ocrJson.ParsedResults : [];
  const lines = parsed.flatMap((p) => p?.TextOverlay?.Lines || []);
  if (!lines.length) return [];

  const normalizedLines = lines
    .map((line) => {
      const words = Array.isArray(line?.Words) ? line.Words : [];
      const text = cleanSubject(words.map((w) => w?.WordText || "").join(" "));
      const left = words.length ? Math.min(...words.map((w) => Number(w.Left) || 0)) : Number(line?.MinLeft) || 0;
      const right = words.length
        ? Math.max(...words.map((w) => (Number(w.Left) || 0) + (Number(w.Width) || 0)))
        : left + (Number(line?.MaxWidth) || 0);
      const top = Number(line?.MinTop) || 0;
      const height = Number(line?.MaxHeight) || 0;
      return { text, left, right, top, height, words };
    })
    .filter((l) => l.text);

  const anchors = [];
  for (const line of normalizedLines) {
    const m = line.text.match(/^(monday|mon|tuesday|tue|tues|wednesday|wed|thursday|thu|thurs|friday|fri)\b/i);
    if (!m) continue;
    anchors.push({
      day: dayMap[m[1].toLowerCase()],
      top: line.top,
      height: line.height || 24,
      right: line.right,
      raw: line,
    });
  }
  if (!anchors.length) return [];

  anchors.sort((a, b) => a.top - b.top);
  const dayAnchors = [];
  for (const a of anchors) {
    if (!dayAnchors.some((d) => d.day === a.day)) dayAnchors.push(a);
  }
  dayAnchors.sort((a, b) => a.top - b.top);

  const dayColumnRight = Math.max(...dayAnchors.map((a) => a.right), 100);
  const candidateLeftThreshold = dayColumnRight + 16;

  const events = [];
  for (let i = 0; i < dayAnchors.length; i += 1) {
    const current = dayAnchors[i];
    const prev = dayAnchors[i - 1];
    const next = dayAnchors[i + 1];
    const rowTop = prev ? Math.round((prev.top + current.top) / 2) : current.top - 14;
    const rowBottom = next ? Math.round((current.top + next.top) / 2) : current.top + Math.max(current.height, 30) + 80;

    const rowLines = normalizedLines.filter(
      (l) =>
        l.top >= rowTop &&
        l.top < rowBottom &&
        l.left >= candidateLeftThreshold &&
        !/^(monday|mon|tuesday|tue|tues|wednesday|wed|thursday|thu|thurs|friday|fri)\b/i.test(l.text)
    );
    if (!rowLines.length) continue;

    const minX = Math.min(...rowLines.map((l) => l.left));
    const maxX = Math.max(...rowLines.map((l) => l.right));
    const width = Math.max(maxX - minX, 1);
    const slotWidth = width / SLOT_TIMES.length;
    const slotTexts = Array.from({ length: SLOT_TIMES.length }, () => []);

    for (const line of rowLines) {
      const center = (line.left + line.right) / 2;
      let slotIdx = Math.floor((center - minX) / slotWidth);
      if (slotIdx < 0) slotIdx = 0;
      if (slotIdx >= SLOT_TIMES.length) slotIdx = SLOT_TIMES.length - 1;
      slotTexts[slotIdx].push(line.text);
    }

    for (let slotIdx = 0; slotIdx < SLOT_TIMES.length; slotIdx += 1) {
      const subject = cleanSubject(Array.from(new Set(slotTexts[slotIdx])).join(" / "));
      if (isLikelyFreeSlot(subject)) continue;
      const confidence = scoreSubjectConfidence(subject);
      // Slightly lower confidence when slot had fragmented/very sparse OCR lines.
      if (slotTexts[slotIdx].length <= 1 && confidence.label === "high") {
        confidence.label = "medium";
        confidence.score = Math.min(confidence.score, 0.74);
      }
      events.push({
        day: current.day,
        subject,
        start_time: SLOT_TIMES[slotIdx][0],
        end_time: SLOT_TIMES[slotIdx][1],
        confidence: confidence.label,
        confidence_score: Number(confidence.score.toFixed(2)),
      });
    }
  }

  return events;
}

app.post("/detect", upload.single("image"), async (req, res) => {
  try {
    if (!process.env.OCRSPACE_API_KEY) {
      return res.status(500).json({ error: "Missing OCRSPACE_API_KEY env var" });
    }

    if (!req.file || !req.file.buffer) {
      return res.status(400).json({ error: "Missing image upload" });
    }

    const mimeType = req.file.mimetype || "image/jpeg";
    const imageBlob = new Blob([req.file.buffer], { type: mimeType });

    const requestOcr = async ({ engine, isTable, overlay }) => {
      const form = new FormData();
      form.append("file", imageBlob, req.file.originalname || "timetable.jpg");
      form.append("language", "eng");
      form.append("isTable", isTable ? "true" : "false");
      form.append("isOverlayRequired", overlay ? "true" : "false");
      form.append("detectOrientation", "true");
      form.append("OCREngine", String(engine));
      form.append("scale", "true");

      const ocrRes = await fetch("https://api.ocr.space/parse/image", {
        method: "POST",
        headers: { apikey: process.env.OCRSPACE_API_KEY },
        body: form,
      });
      const ocrJson = await ocrRes.json().catch(() => null);
      return { ocrRes, ocrJson };
    };

    const extractEvents = (ocrJson) => {
      const rawText = (ocrJson?.ParsedResults || [])
        .map((p) => p?.ParsedText || "")
        .join("\n");
      let events = parseTimetableFromOverlay(ocrJson);
      if (!events.length) events = parseTimetableFromText(rawText);
      return events;
    };

    const primary = await requestOcr({ engine: 2, isTable: true, overlay: true });
    if (!primary.ocrRes.ok || !primary.ocrJson) {
      return res.status(502).json({
        error: "OCR.space request failed",
        details: `HTTP ${primary.ocrRes.status}`,
      });
    }
    if (primary.ocrJson.IsErroredOnProcessing) {
      const msg = Array.isArray(primary.ocrJson.ErrorMessage)
        ? primary.ocrJson.ErrorMessage.join(" | ")
        : String(primary.ocrJson.ErrorMessage || "OCR processing error");
      return res.status(422).json({ error: "OCR.space could not read image", details: msg });
    }

    let events = extractEvents(primary.ocrJson);

    // Reliability fallback: retry with alternate strategy when first parse is sparse.
    if (events.length < 4) {
      const fallback = await requestOcr({ engine: 1, isTable: false, overlay: true });
      if (fallback.ocrRes.ok && fallback.ocrJson && !fallback.ocrJson.IsErroredOnProcessing) {
        const fallbackEvents = extractEvents(fallback.ocrJson);
        if (fallbackEvents.length > events.length) {
          events = fallbackEvents;
        }
      }
    }

    if (events.length) {
      // Deduplicate repeated OCR fragments for same day/time.
      const seen = new Set();
      events = events.filter((e) => {
        const key = `${e.day}|${e.start_time}|${e.end_time}|${e.subject.toLowerCase()}`;
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
    }

    if (!events.length) {
      return res.status(422).json({
        error: "Could not parse classes from OCR text. Try a clearer photo.",
      });
    }

    res.json(events);
  } catch (e) {
    console.error(e);
    res.status(500).json({
      error: "OCR.space OCR failed",
      details: String(e?.message ?? e),
    });
  }
});

app.listen(3000, () => console.log("running with OCR.space OCR"));