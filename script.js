/* ===================================================================
   FUNNEL VELOCITY DASHBOARD — data loader
   ===================================================================

   DATA SOURCE
   Pulls live data from your FONO_Acquirer_Tracker Google Sheet using
   the public "gviz" JSON endpoint — no backend / Apps Script / API
   key required. The sheet must be shared as "Anyone with the link —
   Viewer" (File > Share > General access).

   Paste the Sheet ID (the string in the URL between /d/ and /edit)
   into CONFIG.SHEET_ID below.

   TABS USED
     - "Acquirer Performance"  → FTD block + MTD block, per acquirer
     - "Visit Log"             → raw dated events, used for the weekly
                                  sparkline AND for any custom date
                                  range that isn't exactly "this month"
                                  or the FTD preset.

   TIMEFRAME MODES
     - MTD  ("This Month" preset) → read straight from the sheet's
       "MTD Performance" block. These are your validated, reconciled
       numbers — not recomputed.
     - FTD  ("All Time" preset)   → read straight from the sheet's
       "FTD Performance" block (assumed = cumulative since inception;
       flagged for confirmation, see README note below).
     - Custom range (any dates picked on the calendar that aren't
       exactly one of the above) → computed live from Visit Log by
       filtering events to the picked range and re-bucketing by
       Stage After. If the picked range happens to exactly match the
       current calendar month, it silently snaps back to the MTD
       block instead of recomputing, so you always get the validated
       sheet numbers when it really is "this month".

   STAGE MAPPING (Visit Log "Stage After" → funnel columns), confirmed:
     Visited                        → Visits
     Lead                           → Pipeline
     Signed (LOI)                   → Contracting
     Onboarded (Live) / (Takeover)  → Contracted
   "Lost/Dropped" and "Stalled" are excluded (outcomes, not stages).

   ASSUMPTIONS FLAGGED FOR REVIEW
   - FTD is treated as "cumulative since inception" — confirm this
     matches your terminology.
   - For custom ranges, a prospect is counted once per funnel stage
     within the picked range (first log entry for that stage in that
     window), so repeat updates don't inflate nest counts.
   - "Target" is always the fixed Target Count column from the sheet
     — it does not scale down for shorter custom ranges (a 3-day
     range is still compared against the full target).
   - "Properties signed" = count of unique prospects reaching
     "Contracted" (MTD/FTD: the sheet's own Contracted Count column;
     custom range: unique prospects in range).
   ================================================================ */

const CONFIG = {
  // Paste your Google Sheet ID here (from the sheet's URL).
  SHEET_ID: "https://docs.google.com/spreadsheets/d/1pZNwOip3teKUuV2bKQGuKnLMikKl5DVieNGVPBtEzqE/edit?gid=2059008135#gid=2059008135",

  TABS: {
    acquirerPerformance: "Acquirer Performance",
    visitLog: "Visit Log",
  },

  // Re-fetch the sheet every N milliseconds so the dashboard stays
  // live without a page refresh. Set to 0 to disable auto-refresh.
  REFRESH_MS: 60000,
};

const THEATRE_ORDER = ["RN", "CORO", "WLG", "DCN"];
const THEATRE_NAMES = { RN: "RN Theatre", CORO: "CORO Theatre", WLG: "WLG Theatre", DCN: "DCN Theatre" };

const STAGE_MAP = {
  "visited": "visits",
  "lead": "pipeline",
  "signed (loi)": "contracting",
  "onboarded (live)": "contracted",
  "onboarded (takeover)": "contracted",
};

const MTD_FIELD_MAP = {
  visitNests: "Visit Nests",
  pipelineNest: "Pipeline Nest",
  contractingNest: "Contracting Nest",
  contractedNest: "Contracted Nest",
  contractedCount: "Contracted Count",
};

const FTD_FIELD_MAP = {
  visitNests: "Visit Nests",
  pipelineNest: "Funnel Nest",
  contractingNest: "Contracting Nest",
  contractedNest: "Contracted Nest",
  contractedCount: "Contracted Count",
};

/* ---------------------------------------------------------------
   Google Sheets gviz raw fetch (no header assumptions — returns a
   plain 2D array of cell values so each tab can be parsed by
   scanning for its own header text).
   --------------------------------------------------------------- */

function gvizUrl(sheetName) {
  return `https://docs.google.com/spreadsheets/d/${CONFIG.SHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(sheetName)}`;
}

function parseGvizCellValue(cell) {
  if (!cell) return "";
  const v = cell.v;
  if (v === null || v === undefined) return "";
  // Date / DateTime cells come back as "Date(2026,5,6)" (0-indexed month)
  if (typeof v === "string" && v.startsWith("Date(")) {
    const parts = v.slice(5, -1).split(",").map((n) => parseInt(n, 10));
    const [y, mo, d = 1, h = 0, mi = 0, s = 0] = parts;
    return new Date(y, mo, d, h, mi, s);
  }
  return v;
}

async function fetchRawRows(sheetName) {
  const res = await fetch(gvizUrl(sheetName));
  if (!res.ok) {
    throw new Error(`Could not load tab "${sheetName}" (HTTP ${res.status}). Check the tab name and sharing settings.`);
  }
  const text = await res.text();
  const jsonStart = text.indexOf("{");
  const jsonEnd = text.lastIndexOf("}");
  if (jsonStart === -1 || jsonEnd === -1) {
    throw new Error(`Unexpected response for tab "${sheetName}". Is the sheet ID correct and shared publicly?`);
  }
  const parsed = JSON.parse(text.slice(jsonStart, jsonEnd + 1));
  if (parsed.status === "error") {
    const msg = (parsed.errors && parsed.errors[0] && parsed.errors[0].detailed_message) || "Unknown sheet error";
    throw new Error(`Tab "${sheetName}": ${msg}`);
  }
  return parsed.table.rows.map((r) => r.c.map(parseGvizCellValue));
}

function toNumber(v, fallback = 0) {
  if (v === null || v === undefined || v === "") return fallback;
  if (v instanceof Date) return fallback;
  const n = Number(String(v).replace(/[, ]/g, ""));
  return Number.isFinite(n) ? n : fallback;
}

function norm(v) {
  return String(v || "").trim().replace(/\s+/g, " ").toLowerCase();
}

function toMidnight(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function fmtDateInput(d) {
  // yyyy-mm-dd for <input type="date">
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function fmtDateLabel(d) {
  return d.toLocaleDateString("default", { month: "short", day: "numeric", year: "numeric" });
}

/* ---------------------------------------------------------------
   Parse "Acquirer Performance" — merged multi-row header. Reads
   either the "FTD Performance" or "MTD Performance" block by
   locating that section's label column, then finding each field's
   header *starting from* that column, so duplicate column names in
   the other section never get matched by accident.
   --------------------------------------------------------------- */

function findSectionStart(rows, label) {
  for (let r = 0; r < Math.min(rows.length, 6); r++) {
    const idx = rows[r].findIndex((c) => norm(c).includes(norm(label)));
    if (idx !== -1) return { col: idx, headerRowIdx: r + 1 };
  }
  return null;
}

function parsePerformanceSection(rows, sectionLabel, fieldMap) {
  const sec = findSectionStart(rows, sectionLabel);
  if (!sec) {
    throw new Error(`Couldn't find a "${sectionLabel}" section header in "${CONFIG.TABS.acquirerPerformance}".`);
  }
  const headerRow = rows[sec.headerRowIdx] || [];
  const findCol = (label) => headerRow.findIndex((c, i) => i >= sec.col && norm(c) === norm(label));

  const targetIdx = headerRow.findIndex((c) => norm(c) === norm("Target Count"));
  const nameIdx = 0;
  const theatreIdx = 1;

  const cols = {};
  Object.entries(fieldMap).forEach(([key, label]) => { cols[key] = findCol(label); });

  const missing = Object.entries(cols).filter(([, i]) => i === -1).map(([key]) => fieldMap[key]);
  if (missing.length) {
    throw new Error(`Couldn't find column(s) ${missing.join(", ")} in the ${sectionLabel} block of "${CONFIG.TABS.acquirerPerformance}".`);
  }

  const acquirers = [];
  for (let r = sec.headerRowIdx + 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const name = row[nameIdx];
    const theatre = row[theatreIdx];
    // Real acquirer rows have both a name AND a theatre code populated.
    // Section headers ("RN Theatre") and totals ("RN Total") have one
    // or the other blank, so this filters them out automatically.
    if (!name || !theatre) continue;
    const rec = {
      name: String(name).trim(),
      theatre: String(theatre).trim().toUpperCase(),
      target: toNumber(row[targetIdx]),
    };
    Object.keys(fieldMap).forEach((key) => { rec[key] = toNumber(row[cols[key]]); });
    acquirers.push(rec);
  }
  return acquirers;
}

/* ---------------------------------------------------------------
   Parse "Visit Log" — every row, mapped to a funnel stage (or null
   for terminal / non-funnel outcomes like Lost/Dropped or Stalled).
   Returned as a flat event list so it can be filtered to ANY date
   range picked on the calendar.
   --------------------------------------------------------------- */

function parseVisitLogAll(rows) {
  let headerRowIdx = -1;
  for (let r = 0; r < Math.min(rows.length, 6); r++) {
    const rowNorm = rows[r].map(norm);
    if (rowNorm.includes("date") && rowNorm.includes("stage after")) { headerRowIdx = r; break; }
  }
  if (headerRowIdx === -1) {
    throw new Error(`Couldn't find the header row (expecting "Date" and "Stage After" columns) in "${CONFIG.TABS.visitLog}".`);
  }
  const headerRow = rows[headerRowIdx];
  const findCol = (label) => headerRow.findIndex((c) => norm(c) === norm(label));

  const dateIdx = findCol("Date");
  const theatreIdx = findCol("Theatre");
  const acquirerIdx = findCol("Acquirer");
  const prospectIdx = findCol("Prospect (PG / owner)") !== -1 ? findCol("Prospect (PG / owner)") : findCol("Prospect");
  const stageAfterIdx = findCol("Stage After");
  const nestsIdx = findCol("Nests Potential");

  if ([dateIdx, theatreIdx, stageAfterIdx, nestsIdx].includes(-1)) {
    throw new Error(`Missing expected column(s) in "${CONFIG.TABS.visitLog}" (Date / Theatre / Stage After / Nests Potential).`);
  }

  const events = [];
  for (let r = headerRowIdx + 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || !row[dateIdx] || !(row[dateIdx] instanceof Date)) continue;
    events.push({
      date: toMidnight(row[dateIdx]),
      theatre: String(row[theatreIdx] || "").trim().toUpperCase(),
      acquirer: acquirerIdx !== -1 ? String(row[acquirerIdx] || "").trim() : "",
      prospect: prospectIdx !== -1 ? String(row[prospectIdx] || "").trim() : `row${r}`,
      funnelStage: STAGE_MAP[norm(row[stageAfterIdx])] || null,
      nests: toNumber(row[nestsIdx]),
    });
  }
  return events;
}

/* ---------------------------------------------------------------
   Custom-range aggregation from Visit Log events
   --------------------------------------------------------------- */

function aggregateEventsForRange(events, start, end) {
  // key: theatre||acquirer||stage -> Map(prospect -> nests)
  const buckets = new Map();
  events.forEach((e) => {
    if (!e.date || !e.funnelStage) return;
    if (e.date < start || e.date > end) return;
    const key = `${e.theatre}||${e.acquirer}||${e.funnelStage}`;
    if (!buckets.has(key)) buckets.set(key, new Map());
    const m = buckets.get(key);
    if (!m.has(e.prospect)) m.set(e.prospect, e.nests);
  });

  // key: theatre||acquirer -> { visitNests, pipelineNest, contractingNest, contractedNest, contractedCount }
  const result = new Map();
  buckets.forEach((prospectMap, key) => {
    const [theatre, acquirer, stage] = key.split("||");
    const sum = [...prospectMap.values()].reduce((a, b) => a + b, 0);
    const rKey = `${theatre}||${acquirer}`;
    if (!result.has(rKey)) result.set(rKey, { visitNests: 0, pipelineNest: 0, contractingNest: 0, contractedNest: 0, contractedCount: 0 });
    const obj = result.get(rKey);
    if (stage === "visits") obj.visitNests += sum;
    else if (stage === "pipeline") obj.pipelineNest += sum;
    else if (stage === "contracting") obj.contractingNest += sum;
    else if (stage === "contracted") { obj.contractedNest += sum; obj.contractedCount += prospectMap.size; }
  });
  return result;
}

function buildAcquirersForRange(mode, start, end, ctx) {
  if (mode === "mtd") return ctx.mtdAcquirers;
  if (mode === "ftd") return ctx.ftdAcquirers;

  // custom
  const agg = aggregateEventsForRange(ctx.events, start, end);
  return ctx.metaList.map((meta) => {
    const hit = agg.get(`${meta.theatre}||${meta.name}`);
    return {
      name: meta.name,
      theatre: meta.theatre,
      target: meta.target,
      visitNests: hit ? hit.visitNests : 0,
      pipelineNest: hit ? hit.pipelineNest : 0,
      contractingNest: hit ? hit.contractingNest : 0,
      contractedNest: hit ? hit.contractedNest : 0,
      contractedCount: hit ? hit.contractedCount : 0,
    };
  });
}

/* ---------------------------------------------------------------
   Weekly trend (sparkline) — buckets Visit-stage events within the
   picked range into up to 5 cumulative buckets, so it adapts to
   ranges shorter or longer than a calendar month.
   --------------------------------------------------------------- */

function computeTrend(events, start, end) {
  const inRange = events.filter((e) => e.funnelStage === "visits" && e.date >= start && e.date <= end);

  // Dedup: earliest "Visited" event per Theatre+Prospect within range
  const seen = new Map();
  inRange.forEach((e) => {
    const key = `${e.theatre}::${e.prospect}`;
    const existing = seen.get(key);
    if (!existing || e.date < existing.date) seen.set(key, e);
  });
  const uniqueVisits = [...seen.values()];

  const byTheatre = {};
  THEATRE_ORDER.forEach((t) => { byTheatre[t] = []; });
  if (!uniqueVisits.length) {
    const trend = {};
    THEATRE_ORDER.forEach((t) => { trend[t] = {}; });
    return trend;
  }

  const msDay = 86400000;
  const totalDays = Math.max(1, Math.round((end - start) / msDay) + 1);
  const bucketCount = Math.min(5, Math.max(2, Math.ceil(totalDays / 7)));
  const bucketSize = Math.max(1, Math.ceil(totalDays / bucketCount));
  const bucketOf = (d) => Math.min(bucketCount - 1, Math.floor((d - start) / msDay / bucketSize));

  const counts = {};
  THEATRE_ORDER.forEach((t) => { counts[t] = new Array(bucketCount).fill(0); });
  uniqueVisits.forEach((e) => {
    if (!counts[e.theatre]) counts[e.theatre] = new Array(bucketCount).fill(0);
    counts[e.theatre][bucketOf(e.date)] += e.nests;
  });

  const trend = {};
  Object.entries(counts).forEach(([theatre, buckets]) => {
    let running = 0;
    const cumulative = buckets.map((b) => (running += b));
    const labeled = {};
    cumulative.forEach((val, i) => {
      const label = i === cumulative.length - 1 ? "Now" : `W${i + 1}`;
      labeled[label] = val;
    });
    trend[theatre] = labeled;
  });
  return trend;
}

/* ---------------------------------------------------------------
   Roll acquirer-level rows up into theatre-level + summary figures
   --------------------------------------------------------------- */

function buildDashboardData(acquirers, trendByTheatre, timeframe) {
  const theatresPresent = [...new Set(acquirers.filter((a) => a.visitNests || a.pipelineNest || a.contractingNest || a.contractedNest || a.target).map((a) => a.theatre))];
  const allSeenTheatres = [...new Set(acquirers.map((a) => a.theatre))];
  const orderedTheatres = THEATRE_ORDER.filter((t) => allSeenTheatres.includes(t))
    .concat(allSeenTheatres.filter((t) => !THEATRE_ORDER.includes(t)));

  const theatreRows = orderedTheatres.map((code) => {
    const rows = acquirers.filter((a) => a.theatre === code);
    const sum = (key) => rows.reduce((acc, a) => acc + a[key], 0);
    const visits = sum("visitNests");
    const pipeline = sum("pipelineNest");
    const contractG = sum("contractingNest");
    const contracted = sum("contractedNest");

    const topByContracted = [...rows].sort((a, b) => b.contractedNest - a.contractedNest)[0];
    const conv = visits > 0 ? Math.round((pipeline / visits) * 100) : 0;
    let insight;
    if (contracted === 0 && contractG === 0 && visits === 0) {
      insight = `No activity logged for ${rows.length} acquirer${rows.length === 1 ? "" : "s"} in this timeframe.`;
    } else if (contracted === 0 && contractG === 0) {
      insight = `${rows.length} acquirer${rows.length === 1 ? "" : "s"} active. ${conv}% visit\u2192pipeline rate. No contracting or contracted activity yet in this timeframe.`;
    } else if (topByContracted && topByContracted.contractedNest > 0) {
      insight = `${conv}% visit\u2192pipeline rate. <b>${topByContracted.name}</b> leads with ${topByContracted.contractedNest.toLocaleString()} contracted nests across ${rows.length} acquirer${rows.length === 1 ? "" : "s"}.`;
    } else {
      insight = `${conv}% visit\u2192pipeline rate across ${rows.length} acquirer${rows.length === 1 ? "" : "s"}.`;
    }

    return {
      Theatre: THEATRE_NAMES[code] || `${code} Theatre`,
      code,
      Visits: visits, Pipeline: pipeline, ContractG: contractG, Contracted: contracted,
      Insight: insight,
    };
  });

  const acquirerRows = acquirers.map((a) => {
    const stages = [
      { type: "Pipe", value: a.pipelineNest },
      { type: "Cont'g", value: a.contractingNest },
      { type: "Cont.", value: a.contractedNest },
    ].filter((s) => s.value > 0);
    const shown = stages.slice(-2);
    return {
      Theatre: THEATRE_NAMES[a.theatre] || `${a.theatre} Theatre`,
      Name: a.name,
      Metric1Type: shown[0] ? shown[0].type : "",
      Metric1Value: shown[0] ? shown[0].value : "",
      Metric2Type: shown[1] ? shown[1].type : "",
      Metric2Value: shown[1] ? shown[1].value : "",
    };
  });

  const trendRows = orderedTheatres.map((code) => ({
    Theatre: THEATRE_NAMES[code] || `${code} Theatre`,
    ...(trendByTheatre[code] || {}),
  }));

  const totalVisits = theatreRows.reduce((a, t) => a + t.Visits, 0);
  const totalPipeline = theatreRows.reduce((a, t) => a + t.Pipeline, 0);
  const totalContracted = theatreRows.reduce((a, t) => a + t.Contracted, 0);
  const totalContractedProps = acquirers.reduce((a, x) => a + x.contractedCount, 0);
  const totalTarget = acquirers.reduce((a, x) => a + x.target, 0);

  const leadTheatre = [...theatreRows].sort((a, b) => b.Visits - a.Visits)[0];
  const summaryNarrative = leadTheatre && leadTheatre.Visits > 0
    ? `<b>${leadTheatre.Theatre.replace(" Theatre", "")}</b> leads with the highest visit volume (${leadTheatre.Visits.toLocaleString()} nests) ${timeframe.narrativeSuffix}. ${theatreRows.length} theatres, ${acquirers.length} acquirers in view.`
    : `No visit activity recorded ${timeframe.narrativeSuffix}.`;

  const summary = [
    { Key: "TimeframeLabel", Value: timeframe.label },
    { Key: "TimeframeDates", Value: timeframe.dates },
    { Key: "SummaryNarrative", Value: summaryNarrative },
    { Key: "TotalVisitNests", Value: totalVisits },
    { Key: "TotalTheatres", Value: theatreRows.length },
    { Key: "TotalAcquirers", Value: acquirers.length },
    { Key: "PipelineConversionPct", Value: totalVisits > 0 ? Math.round((totalPipeline / totalVisits) * 100) : 0 },
    { Key: "ContractedNests", Value: totalContracted },
    { Key: "ContractedProperties", Value: totalContractedProps },
    { Key: "TargetNests", Value: totalTarget },
  ];

  return { summary, theatres: theatreRows, acquirers: acquirerRows, trend: trendRows };
}

/* ---------------------------------------------------------------
   Rendering
   --------------------------------------------------------------- */

const PILL_CLASS = {
  "Cont'g": "pill--orange",
  "Pipe": "pill--blue",
  "Cont.": "pill--green",
};

function renderSummary(summaryRows) {
  const kv = {};
  summaryRows.forEach((r) => { kv[r.Key] = r.Value; });

  document.getElementById("timeframe-pill").textContent = kv.TimeframeLabel || "";
  document.getElementById("timeframe-dates").textContent = kv.TimeframeDates || "";
  document.getElementById("footer-mid").textContent = `${kv.TimeframeLabel || ""} \u00b7 ${kv.TimeframeDates || ""}`;

  document.getElementById("summary-narrative").innerHTML = kv.SummaryNarrative || "";

  const totalNests = toNumber(kv.TotalVisitNests);
  document.getElementById("stat-total-nests").textContent = totalNests.toLocaleString();
  document.getElementById("stat-total-caption").textContent =
    `Across ${toNumber(kv.TotalTheatres)} Theatres / ${toNumber(kv.TotalAcquirers)} Acquirers`;

  document.getElementById("stat-pipeline-conv").textContent = `${toNumber(kv.PipelineConversionPct)}%`;

  const contracted = toNumber(kv.ContractedNests);
  document.getElementById("stat-contracted").textContent = contracted.toLocaleString();
  document.getElementById("stat-contracted-caption").textContent =
    `${toNumber(kv.ContractedProperties)} Properties signed`;

  document.getElementById("target-pipeline").innerHTML = `${totalNests.toLocaleString()} <span>Nests</span>`;

  const target = toNumber(kv.TargetNests, 1);
  const pct = target > 0 ? Math.round((contracted / target) * 100) : 0;
  document.getElementById("target-pct").textContent = `${pct}%`;
  document.getElementById("target-fraction").textContent =
    `${contracted.toLocaleString()} / ${target.toLocaleString()} Target`;
  document.getElementById("target-bar-fill").style.width = `${Math.min(pct, 100)}%`;
}

function renderTheatres(theatreRows, acquirerRows, trendRows) {
  const container = document.getElementById("theatre-columns");
  container.innerHTML = "";
  const template = document.getElementById("theatre-template");

  theatreRows.forEach((t) => {
    const node = template.content.cloneNode(true);

    const visits = t.Visits, pipeline = t.Pipeline, contractG = t.ContractG, contracted = t.Contracted;

    node.querySelector(".theatre-name").textContent = t.Theatre;
    node.querySelector(".theatre-visits").textContent = `${visits.toLocaleString()} Visit Nests`;

    const stageValues = { visits, pipeline, contracting: contractG, contracted };
    Object.entries(stageValues).forEach(([stage, value]) => {
      const row = node.querySelector(`[data-stage="${stage}"]`);
      row.querySelector(".funnel-value").textContent = value.toLocaleString();
      row.classList.toggle("is-zero", value === 0);
    });

    // Sparkline
    const trend = trendRows.find((r) => r.Theatre === t.Theatre);
    const sparkEl = node.querySelector(".sparkline");
    if (trend) {
      const weekCols = Object.keys(trend).filter((k) => k !== "Theatre");
      if (weekCols.length) {
        const values = weekCols.map((c) => toNumber(trend[c]));
        const max = Math.max(...values, 1);
        weekCols.forEach((c) => {
          const bar = document.createElement("div");
          bar.className = "bar";
          const h = Math.max((toNumber(trend[c]) / max) * 100, 4);
          bar.style.height = `${h}%`;
          sparkEl.appendChild(bar);
        });
      }
    }

    // Acquirers
    const acquirers = acquirerRows.filter((r) => r.Theatre === t.Theatre);
    const active = acquirers.filter((a) => a.Metric1Type || a.Metric2Type).length;
    node.querySelector(".acquirer-active").textContent = `${active} / ${acquirers.length} Active`;

    const list = node.querySelector(".acquirer-list");
    const zeroNames = [];
    acquirers.forEach((a) => {
      if (!a.Metric1Type && !a.Metric2Type) {
        zeroNames.push(a.Name);
        return;
      }
      const li = document.createElement("li");
      li.className = "acquirer-row";
      const pills = [];
      if (a.Metric1Type) {
        pills.push(`<span class="pill ${PILL_CLASS[a.Metric1Type] || "pill--blue"}">${toNumber(a.Metric1Value).toLocaleString()} ${a.Metric1Type}</span>`);
      }
      if (a.Metric2Type) {
        pills.push(`<span class="pill ${PILL_CLASS[a.Metric2Type] || "pill--green"}">${toNumber(a.Metric2Value).toLocaleString()} ${a.Metric2Type}</span>`);
      }
      li.innerHTML = `<span class="acquirer-name">${a.Name}</span><span class="acquirer-pills">${pills.join("")}</span>`;
      list.appendChild(li);
    });

    const zeroEl = node.querySelector(".acquirer-zero");
    if (zeroNames.length) {
      zeroEl.textContent = `${zeroNames.join(", ")} \u00b7 0 activity`;
    } else {
      zeroEl.remove();
    }

    node.querySelector(".theatre-insight").innerHTML = t.Insight || "";

    container.appendChild(node);
  });
}

/* ---------------------------------------------------------------
   Demo data — shown automatically until a real SHEET_ID is set,
   so the dashboard is never blank while you wire up the sheet.
   The filter controls are disabled in demo mode.
   --------------------------------------------------------------- */

const DEMO_DATA = {
  summary: [
    { Key: "TimeframeLabel", Value: "MTD (This Month)" },
    { Key: "TimeframeDates", Value: "Jun 1 \u2013 Jun 9, 2026" },
    { Key: "SummaryNarrative", Value: "<b>RN</b> Theatre leads with the highest visit volume. Demo data shown \u2014 connect your sheet to see live numbers." },
    { Key: "TotalVisitNests", Value: 2294 },
    { Key: "TotalTheatres", Value: 4 },
    { Key: "TotalAcquirers", Value: 11 },
    { Key: "PipelineConversionPct", Value: 0 },
    { Key: "ContractedNests", Value: 299 },
    { Key: "ContractedProperties", Value: 6 },
    { Key: "TargetNests", Value: 5750 },
  ],
  theatres: [
    { Theatre: "RN Theatre", Visits: 1628, Pipeline: 0, ContractG: 1, Contracted: 195, Insight: "Demo data \u2014 connect your sheet." },
    { Theatre: "CORO Theatre", Visits: 194, Pipeline: 0, ContractG: 0, Contracted: 52, Insight: "Demo data \u2014 connect your sheet." },
    { Theatre: "WLG Theatre", Visits: 472, Pipeline: 0, ContractG: 1, Contracted: 52, Insight: "Demo data \u2014 connect your sheet." },
    { Theatre: "DCN Theatre", Visits: 0, Pipeline: 0, ContractG: 0, Contracted: 0, Insight: "Demo data \u2014 connect your sheet." },
  ],
  acquirers: [
    { Theatre: "RN Theatre", Name: "Sandeep", Metric1Type: "Cont'g", Metric1Value: 40, Metric2Type: "Cont.", Metric2Value: 167 },
    { Theatre: "RN Theatre", Name: "Karan", Metric1Type: "Cont'g", Metric1Value: 40, Metric2Type: "Cont.", Metric2Value: 28 },
    { Theatre: "RN Theatre", Name: "Sagar", Metric1Type: "", Metric1Value: "", Metric2Type: "", Metric2Value: "" },
    { Theatre: "CORO Theatre", Name: "Muruganandam", Metric1Type: "Cont.", Metric1Value: 52, Metric2Type: "", Metric2Value: "" },
    { Theatre: "WLG Theatre", Name: "Srinivas", Metric1Type: "Cont.", Metric1Value: 52, Metric2Type: "", Metric2Value: "" },
    { Theatre: "WLG Theatre", Name: "Surya", Metric1Type: "Cont'g", Metric1Value: 150, Metric2Type: "Cont.", Metric2Value: 52 },
    { Theatre: "DCN Theatre", Name: "Vijay Arude", Metric1Type: "", Metric1Value: "", Metric2Type: "", Metric2Value: "" },
  ],
  trend: [
    { Theatre: "RN Theatre", W1: 400, W2: 900, Now: 1628 },
    { Theatre: "CORO Theatre", W1: 60, W2: 130, Now: 194 },
    { Theatre: "WLG Theatre", W1: 90, W2: 260, Now: 472 },
    { Theatre: "DCN Theatre", W1: 0, W2: 0, Now: 0 },
  ],
};

const isDemoMode = () => !CONFIG.SHEET_ID || CONFIG.SHEET_ID === "YOUR_GOOGLE_SHEET_ID_HERE";

/* ---------------------------------------------------------------
   Filter state + wiring
   --------------------------------------------------------------- */

const state = {
  mode: "mtd",       // "mtd" | "ftd" | "custom"
  rangeStart: null,  // Date, midnight
  rangeEnd: null,    // Date, midnight
  ctx: null,         // { mtdAcquirers, ftdAcquirers, metaList, events }
};

function currentMonthRange() {
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), 1);
  const end = toMidnight(now);
  return { start, end };
}

function fullHistoryRange(events) {
  const dated = events.filter((e) => e.date);
  if (!dated.length) return currentMonthRange();
  const start = dated.reduce((min, e) => (e.date < min ? e.date : min), dated[0].date);
  const end = toMidnight(new Date());
  return { start, end };
}

function isSameDay(a, b) {
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}

function detectMode(start, end) {
  const mtd = currentMonthRange();
  if (isSameDay(start, mtd.start) && isSameDay(end, mtd.end)) return "mtd";
  return "custom";
}

function timeframeMeta(mode, start, end) {
  if (mode === "mtd") {
    const now = new Date();
    return {
      label: "MTD (This Month)",
      dates: `${fmtDateLabel(start)} \u2013 ${fmtDateLabel(end)}`,
      narrativeSuffix: "this month",
    };
  }
  if (mode === "ftd") {
    return {
      label: "FTD (All Time)",
      dates: `Through ${fmtDateLabel(end)}`,
      narrativeSuffix: "to date",
    };
  }
  return {
    label: "Custom Range",
    dates: `${fmtDateLabel(start)} \u2013 ${fmtDateLabel(end)}`,
    narrativeSuffix: "in this range",
  };
}

function renderFromState() {
  if (isDemoMode() || !state.ctx) {
    renderSummary(DEMO_DATA.summary);
    renderTheatres(DEMO_DATA.theatres, DEMO_DATA.acquirers, DEMO_DATA.trend);
    return;
  }
  const { mode, rangeStart, rangeEnd, ctx } = state;
  const acquirers = buildAcquirersForRange(mode, rangeStart, rangeEnd, ctx);
  const trend = computeTrend(ctx.events, rangeStart, rangeEnd);
  const timeframe = timeframeMeta(mode, rangeStart, rangeEnd);
  const data = buildDashboardData(acquirers, trend, timeframe);
  renderSummary(data.summary);
  renderTheatres(data.theatres, data.acquirers, data.trend);
  syncControls();
}

function syncControls() {
  const startInput = document.getElementById("range-start");
  const endInput = document.getElementById("range-end");
  if (state.rangeStart) startInput.value = fmtDateInput(state.rangeStart);
  if (state.rangeEnd) endInput.value = fmtDateInput(state.rangeEnd);
  document.querySelectorAll(".preset-btn").forEach((btn) => {
    btn.classList.toggle("is-active", btn.dataset.preset === state.mode);
  });
}

function setupControls() {
  const startInput = document.getElementById("range-start");
  const endInput = document.getElementById("range-end");
  const applyBtn = document.getElementById("range-apply");
  const presetBtns = document.querySelectorAll(".preset-btn");

  if (isDemoMode()) {
    [startInput, endInput, applyBtn, ...presetBtns].forEach((el) => { el.disabled = true; });
    return;
  }

  presetBtns.forEach((btn) => {
    btn.addEventListener("click", () => {
      if (btn.dataset.preset === "mtd") {
        const r = currentMonthRange();
        state.mode = "mtd"; state.rangeStart = r.start; state.rangeEnd = r.end;
      } else if (btn.dataset.preset === "ftd") {
        const r = fullHistoryRange(state.ctx ? state.ctx.events : []);
        state.mode = "ftd"; state.rangeStart = r.start; state.rangeEnd = r.end;
      }
      renderFromState();
    });
  });

  applyBtn.addEventListener("click", () => {
    if (!startInput.value || !endInput.value) return;
    let start = toMidnight(new Date(startInput.value));
    let end = toMidnight(new Date(endInput.value));
    if (start > end) { const tmp = start; start = end; end = tmp; }
    state.mode = detectMode(start, end);
    state.rangeStart = start;
    state.rangeEnd = end;
    renderFromState();
  });
}

/* ---------------------------------------------------------------
   Boot
   --------------------------------------------------------------- */

function setStatus(mode, text) {
  const bar = document.getElementById("status-bar");
  bar.hidden = false;
  bar.className = `status-bar ${mode}`;
  document.getElementById("status-text").textContent = text;
}

function showError(err) {
  const grid = document.getElementById("dashboard-grid");
  const box = document.createElement("div");
  box.className = "load-error";
  box.innerHTML = `<strong>Couldn't load live data.</strong><br>${err.message}<br><br>
    Check that <code>CONFIG.SHEET_ID</code> in <code>script.js</code> is set, the
    spreadsheet is shared as "Anyone with the link", and the tab names match
    <code>${CONFIG.TABS.acquirerPerformance}</code> and <code>${CONFIG.TABS.visitLog}</code>.`;
  grid.prepend(box);
}

async function loadDashboard() {
  setupControls();

  if (isDemoMode()) {
    renderFromState();
    setStatus("", "Showing demo data \u2014 set CONFIG.SHEET_ID in script.js to connect your live Google Sheet");
    return;
  }

  try {
    setStatus("", "Connecting to live sheet\u2026");
    const [acquirerPerfRaw, visitLogRaw] = await Promise.all([
      fetchRawRows(CONFIG.TABS.acquirerPerformance),
      fetchRawRows(CONFIG.TABS.visitLog),
    ]);

    const mtdAcquirers = parsePerformanceSection(acquirerPerfRaw, "MTD Performance", MTD_FIELD_MAP);
    const ftdAcquirers = parsePerformanceSection(acquirerPerfRaw, "FTD Performance", FTD_FIELD_MAP);
    const events = parseVisitLogAll(visitLogRaw);
    const metaList = mtdAcquirers.map((a) => ({ name: a.name, theatre: a.theatre, target: a.target }));

    state.ctx = { mtdAcquirers, ftdAcquirers, metaList, events };
    if (!state.rangeStart) {
      const r = currentMonthRange();
      state.mode = "mtd"; state.rangeStart = r.start; state.rangeEnd = r.end;
    }

    renderFromState();
    setStatus("ok", `Live \u00b7 last synced ${new Date().toLocaleTimeString()}`);
  } catch (err) {
    console.error(err);
    setStatus("error", "Live sync failed \u2014 see details below");
    showError(err);
  }
}

document.addEventListener("DOMContentLoaded", () => {
  loadDashboard();
  if (CONFIG.REFRESH_MS > 0) {
    setInterval(loadDashboard, CONFIG.REFRESH_MS);
  }
});
