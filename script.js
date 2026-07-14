const CONFIG = {
  API_ENDPOINT: "/api/fono",

  TABS: {
    acquirerPerformance: "Acquirer Performance",
    visitLog: "Visit Log",
  },

  // Re-fetch the sheet every N milliseconds so the dashboard stays
  // live without a page refresh. Set to 0 to disable auto-refresh.
  REFRESH_MS: 60000,

  // Set true to force demo data regardless of API availability
  // (useful for previewing layout/styling changes offline).
  FORCE_DEMO: false,
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

// This reads the sheet's "FTD Performance" section — confirmed to mean
// "For The Day" (today only), so we treat/label it as the Today block.
const TODAY_FIELD_MAP = {
  visitNests: "Visit Nests",
  pipelineNest: "Funnel Nest",
  contractingNest: "Contracting Nest",
  contractedNest: "Contracted Nest",
  contractedCount: "Contracted Count",
};

/* ---------------------------------------------------------------
   Fetch both tabs from the authenticated /api/fono endpoint. Returns
   { acquirerPerformance: [][], visitLog: [][] } — already plain 2D
   arrays, same shape the parsers below expect.
   --------------------------------------------------------------- */

async function fetchFonoData() {
  const res = await fetch(CONFIG.API_ENDPOINT);
  if (!res.ok) {
    let detail = "";
    try { detail = (await res.json()).error || ""; } catch (_) { /* non-JSON error body */ }
    throw new Error(`Request to ${CONFIG.API_ENDPOINT} failed (HTTP ${res.status})${detail ? `: ${detail}` : ""}.`);
  }
  const data = await res.json();
  if (!data || !Array.isArray(data.acquirerPerformance) || !Array.isArray(data.visitLog)) {
    throw new Error(`Unexpected response from ${CONFIG.API_ENDPOINT} — expected { acquirerPerformance, visitLog } arrays.`);
  }
  return data;
}

// The API returns dates as Sheets/Excel serial numbers (days since
// 1899-12-30) rather than formatted strings, so parsing is exact and
// timezone/locale-independent.
function serialToDate(serial) {
  if (typeof serial !== "number" || !Number.isFinite(serial)) return null;
  const excelEpochUTC = Date.UTC(1899, 11, 30);
  const d = new Date(excelEpochUTC + Math.round(serial) * 86400000);
  return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
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
    if (!row || row[dateIdx] === undefined || row[dateIdx] === "") continue;
    const eventDate = typeof row[dateIdx] === "number" ? serialToDate(row[dateIdx]) : null;
    if (!eventDate) continue;
    events.push({
      date: eventDate,
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
  if (mode === "today") return ctx.todayAcquirers;

  // "all" (All Time) and "custom" both compute live from Visit Log —
  // "all" just uses the full-history range (earliest event -> today).
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
    { Key: "ShortLabel", Value: timeframe.shortLabel },
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

  const shortLabel = kv.ShortLabel || "MTD";
  document.getElementById("summary-heading").textContent = `${shortLabel} Summary`;
  document.getElementById("stat-total-label").textContent = `Total ${shortLabel} Visit Nests`;
  document.getElementById("stat-contracted-label").textContent = `${shortLabel} Contracted Nests`;
  document.getElementById("target-eyebrow").textContent = `${shortLabel} Pipeline vs. Target`;

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

function renderTheatres(theatreRows, acquirerRows, trendRows, shortLabel) {
  const container = document.getElementById("theatre-columns");
  container.innerHTML = "";
  const template = document.getElementById("theatre-template");

  theatreRows.forEach((t) => {
    const node = template.content.cloneNode(true);

    const visits = t.Visits, pipeline = t.Pipeline, contractG = t.ContractG, contracted = t.Contracted;

    node.querySelector(".theatre-name").textContent = t.Theatre;
    node.querySelector(".theatre-visits").textContent = `${visits.toLocaleString()} Visit Nests`;
    node.querySelector(".acquirer-heading").textContent = `Acquirer ${shortLabel || "MTD"}`;

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
   Demo data — only shown if CONFIG.FORCE_DEMO is set to true.
   The filter controls are disabled in demo mode.
   --------------------------------------------------------------- */

const DEMO_DATA = {
  summary: [
    { Key: "TimeframeLabel", Value: "MTD (This Month)" },
    { Key: "ShortLabel", Value: "MTD" },
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

const isDemoMode = () => CONFIG.FORCE_DEMO === true;

/* ---------------------------------------------------------------
   Filter state + wiring
   --------------------------------------------------------------- */

const state = {
  mode: "mtd",       // "mtd" | "today" | "all" | "custom"
  rangeStart: null,  // Date, midnight
  rangeEnd: null,    // Date, midnight
  ctx: null,         // { mtdAcquirers, todayAcquirers, metaList, events }
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
    return {
      label: "MTD (This Month)",
      shortLabel: "MTD",
      dates: `${fmtDateLabel(start)} \u2013 ${fmtDateLabel(end)}`,
      narrativeSuffix: "this month",
    };
  }
  if (mode === "today") {
    return {
      label: "Today",
      shortLabel: "Today",
      dates: fmtDateLabel(end),
      narrativeSuffix: "today",
    };
  }
  if (mode === "all") {
    return {
      label: "All Time",
      shortLabel: "All Time",
      dates: `${fmtDateLabel(start)} \u2013 ${fmtDateLabel(end)}`,
      narrativeSuffix: "all-time",
    };
  }
  return {
    label: "Custom Range",
    shortLabel: "Range",
    dates: `${fmtDateLabel(start)} \u2013 ${fmtDateLabel(end)}`,
    narrativeSuffix: "in this range",
  };
}

function renderFromState() {
  if (isDemoMode() || !state.ctx) {
    renderSummary(DEMO_DATA.summary);
    renderTheatres(DEMO_DATA.theatres, DEMO_DATA.acquirers, DEMO_DATA.trend, "MTD");
    return;
  }
  const { mode, rangeStart, rangeEnd, ctx } = state;
  const acquirers = buildAcquirersForRange(mode, rangeStart, rangeEnd, ctx);
  const trend = computeTrend(ctx.events, rangeStart, rangeEnd);
  const timeframe = timeframeMeta(mode, rangeStart, rangeEnd);
  const data = buildDashboardData(acquirers, trend, timeframe);
  renderSummary(data.summary);
  renderTheatres(data.theatres, data.acquirers, data.trend, timeframe.shortLabel);
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
      } else if (btn.dataset.preset === "today") {
        const today = toMidnight(new Date());
        state.mode = "today"; state.rangeStart = today; state.rangeEnd = today;
      } else if (btn.dataset.preset === "all") {
        const r = fullHistoryRange(state.ctx ? state.ctx.events : []);
        state.mode = "all"; state.rangeStart = r.start; state.rangeEnd = r.end;
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
    Check that <code>${CONFIG.API_ENDPOINT}</code> is deployed, the
    <code>GOOGLE_SERVICE_ACCOUNT</code> environment variable is set, the
    sheet is shared with the service account's email as Viewer, and the tab
    names match <code>${CONFIG.TABS.acquirerPerformance}</code> and
    <code>${CONFIG.TABS.visitLog}</code>.`;
  grid.prepend(box);
}

async function loadDashboard() {
  setupControls();

  if (isDemoMode()) {
    renderFromState();
    setStatus("", "Showing demo data \u2014 set CONFIG.FORCE_DEMO to false in script.js to use the live sheet");
    return;
  }

  try {
    setStatus("", "Connecting to live sheet\u2026");
    const { acquirerPerformance: acquirerPerfRaw, visitLog: visitLogRaw } = await fetchFonoData();

    const mtdAcquirers = parsePerformanceSection(acquirerPerfRaw, "MTD Performance", MTD_FIELD_MAP);
    const todayAcquirers = parsePerformanceSection(acquirerPerfRaw, "FTD Performance", TODAY_FIELD_MAP);
    const events = parseVisitLogAll(visitLogRaw);
    const metaList = mtdAcquirers.map((a) => ({ name: a.name, theatre: a.theatre, target: a.target }));

    state.ctx = { mtdAcquirers, todayAcquirers, metaList, events };
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
