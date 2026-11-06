// LPH Daily Huddle Tool — ingest + analysis.
//
//   node huddle.js ingest  <repoDir> <file.xlsx ...>   save board snapshots / completed jobs
//   node huddle.js analyze <repoDir>                    rebuild data/days/*.json + data/index.json
//
// Reports (auto-detected by header row):
//   "Sold Work" (email subject "Todays Sold Jobs")  -> every job on TODAY's board (Plumbing/Sewer BUs)
//                                                       = the morning plan  -> data/board/<date>.json
//   "Daily Job Detail" (subject "Daily Job Report") -> completed jobs, trailing 7 days
//                                                       = what actually happened -> data/actual/<date>.json
//
// Counting rules (Matt, 11/6/2026):
//   - one job counts once no matter how many techs are on it
//   - Demand = DEM or EST tag; SPP = SYS CK tag; SLD WRK (or "Sold Work" job type) = sold work only
//   - multi-day sold work counts on day 1 only
//   - Sewer Service SPP is always 0
//   - conversion uses ServiceTitan's Converted flag; sold work and warranty/recall/no-charge
//     (no Opportunity) calls are kept out of the conversion rate
//   - revenue = "Jobs Total Revenue" (matches the ServiceTitan dashboard); sold work on the
//     board = "Jobs Subtotal" (pre-tax, matches what Matt enters)

const fs = require('fs');
const path = require('path');
let XLSX;
try { XLSX = require('xlsx'); } catch (e) { console.error('ERROR missing xlsx package: npm i xlsx@0.18.5'); process.exit(1); }

const CFG = JSON.parse(fs.readFileSync(path.join(__dirname, 'config.json'), 'utf8'));
const DEPTS = Object.keys(CFG.departments);

// ------------------------------------------------------------------ helpers
const str = (v) => (v == null ? '' : String(v)).trim();
const num = (v) => { if (v == null || v === '') return 0; if (typeof v === 'number') return isFinite(v) ? v : 0; const n = parseFloat(String(v).replace(/[$,\s]/g, '')); return isFinite(n) ? n : 0; };
const bool = (v) => v === true || /^true$/i.test(str(v));
const r2 = (n) => Math.round(n * 100) / 100;
function serialToISO(v) {
  if (v == null || v === '') return '';
  if (typeof v === 'number') { const d = new Date(Math.round((v - 25569) * 86400000)); return d.toISOString().slice(0, 10); }
  const s = str(v); let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/); if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/); if (m) { const y = m[3].length === 2 ? '20' + m[3] : m[3]; return `${y}-${m[1].padStart(2, '0')}-${m[2].padStart(2, '0')}`; }
  return '';
}
function timeStr(v) {
  if (typeof v === 'number') { const mins = Math.round((v % 1) * 1440); return `${String(Math.floor(mins / 60)).padStart(2, '0')}:${String(mins % 60).padStart(2, '0')}`; }
  return str(v);
}
function deptOf(bu) { for (const d of DEPTS) if (CFG.departments[d].businessUnits.some((b) => b.toLowerCase() === str(bu).toLowerCase())) return d; return null; }
const tagList = (t) => str(t).split(',').map((x) => x.trim().toUpperCase()).filter(Boolean);
const firstName = (t) => str(t).split(',')[0].trim().split(' ')[0] || 'Unassigned';
function shortJob(type) {
  const t = str(type);
  if (/water heater/i.test(t)) return /tankless/i.test(t) ? 'tankless water heater' : 'water heater';
  if (/^sold work/i.test(t)) return 'sold work';
  if (/system check/i.test(t)) return 'system check';
  return t.replace(/^(Estimate|Leak) - (Plumbing - )?/i, '').replace(/ - .*$/, '').toLowerCase();
}
const money = (n) => (n < 0 ? '-' : '') + '$' + Math.abs(Math.round(n)).toLocaleString('en-US');
const signed = (n) => (n >= 0 ? '+' : '-') + '$' + Math.abs(Math.round(n)).toLocaleString('en-US');
const pct = (x) => Math.round(x * 100) + '%';
const readJSON = (p, d) => { try { return JSON.parse(fs.readFileSync(p, 'utf8')); } catch (e) { return d; } };
const writeJSON = (p, o) => { fs.mkdirSync(path.dirname(p), { recursive: true }); fs.writeFileSync(p, JSON.stringify(o, null, 1)); };

function readSheet(file) {
  const wb = XLSX.readFile(file);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
  const hi = rows.findIndex((r) => r.some((c) => str(c) === 'Job #'));
  if (hi < 0) return { kind: null };
  const H = rows[hi].map(str);
  const recs = rows.slice(hi + 1).filter((r) => str(r[H.indexOf('Job #')]) && /^\d+$/.test(str(r[H.indexOf('Job #')])))
    .map((r) => Object.fromEntries(H.map((h, i) => [h, r[i]])));
  let kind = null;
  if (H.includes('Tags') && H.includes('Jobs Subtotal')) kind = 'board';
  else if (H.includes('Jobs Total Revenue') && H.includes('Converted')) kind = 'detail';
  return { kind, recs, H };
}
function dateFromName(file) {
  const m = path.basename(file).match(/Dated (\d{2})_(\d{2})_(\d{2}) - (\d{2})_(\d{2})_(\d{2})/);
  return m ? { from: `20${m[3]}-${m[1]}-${m[2]}`, to: `20${m[6]}-${m[4]}-${m[5]}` } : null;
}

// ------------------------------------------------------------------ classify a board row
function classifyBoard(j, dept) {
  const tags = tagList(j.Tags);
  const type = str(j['Job Type']);
  const status = str(j.Status);
  const techs = str(j['Assigned Technicians']);
  const appt = timeStr(j['Upcoming Appt Start Time']);
  const out = { cls: null, excluded: null };
  if (/^cancel/i.test(status)) out.excluded = 'Canceled';
  else if (/^hold$/i.test(status)) out.excluded = 'On hold';
  if (tags.includes('SLD WRK') || /^sold work/i.test(type)) out.cls = 'sold';
  else if (tags.includes('WARR') || tags.includes('CB') || tags.includes('RECALL') || /^non-repair/i.test(type)) out.cls = 'wcn';
  else if (!bool(j.Opportunity) && !tags.some((t) => t === 'DEM' || t === 'EST' || t === 'SYS CK')) out.cls = 'wcn';
  else if (tags.includes('SYS CK') && dept === 'Plumbing') out.cls = 'spp';
  else if (tags.includes('DEM') || tags.includes('EST')) out.cls = 'demand';
  else if (/system c/i.test(str(j['Business Unit'])) && dept === 'Plumbing') out.cls = 'spp';
  else out.cls = 'demand';
  if (dept === 'Sewer Service' && out.cls === 'spp') out.cls = 'demand';
  if (dept === 'Sewer Install' && out.cls !== 'sold') out.cls = 'lead';
  if (!out.excluded) {
    const placeholder = CFG.placeholderTechs.some((p) => techs.toLowerCase().includes(p.toLowerCase()));
    if (out.cls === 'sold' && appt === CFG.placeholderApptTime) out.excluded = '5 PM placeholder';
    else if (out.cls === 'lead') { /* sewer install leads count even before a rep is assigned */ }
    else if (!techs) out.excluded = 'No tech assigned';
    else if (placeholder) out.excluded = 'Parking Lot (unassigned)';
  }
  return out;
}

// ------------------------------------------------------------------ ingest
function ingest(repo, files) {
  for (const f of files) {
    const { kind, recs } = readSheet(f);
    const d = dateFromName(f);
    if (!kind) { console.log(`SKIP ${path.basename(f)} (not a Daily Huddle report)`); continue; }
    if (kind === 'board') {
      const date = d ? d.to : null;
      if (!date) { console.log(`ERROR ${path.basename(f)}: cannot read report date from file name`); continue; }
      const jobs = [];
      for (const j of recs) {
        const dept = deptOf(j['Business Unit']); if (!dept) continue;
        const c = classifyBoard(j, dept);
        jobs.push({
          job: str(j['Job #']), type: str(j['Job Type']), bu: str(j['Business Unit']), dept, status: str(j.Status),
          techs: str(j['Assigned Technicians']), tags: str(j.Tags), opp: bool(j.Opportunity),
          subtotal: r2(num(j['Jobs Subtotal'])), total: r2(num(j['Jobs Total'])), soldBy: str(j['Sold By']), soldOn: serialToISO(j['Sold On']),
          appt: timeStr(j['Upcoming Appt Start Time']), cls: c.cls, excluded: c.excluded,
        });
      }
      const p = path.join(repo, 'data/board', date + '.json');
      const prev = readJSON(p, null);
      // keep the FIRST (morning) snapshot as the plan; later copies are stored as "latest" for reference
      const base = path.basename(f).replace(/^\d+_/, '');
      if (prev && String(prev.capturedFrom).replace(/^\d+_/, '') === base) { console.log(`BOARD ${date} already saved`); continue; }
      if (prev && prev.jobs) {
        prev.latest = { capturedFrom: path.basename(f), jobs };
        writeJSON(p, prev);
        console.log(`BOARD ${date} kept morning snapshot, stored later copy (${jobs.length} jobs)`);
      } else {
        writeJSON(p, { date, capturedFrom: path.basename(f), savedAt: new Date().toISOString(), jobs });
        console.log(`BOARD ${date} ${jobs.length} jobs saved`);
      }
    } else {
      const byDate = {};
      for (const j of recs) {
        const dept = deptOf(j['Business Unit']); if (!dept) continue;
        if (str(j.Status) && !/^completed$/i.test(str(j.Status))) continue;
        const cd = serialToISO(j['Completion Date']); if (!cd) continue;
        (byDate[cd] = byDate[cd] || []).push({
          job: str(j['Job #']), type: str(j['Job Type']), bu: str(j['Business Unit']), dept,
          tech: str(j['Primary Technician']) || str(j['Assigned Technicians']).split(',')[0], techs: str(j['Assigned Technicians']),
          opp: bool(j.Opportunity), converted: bool(j.Converted), warranty: bool(j.Warranty), recall: bool(j.Recall),
          revenue: r2(num(j['Jobs Total Revenue'])), soldBy: str(j['Sold By']), estimates: num(j.Estimates),
          estSales: r2(num(j['Jobs Estimate Sales Subtotal'])), scheduled: serialToISO(j['Scheduled Date']), isLead: bool(j['Is Lead']),
        });
      }
      const src = path.basename(f);
      for (const [date, jobs] of Object.entries(byDate)) {
        const p = path.join(repo, 'data/actual', date + '.json');
        const prev = readJSON(p, null);
        // newer report wins (ServiceTitan edits made after the fact flow in); remember the first total too
        const firstTotal = prev ? (prev.firstTotal || prev.total) : null;
        const total = r2(jobs.reduce((s, j) => s + j.revenue, 0));
        if (prev && JSON.stringify(prev.jobs) === JSON.stringify(jobs)) continue;
        writeJSON(p, { date, source: src, savedAt: new Date().toISOString(), total, firstTotal: firstTotal == null ? total : firstTotal, jobs });
      }
      console.log(`DETAIL ${src}: ${Object.keys(byDate).length} day(s) saved (${Object.keys(byDate).sort().join(', ')})`);
    }
  }
}

// ------------------------------------------------------------------ analysis
function listDates(dir) { try { return fs.readdirSync(dir).filter((f) => /^\d{4}-\d{2}-\d{2}\.json$/.test(f)).map((f) => f.slice(0, 10)).sort(); } catch (e) { return []; } }
const isWeekday = (iso) => { const g = new Date(iso + 'T12:00:00Z').getUTCDay(); return g >= 1 && g <= 5; };

// sold work counts on day 1 only: skip if this job was already counted on an earlier board AND it's a continuation
function priorSoldDates(repo, date, job) {
  const out = [];
  for (const d of listDates(path.join(repo, 'data/board'))) {
    if (d >= date) break;
    const b = readJSON(path.join(repo, 'data/board', d + '.json'), null);
    const hit = b && b.jobs.find((x) => x.job === job && (/SLD WRK/i.test(x.tags) || /^sold work/i.test(x.type)) && !/cancel|hold/i.test(x.status));
    if (hit) out.push({ date: d, status: hit.status });
  }
  return out;
}

function plan(repo, date) {
  const b = readJSON(path.join(repo, 'data/board', date + '.json'), null);
  if (!b) return null;
  const actualDates = listDates(path.join(repo, 'data/actual'));
  const res = {};
  for (const dept of DEPTS) {
    // re-classify from the saved raw fields so rule changes apply to old snapshots too
    const jobs = b.jobs.filter((j) => j.dept === dept).map((j) => {
      const c = classifyBoard({ Tags: j.tags, 'Job Type': j.type, Status: j.status, 'Assigned Technicians': j.techs, 'Upcoming Appt Start Time': j.appt, Opportunity: j.opp, 'Business Unit': j.bu }, dept);
      return { ...j, cls: c.cls, excluded: c.excluded };
    });
    for (const j of jobs) {
      if (j.cls === 'sold' && !j.excluded) {
        const prior = priorSoldDates(repo, date, j.job);
        // continuation of a multi-day job: already counted earlier and work had started (In Progress now or before),
        // or it completed partially. A job that was simply rescheduled (never started) still counts.
        const started = /in progress/i.test(j.status) || prior.some((p) => /in progress|working/i.test(p.status));
        const doneEarlier = prior.length && actualDates.some((d) => d < date && d >= prior[0].date && (readJSON(path.join(repo, 'data/actual', d + '.json'), { jobs: [] }).jobs.some((x) => x.job === j.job)));
        if (prior.length && (started || doneEarlier)) j.excluded = `Multi-day — counted on ${prior[0].date.slice(5).replace('-', '/')}`;
      }
    }
    const live = jobs.filter((j) => !j.excluded);
    const cnt = (c) => live.filter((j) => j.cls === c).length;
    const sold = live.filter((j) => j.cls === 'sold');
    const soldTotal = r2(sold.reduce((s, j) => s + j.subtotal, 0));
    const rates = CFG.departments[dept].rates;
    const wcnJobs = live.filter((j) => j.cls === 'wcn');
    const isSppish = (j) => dept === 'Plumbing' && (/SYS CK/i.test(j.tags) || /system c/i.test(j.bu));
    const out = { jobs, demand: cnt('demand'), spp: cnt('spp'), wcn: wcnJobs.length, wcnDemand: wcnJobs.filter((j) => !isSppish(j)).length, wcnSpp: wcnJobs.filter(isSppish).length, soldJobs: sold.length, soldTotal, leads: cnt('lead') };
    if (rates) {
      out.demandRev = r2(out.demand * rates.demand.conv * rates.demand.avg);
      out.sppRev = r2(out.spp * rates.spp.conv * rates.spp.avg);
      out.projection = r2(out.demandRev + out.sppRev + soldTotal);
    } else out.projection = soldTotal;
    res[dept] = out;
  }
  return { date, capturedFrom: b.capturedFrom, depts: res };
}

function actualFor(repo, date) {
  const a = readJSON(path.join(repo, 'data/actual', date + '.json'), null);
  return a;
}

// Where did each planned job go? (completed on the day / later / never)
function findLater(repo, date, job) {
  for (const d of listDates(path.join(repo, 'data/actual'))) {
    if (d <= date) continue;
    const a = readJSON(path.join(repo, 'data/actual', d + '.json'), { jobs: [] });
    if (a.jobs.some((x) => x.job === job)) return d;
  }
  for (const d of listDates(path.join(repo, 'data/board'))) {
    if (d <= date) continue;
    const b = readJSON(path.join(repo, 'data/board', d + '.json'), { jobs: [] });
    if (b.jobs.some((x) => x.job === job && !/cancel/i.test(x.status))) return 'board:' + d;
  }
  return null;
}

function bridge(repo, date, P) {
  const A = actualFor(repo, date);
  if (!A) return null;
  const out = {};
  for (const dept of DEPTS) {
    const p = P ? P.depts[dept] : null;
    const rates = CFG.departments[dept].rates;
    const planJobs = p ? p.jobs : [];
    const planByJob = Object.fromEntries(planJobs.map((j) => [j.job, j]));
    const done = A.jobs.filter((j) => j.dept === dept);
    // classify actual jobs (use the morning plan's class when the job was on the board — it had tags)
    for (const j of done) {
      const pj = planByJob[j.job];
      if (/^sold work/i.test(j.type) || (pj && pj.cls === 'sold')) j.cls = 'sold';
      else if (!j.opp || j.warranty || j.recall) j.cls = 'wcn';
      else if (dept === 'Sewer Install') j.cls = 'lead';
      else if (pj && (pj.cls === 'demand' || pj.cls === 'spp')) j.cls = pj.cls;
      else if (dept === 'Plumbing' && /system c/i.test(j.bu)) j.cls = 'spp';
      else j.cls = 'demand';
      j.planned = !!(pj && !pj.excluded);
      j.plannedAs = pj ? (pj.excluded ? 'excluded: ' + pj.excluded : pj.cls) : null;
    }
    const actualTotal = r2(done.reduce((s, j) => s + j.revenue, 0));
    const res = { actual: actualTotal, projection: p ? p.projection : null, jobsCompleted: done.length, effects: [], cats: {}, lists: {} };
    if (!p) { out[dept] = res; continue; }
    res.variance = r2(actualTotal - p.projection);

    const doneIds = new Set(done.map((j) => j.job));
    const missed = planJobs.filter((j) => !j.excluded && !doneIds.has(j.job)).map((j) => ({ ...j, later: findLater(repo, date, j.job) }));
    const added = done.filter((j) => !j.planned);
    res.lists.missed = missed; res.lists.added = added;

    let explained = 0;
    const push = (e) => { e.amount = r2(e.amount); explained += e.amount; res.effects.push(e); };

    for (const k of rates ? ['demand', 'spp'] : []) {
      const rt = rates[k]; if (!rt || !rt.avg) continue;
      const runs = done.filter((j) => j.cls === k);
      const conv = runs.filter((j) => j.converted);
      const Np = p[k], Na = runs.length, Ca = conv.length;
      const Rconv = conv.reduce((s, j) => s + j.revenue, 0);
      const Rnon = runs.filter((j) => !j.converted).reduce((s, j) => s + j.revenue, 0);
      const big = conv.filter((j) => j.revenue >= rt.avg * CFG.bigTicketMultiple).sort((a, b) => b.revenue - a.revenue);
      res.cats[k] = { planned: Np, ran: Na, converted: Ca, convRate: Na ? Ca / Na : null, planConv: rt.conv, avgTicket: Ca ? Rconv / Ca : null, planAvg: rt.avg, revenue: r2(Rconv + Rnon), planRev: p[k + 'Rev'],
        plannedRan: runs.filter((j) => j.planned).length, addedRan: runs.filter((j) => !j.planned).length,
        missed: missed.filter((j) => j.cls === k).length,
        reclassed: done.filter((j) => j.planned && j.plannedAs === k && j.cls !== k).length };
      push({ key: k + '_volume', cat: k, label: `${k === 'spp' ? 'SPP' : 'Demand'} calls run vs plan`, amount: (Na - Np) * rt.conv * rt.avg, detail: `${Na} run vs ${Np} planned` });
      push({ key: k + '_conversion', cat: k, label: `${k === 'spp' ? 'SPP' : 'Demand'} conversion`, amount: (Ca - Na * rt.conv) * rt.avg, detail: `${Ca} of ${Na} converted (${Na ? pct(Ca / Na) : '—'} vs ${pct(rt.conv)} plan)` });
      push({ key: k + '_ticket', cat: k, label: `${k === 'spp' ? 'SPP' : 'Demand'} average ticket`, amount: Rconv - Ca * rt.avg, detail: `${Ca ? money(Rconv / Ca) : '—'} avg vs ${money(rt.avg)} plan`, jobs: big.map((j) => ({ job: j.job, tech: j.tech, type: j.type, revenue: j.revenue })) });
      if (Rnon) push({ key: k + '_fees', cat: k, label: `Fees on unconverted ${k === 'spp' ? 'SPP' : 'demand'} calls`, amount: Rnon, detail: `${runs.filter((j) => !j.converted && j.revenue).length} call(s) with a dispatch/diagnostic fee` });
    }
    if (dept === 'Sewer Install') {
      const leads = done.filter((j) => j.cls === 'lead');
      res.cats.lead = { planned: p.leads, ran: leads.length, revenue: r2(leads.reduce((s, j) => s + j.revenue, 0)) };
      if (res.cats.lead.revenue) push({ key: 'lead_rev', cat: 'lead', label: 'Revenue on sewer estimate visits', amount: res.cats.lead.revenue, detail: `${leads.length} estimate visit(s)` });
    }
    // sold work
    const soldDone = done.filter((j) => j.cls === 'sold');
    const plannedSold = planJobs.filter((j) => j.cls === 'sold' && !j.excluded);
    const pairs = plannedSold.map((pj) => ({ pj, aj: soldDone.find((a) => a.job === pj.job) }));
    const swDiffDone = pairs.filter((x) => x.aj).map((x) => ({ job: x.pj.job, tech: x.aj.tech, type: x.aj.type, planned: x.pj.subtotal, actual: x.aj.revenue, diff: r2(x.aj.revenue - x.pj.subtotal) }));
    const swMissed = pairs.filter((x) => !x.aj).map((x) => ({ job: x.pj.job, tech: firstName(x.pj.techs), type: x.pj.type, planned: x.pj.subtotal, later: findLater(repo, date, x.pj.job) }));
    const swAdded = soldDone.filter((a) => !plannedSold.some((pj) => pj.job === a.job)).map((a) => ({ job: a.job, tech: a.tech, type: a.type, actual: a.revenue, plannedAs: a.plannedAs }));
    res.cats.sold = { planned: p.soldTotal, plannedJobs: plannedSold.length, actual: r2(soldDone.reduce((s, j) => s + j.revenue, 0)), done: swDiffDone, missed: swMissed, added: swAdded };
    const d1 = swDiffDone.reduce((s, x) => s + x.diff, 0), d2 = -swMissed.reduce((s, x) => s + x.planned, 0), d3 = swAdded.reduce((s, x) => s + x.actual, 0);
    if (Math.abs(d1) >= 0.01) push({ key: 'sold_diff', cat: 'sold', label: 'Planned sold work billed differently than scheduled', amount: d1, detail: swDiffDone.filter((x) => Math.abs(x.diff) >= 1).map((x) => `${firstName(x.tech)} ${money(x.actual)} vs ${money(x.planned)}`).join('; ') });
    if (swMissed.length) push({ key: 'sold_missed', cat: 'sold', label: 'Planned sold work not completed', amount: d2, detail: swMissed.map((x) => `${firstName(x.tech)} ${shortJob(x.type)} ${money(x.planned)}`).join('; '), jobs: swMissed });
    if (swAdded.length) push({ key: 'sold_added', cat: 'sold', label: 'Sold work completed that was not on the morning board', amount: d3, detail: swAdded.map((x) => `${firstName(x.tech)} ${money(x.actual)}`).join('; '), jobs: swAdded });
    // warranty / recall / no-charge
    const w = done.filter((j) => j.cls === 'wcn');
    const wRev = w.reduce((s, j) => s + j.revenue, 0);
    res.cats.wcn = { planned: p.wcn, ran: w.length, revenue: r2(wRev) };
    if (wRev) push({ key: 'wcn', cat: 'wcn', label: 'Revenue on warranty / callback / no-opportunity calls', amount: wRev, detail: `${w.filter((j) => j.revenue).length} job(s)` });

    res.unexplained = r2(res.variance - explained);
    res.check = { jobsAccounted: done.length, sumOfJobs: actualTotal, explained: r2(explained), ok: Math.abs(res.unexplained) < 1 };
    out[dept] = res;
  }
  return { date, source: A.source, firstTotal: A.firstTotal, depts: out };
}

// ------------------------------------------------------------------ the "why" sentence
function whySentence(dept, r, P) {
  if (r.projection == null) return `No morning board snapshot was saved for this day, so there's no projection to compare against. Actual: ${money(r.actual)}.`;
  const v = r.variance;
  const parts = [];
  const head = Math.abs(v) < 50 ? `${dept} landed right on projection (${money(r.actual)} vs ${money(r.projection)}).`
    : `${dept} ${v > 0 ? 'beat' : 'missed'} projection by ${money(Math.abs(v))} (${money(r.actual)} vs ${money(r.projection)}).`;
  // group effects per call type so the sentence reads naturally
  const sig = (x) => Math.abs(x) >= Math.max(150, Math.abs(r.projection) * 0.03);
  for (const k of ['demand', 'spp']) {
    const c = r.cats[k]; if (!c) continue;
    const vol = r.effects.find((e) => e.key === k + '_volume'), cv = r.effects.find((e) => e.key === k + '_conversion'), tk = r.effects.find((e) => e.key === k + '_ticket');
    const name = k === 'spp' ? 'SPP' : 'demand';
    const bits = [];
    if (vol && (sig(vol.amount) || c.ran !== c.planned)) {
      let s = `ran ${c.ran} ${name} calls vs ${c.planned} planned`;
      const extra = [];
      if (c.missed) extra.push(`${c.missed} planned didn't run`);
      if (c.reclassed) extra.push(`${c.reclassed} closed as warranty/no-charge`);
      if (c.addedRan) extra.push(`${c.addedRan} added during the day`);
      if (extra.length) s += ` (${extra.join(', ')})`;
      if (sig(vol.amount)) s += ` ${signed(vol.amount)}`;
      if (c.ran !== c.planned || sig(vol.amount)) bits.push(s);
    }
    if (cv && c.ran && sig(cv.amount)) bits.push(`converted ${c.converted} of ${c.ran} (${pct(c.convRate)} vs ${pct(c.planConv)}) ${signed(cv.amount)}`);
    if (tk && c.converted && sig(tk.amount)) {
      let s = `average ${name} ticket ${money(c.avgTicket)} vs ${money(c.planAvg)} ${signed(tk.amount)}`;
      if (tk.amount > 0 && tk.jobs && tk.jobs.length) s += ` — ${tk.jobs.slice(0, 3).map((j) => `${firstName(j.tech)}'s ${money(j.revenue)} ${shortJob(j.type)}`).join(', ')}`;
      bits.push(s);
    }
    if (bits.length) parts.push(bits.join('; '));
  }
  const s = r.cats.sold;
  if (s) {
    const sw = [];
    const missed = r.effects.find((e) => e.key === 'sold_missed');
    const added = r.effects.find((e) => e.key === 'sold_added');
    const diff = r.effects.find((e) => e.key === 'sold_diff');
    if (missed && sig(missed.amount)) sw.push(`${money(-missed.amount)} of scheduled sold work didn't finish (${s.missed.map((x) => `${firstName(x.tech)}'s ${money(x.planned)} ${shortJob(x.type)}${x.later ? (x.later.startsWith('board:') ? ', now on ' + x.later.slice(11).replace('-', '/') : ', completed ' + x.later.slice(5).replace('-', '/')) : ''}`).join('; ')})`);
    if (added && sig(added.amount)) sw.push(`${money(added.amount)} of sold work was added after the morning board (${s.added.slice(0, 3).map((x) => `${firstName(x.tech)}'s ${money(x.actual)}`).join(', ')})`);
    if (diff && sig(diff.amount)) sw.push(`scheduled sold work billed ${signed(diff.amount)} vs the estimate`);
    if (sw.length) parts.push(sw.join('; '));
  }
  const w = r.effects.find((e) => e.key === 'wcn');
  if (w && sig(w.amount)) parts.push(`${money(w.amount)} came from warranty/callback/no-charge calls`);
  const f = r.effects.filter((e) => /_fees$/.test(e.key)).reduce((a, e) => a + e.amount, 0);
  if (sig(f)) parts.push(`${money(f)} in dispatch fees on unconverted calls`);
  const lead = r.cats.lead;
  if (dept === 'Sewer Install' && lead) parts.push(`${lead.ran} sewer estimate visit(s) ran vs ${lead.planned} scheduled`);
  if (!parts.length) return head;
  const cap = (x) => x.charAt(0).toUpperCase() + x.slice(1);
  return head + ' ' + parts.map((x) => cap(x) + '.').join(' ');
}


// ------------------------------------------------------------------ short one-line "why" (fits the Lovable box)
const kmoney = (n) => { const a = Math.abs(n); return '$' + (a >= 1000 ? (Math.round(a / 100) / 10).toString().replace(/\.0$/, '') + 'K' : Math.round(a).toString()); };
const kdelta = (n) => (n >= 0 ? '+' : '−') + kmoney(n);
const SHORTJ = (t) => shortJob(t).replace('tankless water heater', 'tankless').replace('water heater', 'WH').replace('system check', 'sys ck').replace(/sewer camera.*/, 'camera').replace(/ concern$/, '').replace(/drain stoppage/, 'drain');
function whyShort(dept, r) {
  // Matt's format (11/6): "Demand: Conv 43% vs 72% - Avg $1,309 vs $800 - Ran 7 vs 6 scheduled (1 went NC, 2 added). SPP: ..."
  if (r.projection == null) return 'No 8 AM board saved for this day.';
  const parts = [];
  for (const k of ['demand', 'spp']) {
    const c = r.cats[k]; if (!c || !c.planAvg) continue;
    if (!c.ran && !c.planned) continue;
    const bits = [];
    if (c.ran) bits.push(`Conv ${Math.round(c.convRate * 100)}% vs ${Math.round(c.planConv * 100)}%`);
    if (c.converted) bits.push(`Avg ${money(c.avgTicket)} vs ${money(c.planAvg)}`);
    const x = []; if (c.missed) x.push(`${c.missed} didn't run`); if (c.reclassed) x.push(`${c.reclassed} went NC`); if (c.addedRan) x.push(`${c.addedRan} added`);
    bits.push(`Ran ${c.ran} vs ${c.planned} scheduled${x.length ? ' (' + x.join(', ') + ')' : ''}`);
    parts.push(`${k === 'spp' ? 'SPP' : 'Demand'}: ${bits.join(' - ')}.`);
  }
  const sw = r.cats.sold;
  if (sw && (sw.planned || sw.actual) && Math.abs(sw.actual - sw.planned) >= 500) parts.push(`Sold work: ${money(sw.actual)} vs ${money(sw.planned)} scheduled.`);
  const ld = r.cats.lead;
  if (dept === 'Sewer Install' && ld && (ld.ran || ld.planned)) parts.push(`Leads: Ran ${ld.ran} vs ${ld.planned} scheduled.`);
  return parts.length ? parts.join(' ') : 'On plan.';
}

function prevWorkday(dates, date) { const ds = dates.filter((d) => d < date && isWeekday(d)); return ds.length ? ds[ds.length - 1] : null; }

function analyze(repo, since) {
  const boards = listDates(path.join(repo, 'data/board')).filter(isWeekday);
  const actuals = listDates(path.join(repo, 'data/actual')).filter(isWeekday);
  const all = [...new Set([...boards, ...actuals])].sort().filter((d) => !since || d >= since);
  // keep index entries for older days that weren't re-analyzed this run
  const oldIndex = readJSON(path.join(repo, 'data/index.json'), { days: [] });
  const index = (oldIndex.days || []).filter((x) => since && x.date < since).reverse();
  for (const date of all) {
    const P = boards.includes(date) ? plan(repo, date) : null;
    const B = actuals.includes(date) ? bridge(repo, date, P) : null;
    const day = { date, generatedAt: new Date().toISOString(), plan: P, results: B, why: {} };
    day.whyLong = {};
    if (B) for (const dept of DEPTS) { day.why[dept] = whyShort(dept, B.depts[dept]); day.whyLong[dept] = whySentence(dept, B.depts[dept], P); }
    writeJSON(path.join(repo, 'data/days', date + '.json'), day);
    index.push({ date, hasPlan: !!P, hasResults: !!B,
      projection: P ? Object.fromEntries(DEPTS.map((d) => [d, P.depts[d].projection])) : null,
      actual: B ? Object.fromEntries(DEPTS.map((d) => [d, B.depts[d].actual])) : null,
      checks: B ? Object.fromEntries(DEPTS.map((d) => [d, B.depts[d].check ? B.depts[d].check.ok : null])) : null });
  }
  // the morning view: today's commitment + yesterday's results
  const latest = boards[boards.length - 1] || null;
  const yest = latest ? prevWorkday(actuals.concat(boards), latest) : null;
  index.sort((a, b) => (a.date < b.date ? 1 : -1));
  writeJSON(path.join(repo, 'data/index.json'), { updatedAt: new Date().toISOString(), today: latest, yesterday: yest, days: index, config: CFG });
  console.log(`ANALYZED ${all.length} day(s); today=${latest} yesterday=${yest}`);
  for (const d of [yest, latest].filter(Boolean)) {
    const day = readJSON(path.join(repo, 'data/days', d + '.json'), {});
    for (const dept of DEPTS) {
      const r = day.results && day.results.depts[dept];
      if (r && r.check && !r.check.ok) console.log(`WARN ${d} ${dept}: unexplained ${r.unexplained}`);
    }
  }
}

// ------------------------------------------------------------------ cli
module.exports = { ingest, analyze };
if (require.main !== module) return;
const [cmd, repo, ...files] = process.argv.slice(2);
if (cmd === 'ingest') ingest(repo, files);
else if (cmd === 'analyze') analyze(repo, files[0] || null);
else { console.log('usage: node huddle.js ingest <repoDir> <xlsx...> | analyze <repoDir>'); process.exit(1); }
