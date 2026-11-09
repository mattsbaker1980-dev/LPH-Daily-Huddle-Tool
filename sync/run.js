// LPH Daily Huddle Tool — one-shot morning run (used by the scheduled task).
//
//   GH_TOKEN=... node run.js <rawGmailMessage.json ...>
//
// 1. downloads the last ~45 days of saved data from GitHub (so multi-day / "moved to" checks work)
// 2. extracts the .xlsx attachments from the Gmail RAW messages
// 3. saves the board snapshot / completed jobs, re-analyzes recent days
// 4. publishes everything that changed as ONE commit
// Prints "DONE ..." on success; any line starting with ERROR is a failure.

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const OWNER = 'mattsbaker1980-dev', REPO = 'LPH-Daily-Huddle-Tool', BRANCH = 'main';
const API = `https://api.github.com/repos/${OWNER}/${REPO}`;
const TOKEN = process.env.GH_TOKEN;
if (!TOKEN) { console.log('ERROR GH_TOKEN not set'); process.exit(1); }
const H = { Authorization: `Bearer ${TOKEN}`, Accept: 'application/vnd.github+json', 'Content-Type': 'application/json', 'User-Agent': 'lph-daily-huddle' };
const WORK = path.resolve(process.env.HUDDLE_WORK || '/root/work/huddle');
const LOOKBACK_DAYS = 45;

async function gh(method, url, body) {
  for (let i = 0; i < 3; i++) {
    const r = await fetch(url.startsWith('http') ? url : API + url, { method, headers: H, body: body ? JSON.stringify(body) : undefined });
    if (r.ok) return r.status === 204 ? null : r.json();
    if (r.status >= 500 || r.status === 429) { await new Promise((s) => setTimeout(s, 2000 * (i + 1))); continue; }
    throw new Error(`${method} ${url} -> ${r.status} ${(await r.text()).slice(0, 200)}`);
  }
  throw new Error(`${method} ${url} failed after retries`);
}
const isoDaysAgo = (n) => new Date(Date.now() - n * 86400000).toISOString().slice(0, 10);
const blobSha = (buf) => require('crypto').createHash('sha1').update(Buffer.concat([Buffer.from(`blob ${buf.length}\0`), buf])).digest('hex');

(async () => {
  try {
    fs.rmSync(WORK, { recursive: true, force: true });
    fs.mkdirSync(path.join(WORK, 'xlsx'), { recursive: true });
    const since = isoDaysAgo(LOOKBACK_DAYS);

    // 1. pull recent data
    const ref = await gh('GET', `/git/ref/heads/${BRANCH}`);
    const baseCommit = ref.object.sha;
    const commit = await gh('GET', `/git/commits/${baseCommit}`);
    const tree = await gh('GET', `/git/trees/${commit.tree.sha}?recursive=1`);
    const remote = {};
    for (const t of tree.tree) if (t.type === 'blob') remote[t.path] = t.sha;
    const want = Object.keys(remote).filter((p) => p === 'data/index.json' || (/^data\/(board|actual|days)\/(\d{4}-\d{2}-\d{2})\.json$/.test(p) && p.match(/(\d{4}-\d{2}-\d{2})/)[1] >= since));
    await Promise.all(want.map(async (p) => {
      const b = await gh('GET', `/git/blobs/${remote[p]}`);
      const dest = path.join(WORK, p); fs.mkdirSync(path.dirname(dest), { recursive: true });
      fs.writeFileSync(dest, Buffer.from(b.content, 'base64'));
    }));
    console.log(`PULLED ${want.length} data file(s) since ${since}`);
    // job notes from the Job Review site (public repo) for waived-fee checks
    const months = [...new Set([isoDaysAgo(0).slice(0, 7), isoDaysAgo(LOOKBACK_DAYS).slice(0, 7), isoDaysAgo(20).slice(0, 7)])];
    fs.mkdirSync(path.join(WORK, 'data/jr'), { recursive: true });
    for (const m of months) {
      try {
        const r = await fetch(`https://raw.githubusercontent.com/mattsbaker1980-dev/LPH-Materials-Invoice/main/data/review/${m}.json?cb=${Date.now()}`);
        if (r.ok) { fs.writeFileSync(path.join(WORK, 'data/jr', m + '.json'), Buffer.from(await r.arrayBuffer())); console.log(`NOTES ${m} loaded from Job Review`); }
        else console.log(`NOTE Job Review notes for ${m} not available (${r.status})`);
      } catch (e) { console.log(`NOTE Job Review notes for ${m} not loaded: ${e.message}`); }
    }

    // 2. extract attachments
    // normalize the Gmail tool's saved files: plain {raw}, or [{type:'text', text:'{...raw...}'}]
    const raws = process.argv.slice(2).map((f, i) => {
      try {
        let j = JSON.parse(fs.readFileSync(f, 'utf8'));
        if (Array.isArray(j) && j[0] && typeof j[0].text === 'string') j = JSON.parse(j[0].text);
        if (j && j.raw) { const out = path.join(WORK, `raw_${i}.json`); fs.writeFileSync(out, JSON.stringify({ id: j.id, raw: j.raw })); return out; }
      } catch (e) { /* leave as-is */ }
      return f;
    });
    if (raws.length) {
      const out = execFileSync('node', [path.join(__dirname, 'gmail_extract.js'), path.join(WORK, 'xlsx'), ...raws], { encoding: 'utf8' });
      process.stdout.write(out);
    } else console.log('NOTE no Gmail messages given — re-analyzing saved data only');
    const files = fs.readdirSync(path.join(WORK, 'xlsx')).filter((f) => /\.xlsx$/i.test(f))
      .sort((a, b) => (/Detail/i.test(a) ? 0 : 1) - (/Detail/i.test(b) ? 0 : 1)) // completed jobs first, then boards
      .map((f) => path.join(WORK, 'xlsx', f));

    // 3. ingest + analyze
    const huddle = require('./huddle.js');
    huddle.ingest(WORK, files);
    huddle.analyze(WORK, isoDaysAgo(LOOKBACK_DAYS - 10));

    // 4. publish changed files as one commit
    const changed = [];
    const walk = (dir) => { for (const f of fs.readdirSync(dir)) { const p = path.join(dir, f); if (fs.statSync(p).isDirectory()) walk(p); else changed.push(p); } };
    walk(path.join(WORK, 'data'));
    const entries = [];
    for (const abs of changed) {
      const rel = path.relative(WORK, abs).split(path.sep).join('/');
      if (rel.startsWith('data/jr/')) continue;
      const buf = fs.readFileSync(abs);
      if (remote[rel] === blobSha(buf)) continue;
      const blob = await gh('POST', '/git/blobs', { content: buf.toString('base64'), encoding: 'base64' });
      entries.push({ path: rel, mode: '100644', type: 'blob', sha: blob.sha });
    }
    if (!entries.length) { console.log('DONE nothing changed'); return; }
    const newTree = await gh('POST', '/git/trees', { base_tree: commit.tree.sha, tree: entries });
    const idx = JSON.parse(fs.readFileSync(path.join(WORK, 'data/index.json'), 'utf8'));
    const newCommit = await gh('POST', '/git/commits', { message: `Daily Huddle data ${idx.today || ''} (${entries.length} file(s))`, tree: newTree.sha, parents: [baseCommit] });
    await gh('PATCH', `/git/refs/heads/${BRANCH}`, { sha: newCommit.sha });
    console.log(`DONE published ${entries.length} file(s); today=${idx.today} yesterday=${idx.yesterday}`);
  } catch (e) {
    console.log('ERROR ' + (e && e.message ? e.message : e));
    process.exit(1);
  }
})();
