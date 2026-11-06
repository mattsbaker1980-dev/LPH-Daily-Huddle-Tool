# LPH Daily Huddle Tool

Live page: https://mattsbaker1980-dev.github.io/LPH-Daily-Huddle-Tool/

Every weekday morning this tool reads two ServiceTitan reports that already arrive around 8:05 AM:

| Report (email subject) | What it gives |
|---|---|
| **Todays Sold Jobs** (report "Sold Work") | Every Plumbing / Sewer job on today's board → **Today's Commitment** (Demand calls, SPP calls, warranty/CB/NC, previously sold work, sewer leads) |
| **Daily Job Report** (report "Daily Job Detail") | Every completed job for the past 7 days → **Yesterday's Results** (actual revenue) |

It saves each morning's board as a snapshot, then the next morning compares that plan against what actually completed, job by job, and writes the "why did we win or lose" sentence.

## How the "why" is built

Projection = Demand calls × 72% × $800 + SPP calls × 40% × $350 + previously sold work (Plumbing); Sewer Service uses 60% × $242. The gap between projection and actual is split into:

- **Calls run vs plan** – calls that dropped off / got added
- **Conversion** – converted calls vs the planned rate (ServiceTitan's Converted flag; sold work and warranty/recall/no-charge calls are left out)
- **Average ticket** – actual average vs the planned average, naming the big tickets
- **Sold work** – scheduled sold work that didn't finish, billed differently, or got added
- **Warranty / callback / no-charge** and **dispatch fees** on unconverted calls

The pieces always add up exactly to the variance; the page shows a ✓ when every dollar is explained.

## Files

- `sync/huddle.js` – parsing, counting rules, analysis, sentence writer
- `sync/run.js` – the scheduled morning run (pull data → ingest → analyze → one commit)
- `sync/config.json` – planning rates and business units
- `data/board/` – morning board snapshots · `data/actual/` – completed jobs · `data/days/` – analysis per day
