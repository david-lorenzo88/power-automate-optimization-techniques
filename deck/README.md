# Session deck

`From-Hours-to-Seconds.pptx`: 32 slides on the Experts Live Emirates template, with speaker notes and demo cues on every slide.
The numbers on the slides are the measured lab results from [`lab/RESULTS.md`](../lab/RESULTS.md) and [`poc/RESULTS.md`](../poc/RESULTS.md).

| # | Slide | # | Slide |
|---|---|---|---|
| 1 | Title | 17 | 03 · Aggregate section |
| 2 | Hook: the five-hour flow | 18 | Group by and gotchas |
| 3 | Measured, not promised (177× / 130× / 124×) | 19 | 04 · Bulk section |
| 4 | Why slow flows cost money | 20 | Bulk options (SharePoint ≤ 500 per batch, not atomic) |
| 5 | Six numbers to know by heart | 21 | $batch anatomy + UC4 measured |
| 6 | Agenda | 22 | Changesets: CreateMultiple 0/100 vs loop 19/20 |
| 7 | The Flow Optimization Lab | 23 | 05 · Filter section |
| 8 | 01 · Diagnose section | 24 | UC6 funnel + measured payloads |
| 9 | Optimize the action that owns the run | 25 | UC7 trigger conditions (100 → 10 runs) |
| 10 | 02 · Kill loops section | 26 | 06 · Tune section |
| 11 | UC1: 130× faster, same result | 27 | Six settings |
| 12 | Expressions that replace loops | 28 | UC8: when concurrency bites |
| 13 | UC2 joins (3.0 min → 1.5 s) | 29 | UC10 fan-out + the 120-second rule |
| 14 | To loop or not to loop | 30 | Scoreboard: every pair, measured |
| 15 | 03 · XPath section | 31 | Takeaways + take it home |
| 16 | UC3 XPath (177×) | 32 | Thank you |

Rebuild (needs the template .pptx, python-pptx, and the pptx skill scripts):

1. `node build/render.js`: renders the icon PNGs (react-icons + sharp)
2. `python3 build/structure.py <skill-dir> template.pptx unpacked/`: duplicates the template slides into the 32-slide order, then zips `unpacked/` → `base.pptx`
3. `python3 build/build.py base.pptx deck.pptx <icons-dir>`: fills in all the content
