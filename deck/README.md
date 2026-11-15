# Session deck

`From-Hours-to-Seconds-ExpertsLive.pptx`: 24 slides on the Experts Live Emirates template, with speaker notes.
Every number on a slide is a measured lab result (see `lab/README.md`, `poc/RESULTS.md`).

Rebuild after new measurements (edit the `R` dictionary at the top of `build_deck.py`):

```bash
python deck/make_skeleton.py <Experts Live template.pptx> /tmp/skeleton.pptx <pptx-skill-dir>
python deck/build_deck.py /tmp/skeleton.pptx deck/From-Hours-to-Seconds-ExpertsLive.pptx
```

| # | Slide | # | Slide |
|---|---|---|---|
| 1 | Title | 13 | 03 Go bulk |
| 2 | Hook: up to 177× faster | 14 | UC4 SharePoint $batch + limits |
| 3 | Why slow flows cost money | 15 | UC5 CreateMultiple + rollback |
| 4 | Agenda | 16 | 04 Filter at the source |
| 5 | Find the slow action | 17 | UC6 $filter/$select/$expand |
| 6 | How we measure (the lab) | 18 | UC7 trigger conditions |
| 7 | 01 Kill the loops | 19 | 05 Tune the platform |
| 8 | UC1 Filter array + Select | 20 | UC8 pagination + race traps |
| 9 | UC2 dictionary join | 21 | UC10 fan-out + respond early |
| 10 | Do you need a loop at all? | 22 | Scoreboard (chart) |
| 11 | 02 Aggregate without iterating | 23 | Take home |
| 12 | UC3 xpath | 24 | Thank you |

All numbers come from `lab/RESULTS.md` (UC9 Excel and UC6-DV appear on the scoreboard and the UC6 slide).
