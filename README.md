# From Hours to Seconds: Power Automate Optimization Techniques That Actually Work

Session prep repo (24 Nov 2026, 50 min, Modern Workplace & Endpoint).

| Doc | Content |
|---|---|
| [docs/01-optimization-techniques.md](docs/01-optimization-techniques.md) | Catalog of techniques: the ones in the abstract plus extras, with snippets, limits and licensing notes |
| [docs/02-demo-scenarios.md](docs/02-demo-scenarios.md) | Slow vs fast flow pairs per use case, demo dataset, 50-minute agenda |
| [docs/03-code-app-benchmark-harness.md](docs/03-code-app-benchmark-harness.md) | Feasibility + design of the Code App that launches each variant and shows KPIs and run/edit links |

Status: PoC done ([poc/RESULTS.md](poc/RESULTS.md)) → 30 lab flows built and measured ([lab/RESULTS.md](lab/RESULTS.md)) → Code App published → session deck in [deck/](deck/README.md).

| Folder | Content |
|---|---|
| [poc/](poc/README.md) | Runnable kit for the PoC checklist: Dataverse setup scripts, PoC flow definitions, bulk API benchmark |
| [lab/](lab/README.md) | The slow-vs-fast flows (UC1–UC10) as code, data setup, deploy and CLI runner scripts, results |
| [app/](app/) | Code App (React + Vite + `@microsoft/power-apps`): Scenarios (run SLOW/MID/FAST, KPIs, speed-up, scoreboard), Seed & reset, PoC checks |
| [deck/](deck/README.md) | Session deck on the Experts Live template (the one shared with the co-speaker), updated with the measured results |

## Skills for GitHub Copilot and Claude Code

[`plugins/power-automate-best-practices`](plugins/power-automate-best-practices/README.md) packages everything learned here as
two agent skills plus a flow linter, to use alongside Microsoft's `power-automate` plugin
([microsoft/power-platform-skills](https://github.com/microsoft/power-platform-skills)):

```bash
/plugin marketplace add david-lorenzo88/power-automate-optimization-techniques
```

```bash
/plugin install power-automate-best-practices@power-automate-optimization
```

## License

[MIT](LICENSE)
