# Session deck

`From-Hours-to-Seconds.pptx`: 30 slides on the Experts Live Emirates template, with speaker notes and demo cues on every slide.

Rebuild (needs the template .pptx, python-pptx, and the pptx skill scripts):

1. `node build/render.js`: renders the icon PNGs (react-icons + sharp)
2. `python3 build/structure.py <skill-dir> template.pptx unpacked/`: duplicates the template slides into the 30-slide order, then zips `unpacked/` → `base.pptx`
3. `python3 build/build.py base.pptx deck.pptx <icons-dir>`: fills in all the content
