"""Duplicates the Experts Live template frames into the 24-slide order build_deck.py expects.

    python deck/make_skeleton.py <template.pptx> <skeleton.pptx> <pptx-skill-dir>

Template frames: slide1 title · slide2 section divider (dark) · slide3 light, logo centred ·
slide4 dark content · slide5 light content · slide6 thank you.
<pptx-skill-dir> provides scripts/add_slide.py (copies a slide with all package bookkeeping).
"""
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

SEQ = ['slide1', 'slide5', 'slide4', 'slide3', 'slide5', 'slide5', 'slide2', 'slide5', 'slide5', 'slide3', 'slide2', 'slide5',
       'slide2', 'slide5', 'slide5', 'slide2', 'slide5', 'slide5', 'slide2', 'slide5', 'slide5', 'slide4', 'slide3', 'slide6']


def main(template, out, skill):
    work = Path(tempfile.mkdtemp()) / 'work'
    zipfile.ZipFile(template).extractall(work)
    used, files = set(), []
    for src in SEQ:
        if src not in used:
            used.add(src)
            files.append(f'{src}.xml')
            continue
        res = subprocess.run([sys.executable, f'{skill}/scripts/add_slide.py', f'{work}/', f'{src}.xml'], capture_output=True, text=True, check=True)
        files.append(re.search(r'(slide\d+\.xml)', res.stdout.split('Created')[-1]).group(1))
    rels = (work / 'ppt/_rels/presentation.xml.rels').read_text()
    rid = {t: r for r, t in re.findall(r'Id="(rId\d+)"[^>]*Target="slides/(slide\d+\.xml)"', rels)}
    rid.update({t: r for t, r in re.findall(r'Target="slides/(slide\d+\.xml)"[^>]*Id="(rId\d+)"', rels)})
    pres = (work / 'ppt/presentation.xml').read_text()
    ids = {r: i for i, r in re.findall(r'<p:sldId id="(\d+)" r:id="(rId\d+)"/>', pres)}
    lst = ''.join(f'<p:sldId id="{ids[rid[f]]}" r:id="{rid[f]}"/>' for f in files)
    (work / 'ppt/presentation.xml').write_text(re.sub(r'<p:sldIdLst>.*?</p:sldIdLst>', f'<p:sldIdLst>{lst}</p:sldIdLst>', pres, flags=re.S))
    Path(out).unlink(missing_ok=True)
    shutil.make_archive(str(Path(out).with_suffix('')), 'zip', work)
    Path(out).with_suffix('.zip').rename(out)
    print(f'{out}: {len(files)} slides')


if __name__ == '__main__':
    main(*sys.argv[1:4])
