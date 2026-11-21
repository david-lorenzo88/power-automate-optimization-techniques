import sys, re, subprocess, zipfile, shutil, os
SK, tpl, out_dir = sys.argv[1], sys.argv[2], sys.argv[3]
# T=title(1) S=divider(2) L3=spotlight light(3) D=dark(4) L5=light(5) E=end(6)
src = {'T':'slide1.xml','S':'slide2.xml','L3':'slide3.xml','D':'slide4.xml','L5':'slide5.xml','E':'slide6.xml'}
order = ['T','D','L3', 'L5','D','L3','D', 'S','L5', 'S','L3','L5','D','L3', 'S','D','L5', 'S','L3','D','L5', 'S','L3','D', 'S','L5','D','L5','L3','L5','D','E']
shutil.rmtree(out_dir, ignore_errors=True)
zipfile.ZipFile(tpl).extractall(out_dir)
new = []
for k in order:
    r = subprocess.run([sys.executable, f'{SK}/scripts/add_slide.py', out_dir, src[k]], capture_output=True, text=True, check=True)
    m = re.search(r'(slide\d+\.xml)', r.stdout.split('Created')[1])
    new.append(m.group(1))
pres = open(f'{out_dir}/ppt/presentation.xml').read()
rels = open(f'{out_dir}/ppt/_rels/presentation.xml.rels').read()
rid = {t.split('/')[-1]: i for i, t in re.findall(r'Id="(rId\d+)"[^>]*Target="(slides/[^"]+)"', rels)}
if not rid:
    rid = {t.split('/')[-1]: i for t, i in re.findall(r'Target="(slides/[^"]+)"[^>]*Id="(rId\d+)"', rels)}
ids = dict((r_, i) for i, r_ in re.findall(r'<p:sldId id="(\d+)" r:id="(rId\d+)"/>', pres))
lst = ''.join(f'<p:sldId id="{ids[rid[s]]}" r:id="{rid[s]}"/>' for s in new)
pres = re.sub(r'<p:sldIdLst>.*?</p:sldIdLst>', f'<p:sldIdLst>{lst}</p:sldIdLst>', pres, flags=re.S)
open(f'{out_dir}/ppt/presentation.xml','w').write(pres)
subprocess.run([sys.executable, f'{SK}/scripts/clean.py', out_dir], check=True, capture_output=True)
print(len(new), 'slides:', new[:5], '...')
