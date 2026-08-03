import glob, os, json
from PIL import Image

MAX_H = 1330          # px, at 1400 px wide -> ~6.2 in tall on a 6.5 in column
OVERLAP = 45

out = {}
for f in sorted(glob.glob('shots_annotated/*/*.png')):
    if '__s' in f:
        continue
    im = Image.open(f)
    w, h = im.size
    base = f[:-4]
    if h <= MAX_H:
        out[os.path.basename(base)] = [f]
        continue
    parts, y, i = [], 0, 0
    while y < h:
        bottom = min(h, y + MAX_H)
        p = f'{base}__s{i}.png'
        im.crop((0, y, w, bottom)).save(p)
        parts.append(p)
        if bottom >= h:
            break
        y = bottom - OVERLAP
        i += 1
    out[os.path.basename(base)] = parts

json.dump(out, open('data/slices.json', 'w'), indent=1)
print('screens:', len(out), '| multi-part:', sum(1 for v in out.values() if len(v) > 1))
for k, v in out.items():
    if len(v) > 1:
        print('  ', k, len(v), 'parts')
