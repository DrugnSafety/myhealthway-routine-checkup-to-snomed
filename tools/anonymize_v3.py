"""Same as v5 but matches the old '{이름}_{birth}.xlsx' (no v5 suffix)."""
import os, re, csv, shutil, sys, hashlib
from pathlib import Path

CHOSEONG = ['ㄱ','ㄲ','ㄴ','ㄷ','ㄸ','ㄹ','ㅁ','ㅂ','ㅃ','ㅅ','ㅆ','ㅇ','ㅈ','ㅉ','ㅊ','ㅋ','ㅌ','ㅍ','ㅎ']

def to_choseong(name):
    out = []
    for c in name:
        code = ord(c)
        if 0xAC00 <= code <= 0xD7A3:
            idx = (code - 0xAC00) // (21 * 28)
            out.append(CHOSEONG[idx])
        else:
            out.append(c)
    return ''.join(out)

def short_hash(name, birth):
    return hashlib.sha256(f'{name}|{birth}'.encode('utf-8')).hexdigest()[:4]

src = Path(sys.argv[1]); dst = Path(sys.argv[2])
mapping_csv = Path(sys.argv[3])
dst.mkdir(parents=True, exist_ok=True)
rows = []
pattern = re.compile(r'^(.+?)_(\d{4}-\d{2}-\d{2})(?:_(.+))?\.xlsx$')
seen = set()
for fp in sorted(src.glob('*.xlsx')):
    if fp.name.startswith('_'):
        shutil.copy2(fp, dst / fp.name); continue
    m = pattern.match(fp.name)
    if not m:
        print(f'skip: {fp.name}'); continue
    name, birth, extra = m.group(1), m.group(2), m.group(3)
    cho = to_choseong(name)
    key = (cho, birth)
    if key in seen:
        new = f'routineCheckup_{cho}_{birth}_{short_hash(name,birth)}.xlsx'
    else:
        seen.add(key)
        new = f'routineCheckup_{cho}_{birth}.xlsx'
    if extra: new = new.replace('.xlsx', f'_{extra}.xlsx')
    shutil.copy2(fp, dst / new)
    rows.append({'original_filename': fp.name, 'anonymized_filename': new,
                 'patient_name': name, 'patient_choseong': cho, 'birth_date': birth})

print(f'{len(rows)} files anonymized → {dst}')
mapping_csv.parent.mkdir(parents=True, exist_ok=True)
with open(mapping_csv, 'w', encoding='utf-8', newline='') as f:
    w = csv.DictWriter(f, fieldnames=['original_filename','anonymized_filename','patient_name','patient_choseong','birth_date'])
    w.writeheader(); w.writerows(rows)
print(f'mapping → {mapping_csv}')
