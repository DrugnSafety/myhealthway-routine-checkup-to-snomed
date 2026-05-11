"""Rename per-subject files: replace Korean name with 초성 (consonant initials).
Pattern: '{원본이름}_{birth}_v5.xlsx' → '{초성}_{birth}_v5.xlsx'
Source files are filenames like '신정순_1949-05-30_v5.xlsx'.

Preserves a private mapping CSV that should NEVER be uploaded to GitHub.
"""
import os, re, csv, shutil, sys, hashlib
from pathlib import Path

CHOSEONG = ['ㄱ','ㄲ','ㄴ','ㄷ','ㄸ','ㄹ','ㅁ','ㅂ','ㅃ','ㅅ',
            'ㅆ','ㅇ','ㅈ','ㅉ','ㅊ','ㅋ','ㅌ','ㅍ','ㅎ']

def to_choseong(name: str) -> str:
    out = []
    for c in name:
        code = ord(c)
        if 0xAC00 <= code <= 0xD7A3:
            idx = (code - 0xAC00) // (21 * 28)
            out.append(CHOSEONG[idx])
        else:
            out.append(c)
    return ''.join(out)

def short_hash(name: str, birth: str) -> str:
    """Stable 4-char hex hash for deduplication of same-초성 patients."""
    return hashlib.sha256(f'{name}|{birth}'.encode('utf-8')).hexdigest()[:4]

def process(src_dir: Path, dst_dir: Path, prefix: str = 'routineCheckup'):
    dst_dir.mkdir(parents=True, exist_ok=True)
    mapping = []
    pattern = re.compile(r'^(.+)_(\d{4}-\d{2}-\d{2})_v5(?:_(\d+))?\.xlsx$')
    seen = set()
    for fp in sorted(src_dir.glob('*.xlsx')):
        if fp.name.startswith('_'):
            # index file or similar — keep as-is but rename
            shutil.copy2(fp, dst_dir / fp.name)
            continue
        m = pattern.match(fp.name)
        if not m:
            print(f'  ! skip (no match): {fp.name}')
            continue
        name, birth, suffix = m.group(1), m.group(2), m.group(3)
        cho = to_choseong(name)
        # Disambiguate if same 초성 + birth
        key = (cho, birth)
        if key in seen:
            short = short_hash(name, birth)
            new_name = f'{prefix}_{cho}_{birth}_{short}_v5.xlsx'
        else:
            seen.add(key)
            new_name = f'{prefix}_{cho}_{birth}_v5.xlsx'
        if suffix:
            new_name = new_name.replace('_v5.xlsx', f'_{suffix}_v5.xlsx')
        shutil.copy2(fp, dst_dir / new_name)
        mapping.append({
            'original_filename': fp.name,
            'anonymized_filename': new_name,
            'patient_name': name,
            'patient_choseong': cho,
            'birth_date': birth,
        })
    return mapping

if __name__ == '__main__':
    src = Path(sys.argv[1])
    dst = Path(sys.argv[2])
    mapping_csv = Path(sys.argv[3]) if len(sys.argv) > 3 else None
    rows = process(src, dst)
    print(f'\n✓ {len(rows)} files anonymized → {dst}')
    if mapping_csv:
        mapping_csv.parent.mkdir(parents=True, exist_ok=True)
        with open(mapping_csv, 'w', encoding='utf-8', newline='') as f:
            w = csv.DictWriter(f, fieldnames=['original_filename','anonymized_filename','patient_name','patient_choseong','birth_date'])
            w.writeheader()
            w.writerows(rows)
        print(f'✓ mapping CSV (PRIVATE — DO NOT upload): {mapping_csv}')
