"""Build a deterministic installable skill ZIP, excluding generated caches."""
import argparse
import hashlib
from pathlib import Path
import zipfile


def package(root, output):
    root, output = Path(root).resolve(), Path(output).resolve()
    if not (root/'SKILL.md').is_file():
        raise ValueError('Input directory must contain SKILL.md')
    if root == output or root in output.parents:
        raise ValueError('Save the archive outside the skill directory')
    output.parent.mkdir(parents=True,exist_ok=True)
    files = sorted(p for p in root.rglob('*') if p.is_file() and not any(v in ('.git','__pycache__','.venv','node_modules') for v in p.relative_to(root).parts) and p.suffix not in ('.pyc','.zip'))
    with zipfile.ZipFile(output,'w',compression=zipfile.ZIP_DEFLATED) as archive:
        for p in files:
            entry = zipfile.ZipInfo(f'{root.name}/{p.relative_to(root).as_posix()}',date_time=(2026,1,1,0,0,0))
            entry.compress_type = zipfile.ZIP_DEFLATED
            entry.external_attr = 0o100644 << 16
            archive.writestr(entry,p.read_bytes())
    with zipfile.ZipFile(output) as archive:
        assert archive.testzip() is None
        for p in files:
            assert archive.read(f'{root.name}/{p.relative_to(root).as_posix()}') == p.read_bytes()
    return {'path':str(output),'files':len(files),'sha256':hashlib.sha256(output.read_bytes()).hexdigest()}


if __name__ == '__main__':
    import json
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('skill_directory')
    parser.add_argument('output_zip')
    args = parser.parse_args()
    print(json.dumps(package(args.skill_directory,args.output_zip),indent=2))
