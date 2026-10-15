"""Restore chunked source files and verify immutable original hashes."""
from pathlib import Path
import hashlib
import json

def main():
    root = Path(__file__).resolve().parent
    catalog = json.loads((root / 'source-catalog.json').read_text(encoding='utf-8'))
    for item in catalog['files']:
        path = root / item['file']
        if not path.exists() and item.get('chunks'):
            data = b''.join((root / part).read_bytes() for part in item['chunks'])
            if hashlib.sha256(data).hexdigest() != item['sha256']:
                raise SystemExit('Chunk hash mismatch: ' + item['file'])
            with path.open('xb') as output:
                output.write(data)
        if hashlib.sha256(path.read_bytes()).hexdigest() != item['sha256']:
            raise SystemExit('Original hash mismatch: ' + item['file'])
        print('VERIFIED ' + item['file'])

if __name__ == '__main__':
    main()
