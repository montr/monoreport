"""Exercise Sample 5 through the supplied REST API and retain raw responses."""
import email.policy
from email.parser import BytesParser
import json
from pathlib import Path
import subprocess
import time

BASE = Path(__file__).resolve().parent
EVIDENCE = BASE / 'evidence'
TEMPLATE = BASE / 'output/skill-sample-5.template.xlsx'
CASES = {'original': BASE / 'output/skill-sample-5.json',
         **{name: EVIDENCE / f'{name}-data.json' for name in
            ['modified', 'single-supplier', 'empty-items', 'empty-lot', 'empty-suppliers']}}
results = {}
for name, data in CASES.items():
    headers = EVIDENCE / f'{name}-render.headers'
    response = EVIDENCE / f'{name}-response.mime'
    started = time.monotonic()
    completed = subprocess.run(['curl', '-sS', '--max-time', '30', '-D', str(headers),
                                '-H', 'Accept: multipart/mixed', '-F', f'template=@{TEMPLATE}',
                                '-F', f'data=<{data}', 'http://localhost:3000/api/documents/render',
                                '-o', str(response)], check=True)
    header_bytes = headers.read_bytes()
    status = int(header_bytes.splitlines()[0].split()[1])
    results[name] = {'http_status': status, 'elapsed_seconds': round(time.monotonic() - started, 4)}
    if status != 200:
        error = json.loads(response.read_bytes())
        (EVIDENCE / f'{name}-error.json').write_text(json.dumps(error, indent=2) + '\n')
        results[name]['error'] = error
        continue
    mime_headers = header_bytes.split(b'\n', 1)[1].rstrip(b'\r\n')
    message = BytesParser(policy=email.policy.default).parsebytes(mime_headers + b'\r\n\r\n' + response.read_bytes())
    assert message.is_multipart() and not message.defects, message.defects
    parts = list(message.iter_parts())
    assert len(parts) == 2
    documents = [part for part in parts if part.get_content_type() == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']
    diagnostics = [part for part in parts if part.get_content_type() == 'application/json']
    assert len(documents) == len(diagnostics) == 1
    payload = documents[0].get_payload(decode=True)
    warnings = json.loads(diagnostics[0].get_payload(decode=True))
    assert payload.startswith(b'PK') and warnings['format'] == 'xlsx'
    (EVIDENCE / f'{name}-diagnostics.json').write_text(json.dumps(warnings, indent=2) + '\n')
    destination = BASE / 'output/skill-sample-5.result.xlsx' if name == 'modified' else EVIDENCE / f'{name}.result.xlsx'
    destination.write_bytes(payload)
    results[name]['warnings'] = warnings['warnings']
(EVIDENCE / 'api-cases.json').write_text(json.dumps(results, indent=2) + '\n')
print(json.dumps(results, indent=2))
