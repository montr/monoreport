"""Run in an isolated copy of research-spikes; no GUI/network/build steps hidden here."""
import json
import subprocess
import sys
from pathlib import Path

base=Path(__file__).resolve().parent
def run(*args):
    subprocess.run(args,cwd=base,check=True)

run(sys.executable,str(base/"nested.py"),"prepare")
run(sys.executable,str(base/"nested.py"),"dates")
internal=base/"bin/Debug/net10.0/InternalSpike.dll"
closed=base/"bin/Debug/net10.0/Spike.dll"
outputs=[]
total=0
backends=[("internal",internal)] if "--internal-only" in sys.argv else [("internal",internal),("closed",closed)]
for case in ["full","one-criterion","seven-criteria","no-criteria","no-companies"]:
    for backend,assembly in backends:
        run("dotnet",str(assembly),"transform",str(base/f"nested-evidence/{case}-requests.json"),str(base/f"nested-evidence/{case}-{backend}-formulas.json"))
        outputs.append(str(base/f"nested-evidence/{case}-{backend}.xlsx"))
    total+=len(json.loads((base/f"nested-evidence/{case}-requests.json").read_text()))
run(sys.executable,str(base/"nested.py"),"finish",*(["--internal-only"] if "--internal-only" in sys.argv else []))
run("dotnet",str(internal),"validate",str(base/"nested-evidence/nested-template.xlsx"),*outputs,str(base/"nested-evidence/date-1900.xlsx"),str(base/"nested-evidence/date-1904.xlsx"))
deps=json.loads((base/"bin/Debug/net10.0/InternalSpike.deps.json").read_text())
assert not any("ClosedXML" in name for name in deps["libraries"])
print("Internal runtime dependency manifest has no ClosedXML library. Nested formulas processed per selected backend:",total)
