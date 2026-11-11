"""Recalculate a workbook with headless LibreOffice (for tests and previews).

    python -m tcc.recalc OpenTCC.xlsx outdir [xlsx|pdf]
"""
from __future__ import annotations

import pathlib
import shutil
import subprocess
import sys
import tempfile

_XCU = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry"
 xmlns:xs="http://www.w3.org/2001/XMLSchema"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load">
<prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>
"""


def soffice():
    return shutil.which("soffice") or shutil.which("libreoffice")


def recalc(src, outdir, fmt="xlsx"):
    """Convert src into outdir with every formula recalculated on load."""
    exe = soffice()
    if not exe:
        raise RuntimeError("LibreOffice (soffice) not found")
    outdir = pathlib.Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as prof:
        user = pathlib.Path(prof) / "user"
        user.mkdir()
        (user / "registrymodifications.xcu").write_text(_XCU)
        subprocess.run(
            [exe, f"-env:UserInstallation={pathlib.Path(prof).as_uri()}", "--headless",
             "--convert-to", fmt, "--outdir", str(outdir), str(src)],
            check=True, capture_output=True, timeout=300)
    return outdir / (pathlib.Path(src).stem + "." + fmt)


if __name__ == "__main__":
    print(recalc(sys.argv[1], sys.argv[2], *(sys.argv[3:4])))
