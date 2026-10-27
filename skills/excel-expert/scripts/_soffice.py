"""Utilidades compartidas para ejecutar LibreOffice en modo headless.

- Localiza el binario (soffice / libreoffice / rutas típicas de macOS y Windows).
- Usa un perfil de usuario aislado y desechable (evita bloqueos con instancias abiertas).
- Configura variables de entorno necesarias en entornos sin pantalla (contenedores, CI).
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

CANDIDATES = [
    "soffice",
    "libreoffice",
    "/usr/bin/soffice",
    "/usr/lib/libreoffice/program/soffice",
    "/opt/libreoffice/program/soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
]


def find_soffice() -> str | None:
    env = os.environ.get("SOFFICE_PATH")
    if env and Path(env).exists():
        return env
    for c in CANDIDATES:
        p = shutil.which(c) if not os.path.isabs(c) else (c if Path(c).exists() else None)
        if p:
            return p
    return None


def soffice_env() -> dict:
    env = os.environ.copy()
    # Sin servidor gráfico: backend "svp" (headless puro).
    env.setdefault("SAL_USE_VCLPLUGIN", "svp")
    env.setdefault("HOME", tempfile.gettempdir())
    return env


_RECALC_XCU = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry"
           xmlns:xs="http://www.w3.org/2001/XMLSchema"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>
"""


def new_profile_dir(recalc_always: bool = True) -> Path:
    """Crea un perfil temporal. Con recalc_always=True, LibreOffice recalcula TODAS las
    fórmulas al abrir ficheros OOXML/ODF (por defecto solo calcula las que no tienen valor
    cacheado, y podría dejar valores obsoletos)."""
    profile = Path(tempfile.mkdtemp(prefix="lo_profile_"))
    if recalc_always:
        user = profile / "user"
        user.mkdir(parents=True, exist_ok=True)
        (user / "registrymodifications.xcu").write_text(_RECALC_XCU, encoding="utf-8")
    return profile


def profile_arg(profile: Path) -> str:
    return "-env:UserInstallation=" + profile.resolve().as_uri()


def convert(src: Path, out_dir: Path, target: str, timeout: int = 120,
            infilter: str | None = None) -> Path:
    """Convierte `src` al formato `target` (p.ej. 'xlsx', 'pdf', 'csv') con LibreOffice.

    Devuelve la ruta del fichero generado. Lanza RuntimeError si falla.
    """
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError(
            "LibreOffice no encontrado. Instálalo (apt install libreoffice-calc / "
            "brew install --cask libreoffice) o define SOFFICE_PATH."
        )
    out_dir.mkdir(parents=True, exist_ok=True)
    profile = new_profile_dir()
    cmd = [soffice, profile_arg(profile), "--headless", "--norestore", "--nolockcheck", "--calc"]
    if infilter:
        cmd.append(f"--infilter={infilter}")
    cmd += ["--convert-to", target, "--outdir", str(out_dir), str(src)]
    try:
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout,
                             env=soffice_env())
    except subprocess.TimeoutExpired:
        raise RuntimeError(f"LibreOffice superó el timeout de {timeout}s")
    finally:
        shutil.rmtree(profile, ignore_errors=True)
    ext = target.split(":")[0]
    produced = out_dir / (src.stem + "." + ext)
    if res.returncode != 0 or not produced.exists():
        raise RuntimeError(
            f"La conversión falló (rc={res.returncode}).\nstdout: {res.stdout}\nstderr: {res.stderr}"
        )
    return produced


if __name__ == "__main__":
    print(find_soffice() or "LibreOffice no encontrado", file=sys.stderr)
