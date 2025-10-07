# clean_switch/raw_outputs.py
from __future__ import annotations
import os, re
from typing import Iterable, Tuple, Any

CollectedItem = Tuple[str, Any, Any, Any]
_slug_re = re.compile(r"[^a-z0-9_]+")

def _slugify(cmd: str) -> str:
    s = (cmd or "").strip().lower().replace(" ", "_").replace("/", "-")
    s = _slug_re.sub("_", s)
    return s.strip("_") or "unknown_cmd"

def save_raw_outputs(
    hostname: str,
    collected: Iterable[CollectedItem],
    base_dir: str = "outputs",
    make_timestamp_subdir: bool = False,
    ts: str | None = None,
) -> str:
    """
    Grava 1 ficheiro .txt por comando, com o *raw* exatamente como veio do switch.
    Tupla esperada: (cmd, raw, headers, rows) — usa SEMPRE o 2.º elemento ('raw').
    """
    host = hostname or "unknown"
    # Se base_dir já termina no host, não voltar a juntar
    if os.path.basename(os.path.normpath(base_dir)) == host:
        outdir = base_dir
    else:
        outdir = os.path.join(base_dir, host)
        
    if make_timestamp_subdir and ts:
        safe_ts = ts.replace(":", "-").replace(" ", "_")
        outdir = os.path.join(outdir, safe_ts)
    os.makedirs(outdir, exist_ok=True)

    for (cmd, raw, _headers, _rows) in collected:
        fname = f"{host}_{_slugify(cmd)}.txt"
        fpath = os.path.join(outdir, fname)

        header = (
            f"# Hostname: {host}\n"
            f"# Comando: {cmd}\n"
            f"# Timestamp: {ts}\n"
            f"# ---\n"
        )   

        if isinstance(raw, (bytes, bytearray)):
            # se raw for bytes, prepend também o header em utf-8
            with open(fpath, "wb") as f:
                f.write(header.encode("utf-8"))
                f.write(raw)
        elif isinstance(raw, str):
            with open(fpath, "w", encoding="utf-8", newline="") as f:
                f.write(header)
                f.write(raw)
        elif isinstance(raw, (list, tuple)):
            with open(fpath, "w", encoding="utf-8", newline="") as f:
                f.write(header)
                f.write("\n".join(str(x) for x in raw))
        else:
            with open(fpath, "w", encoding="utf-8", newline="") as f:
                f.write(header)
                f.write("" if raw is None else str(raw))

    return outdir
