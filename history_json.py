# -*- coding: utf-8 -*-
# history_json.py
#
# Utilitários para guardar/ler "snapshots" (as tuplas (cmd, raw, headers, rows))
# em JSON, com retenção e APIs simples para obter as duas últimas execuções.

import os, json, tempfile, shutil
from typing import List, Tuple, Dict, Any, Optional
import logging

CollectedType = List[Tuple[str, str, List[str], List[List[str]]]]

def _host_hist_dir(base_dir: str, hostname: str) -> str:
    """Pasta de histórico por host: <base_dir>/_history/<hostname>"""
    p = os.path.join(base_dir, "_history", hostname)
    os.makedirs(p, exist_ok=True)
    return p

def snapshot_path(base_dir: str, hostname: str, ts: str) -> str:
    """Caminho do ficheiro JSON para um timestamp (YYYY-mm-dd_HH-MM-SS)."""
    return os.path.join(_host_hist_dir(base_dir, hostname), f"{ts}.json")

def list_snapshots(base_dir: str, hostname: str) -> List[str]:
    """Lista de caminhos (ordenados por ts) dos snapshots existentes do host."""
    d = _host_hist_dir(base_dir, hostname)
    files = [os.path.join(d, f) for f in os.listdir(d) if f.endswith(".json")]
    files.sort(key=lambda p: os.path.basename(p).split(".json")[0])
    return files

def load_snapshot(path: str) -> Dict[str, Any]:
    """Lê um snapshot JSON do disco."""
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def _to_jsonable(collected: CollectedType) -> List[Dict[str, Any]]:
    """Converte a lista de tuplos para uma estrutura serializável."""
    out: List[Dict[str, Any]] = []
    for cmd, raw, headers, rows in collected:
        out.append({
            "cmd": cmd,
            "raw": raw,
            "headers": headers or [],
            "rows": rows or []
        })
    return out

def save_snapshot(base_dir: str, hostname: str, ts: str,
                  collected: CollectedType, meta: Optional[Dict[str, Any]] = None,
                  max_keep: int = 10) -> str:
    """
    Guarda o snapshot atual como JSON (write-then-rename para segurança).
    Retém apenas os 'max_keep' mais recentes.
    Devolve o caminho final do ficheiro JSON criado.
    """
    dst = snapshot_path(base_dir, hostname, ts)
    tmp_fd, tmp_path = tempfile.mkstemp(prefix="snap_", suffix=".json",
                                        dir=_host_hist_dir(base_dir, hostname))
    os.close(tmp_fd)
    data = {
        "hostname": hostname,
        "timestamp": ts,
        "items": _to_jsonable(collected),
        "meta": meta or {}
    }
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    # gravação atómica
    shutil.move(tmp_path, dst)

    # retenção
    files = list_snapshots(base_dir, hostname)
    if len(files) > max_keep:
        for old in files[0:len(files)-max_keep]:
            try:
                os.remove(old)
            except OSError:
                pass
    return dst

def get_last_two(base_dir: str, hostname: str) -> Tuple[Optional[Dict[str, Any]], Optional[Dict[str, Any]]]:
    """
    Devolve (prev, curr) snapshots como dicionários, ou (None, None) se não houver.
    """
    files = list_snapshots(base_dir, hostname)
    if not files:
        return None, None
    if len(files) == 1:
        return None, load_snapshot(files[-1])
    return load_snapshot(files[-2]), load_snapshot(files[-1])

def simple_diff(prev: Dict[str, Any], curr: Dict[str, Any]) -> Dict[str, Any]:
    """
    Diferença minimalista entre dois snapshots:
      - nº linhas por comando,
      - conjunto de VLANs observadas,
      - contagem de interfaces 'connected' (aprox).
    Isto serve apenas como apoio rápido; a tua Ideia 6 continua a liderar a comparação Excel↔Excel.
    """
    def by_cmd_map(snap):
        return {item["cmd"].lower(): item for item in snap.get("items", [])}

    dif: Dict[str, Any] = {}
    if not prev or not curr:
        return dif

    pmap, cmap = by_cmd_map(prev), by_cmd_map(curr)

    # nº linhas por comando
    dif["rows_delta"] = {}
    for cmd in set(pmap.keys()) | set(cmap.keys()):
        p = len(pmap.get(cmd, {}).get("rows", []))
        c = len(cmap.get(cmd, {}).get("rows", []))
        dif["rows_delta"][cmd] = {"prev": p, "curr": c, "delta": c - p}

    # VLANs observadas (por 'show vlan brief' e/ou colunas VLAN em 'interfaces status')
    def extract_vlans(item: Dict[str, Any]) -> set:
        vlans = set()
        headers = [h.lower() for h in item.get("headers", [])]
        rows = item.get("rows", [])
        if "show vlan brief" in item.get("cmd", "").lower():
            idx = headers.index("vlan") if "vlan" in headers else None
            if idx is not None:
                for r in rows:
                    v = str(r[idx]).strip()
                    if v.isdigit():
                        vlans.add(v)
        if "show interfaces status" in item.get("cmd", "").lower():
            if "vlan" in headers:
                idx = headers.index("vlan")
                for r in rows:
                    v = str(r[idx]).strip()
                    if v and v.isdigit():
                        vlans.add(v)
        return vlans

    def vlanset(snap):
        all_v = set()
        for it in snap.get("items", []):
            all_v |= extract_vlans(it)
        return all_v

    prev_v, curr_v = vlanset(prev), vlanset(curr)
    dif["vlans_added"] = sorted(curr_v - prev_v, key=int)
    dif["vlans_removed"] = sorted(prev_v - curr_v, key=int)

    # 'connected' em interfaces (aproximação)
    def count_connected(item: Dict[str, Any]) -> int:
        headers = [h.lower() for h in item.get("headers", [])]
        rows = item.get("rows", [])
        if "status" not in headers:
            return 0
        sidx = headers.index("status")
        n = 0
        for r in rows:
            st = str(r[sidx]).lower().strip() if sidx < len(r) else ""
            if st in ("connected", "up"):
                n += 1
        return n

    p_conn = count_connected(pmap.get("show interfaces status", {}))
    c_conn = count_connected(cmap.get("show interfaces status", {}))
    dif["interfaces_connected_delta"] = {"prev": p_conn, "curr": c_conn, "delta": c_conn - p_conn}

    return dif

