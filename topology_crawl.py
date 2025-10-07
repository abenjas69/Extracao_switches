# clean_switch/topology_crawl.py
from __future__ import annotations
import logging, os, socket
from dataclasses import dataclass
from typing import Dict, List, Optional, Set, Tuple
from ipaddress import ip_address, ip_network

from .network_connector import collect  # SSH + recolha + parsing + outputs dir
from .excel_pipeline import create_excel, metrics_from_collected
from history_json import save_snapshot
from ideia6_features import ideia6_run_pipeline
from .raw_outputs import save_raw_outputs


log = logging.getLogger("clean_switch.crawl")

# --------------------------------------------------------------------------------------
# Tipos base
# --------------------------------------------------------------------------------------

@dataclass(frozen=True)
class DeviceKey:
    serial: Optional[str] = None
    mgmt_ip: Optional[str] = None
    def best(self) -> str:
        return self.serial or self.mgmt_ip or "unknown"

@dataclass
class Neighbor:
    device_id: str
    mgmt_ip: Optional[str]

# --------------------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------------------

def _norm(s) -> str:
    return (str(s) if s is not None else "").strip()

def _in_allowed_subnets(ip: str, allowed: Optional[List[str]]) -> bool:
    if not allowed:
        return True
    try:
        ipx = ip_address(ip)
    except Exception:
        return False
    for cidr in allowed:
        try:
            if ipx in ip_network(cidr, strict=False):
                return True
        except Exception:
            continue
    return False

def _resolve_dns(name: str) -> Optional[str]:
    try:
        return socket.gethostbyname(name)
    except Exception:
        return None

def _headers_index_map(headers: List[str]) -> Dict[str, int]:
    return {str(h).strip().lower(): i for i, h in enumerate(headers or [])}

def _find_col_idx(headers: List[str], *, exact: List[str] = None, contains: List[str] = None) -> Optional[int]:
    """
    Procura o índice de uma coluna por nomes alternativos:
      - 'exact': nomes possíveis (case-insensitive, '_' ~ ' ')
      - 'contains': substrings (case-insensitive)
    """
    if not headers:
        return None
    exact = [e.strip().lower().replace(" ", "_") for e in (exact or [])]
    contains = [c.strip().lower() for c in (contains or [])]
    norm = [str(h).strip().lower() for h in headers]
    norm_underscore = [h.replace(" ", "_") for h in norm]

    # match exato (com '_' ~ espaço)
    for i, h in enumerate(norm_underscore):
        if h in exact:
            return i

    # match por substring
    for i, h in enumerate(norm):
        for c in contains:
            if c in h:
                return i
    return None

def _extract_device_key_from_collected(collected) -> DeviceKey:
    """
    Tenta obter um identificador único do equipamento a partir do 'collected':
      1) serial do 'show version'
      2) serial do 'show inventory'
    """
    serial = None
    for cmd, _raw, headers, rows in collected:
        cl = (cmd or "").strip().lower()
        if "show version" in cl and headers and rows:
            hmap = _headers_index_map(headers)
            i = hmap.get("serial")
            if i is not None and rows and i < len(rows[0]):
                serial = _norm(rows[0][i])
                if serial:
                    break
    if not serial:
        for cmd, _raw, headers, rows in collected:
            cl = (cmd or "").strip().lower()
            if "show inventory" in cl and headers and rows:
                i = _find_col_idx(headers, exact=["serial", "sn"], contains=["serial"])
                if i is not None:
                    for r in rows:
                        sv = _norm(r[i]) if i < len(r) else ""
                        if sv:
                            serial = sv
                            break
            if serial:
                break
    return DeviceKey(serial=serial)

def _extract_neighbors_from_collected(collected) -> List[Neighbor]:
    """
    Lê vizinhos a partir de CDP e LLDP, aceitando múltiplas variantes de headers:
      - Nome: Neighbor_name, Neighbour_name, destination_host, device_id, device,
              remote_system_name, system_name, (e variantes com espaços)
      - IP  : Mgmt_address, management_ip, mgmt_ip, ip_address, management_address,
              (e variantes 'ip address', 'management address', 'mgmt addr')
    CDP tem prioridade; LLDP é fallback.
    """
    def collect_from(cmd_name_match: str) -> List[Neighbor]:
        local: List[Neighbor] = []
        for cmd, _raw, headers, rows in collected:
            if cmd_name_match not in (cmd or "").lower():
                continue
            if not headers or not rows:
                continue

            dev_idx = _find_col_idx(
                headers,
                exact=[
                    "neighbor_name", "Neighbour_name",
                    "destination_host", "device_id", "device",
                    "remote_system_name", "system_name",
                ],
                contains=["device id", "destination", "system name", "remote_system"]
            )
            ip_idx = _find_col_idx(
                headers,
                exact=[
                    "mgmt_address", "Mgmt_address",
                    "management_ip", "mgmt_ip",
                    "ip_address", "management_address",
                ],
                contains=["ip address", "management address", "mgmt addr"]
            )

            for r in rows:
                dev = (str(r[dev_idx]).strip() if (dev_idx is not None and dev_idx < len(r)) else "")
                ipv = (str(r[ip_idx]).strip()  if (ip_idx  is not None and ip_idx  < len(r)) else "")
                # normalizar “not advertised”
                dev = dev if dev and dev.lower() != "not advertised" else ""
                ipv = ipv if ipv and ipv.lower() != "not advertised" else ""
                local.append(Neighbor(device_id=dev, mgmt_ip=(ipv or None)))
        return local

    # 1) CDP primeiro
    out = collect_from("cdp neighbors detail")
    if out:
        return out
    # 2) LLDP fallback
    return collect_from("lldp neighbors detail")

def _cleanup_execucao_keep_newest(xlsx_path: str) -> None:
    """
    Mantém apenas a sheet Execução mais recente (mesma lógica do CLI).
    """
    try:
        from openpyxl import load_workbook
        from .excel_utils import _list_execucao_sheets
        wb = load_workbook(xlsx_path)
        execs = _list_execucao_sheets(wb)
        if execs:
            execs_sorted = sorted(execs, key=lambda t: (t[3] is None, t[3], t[2]))
            newest = execs_sorted[-1][0]
            for ws, _name, _idx, _dt in execs:
                if ws is not newest:
                    wb.remove(ws)
            wb.save(xlsx_path)
    except Exception as e:
        log.error("Falhou limpeza de sheets Execução (%s): %s", xlsx_path, e)

# --------------------------------------------------------------------------------------
# API principal
# --------------------------------------------------------------------------------------

def crawl_topology(
    seed_ip: str,
    username: str,
    password: str,
    enable: Optional[str] = None,
    port: int = 22,
    max_depth: int = 1,
    allowed_subnets: Optional[List[str]] = None,
    dns_fallback: bool = True,
    hostmap: Optional[Dict[str, str]] = None,
    out_dir_override: Optional[str] = None,
) -> List[Tuple[str, str]]:
    """
    Explora a topologia por camadas (BFS) a partir de 'seed_ip', gerando 1 Excel por switch.

    Returns:
        Lista [(hostname, xlsx_path)] para todos os switches processados.
    """
    hostmap = hostmap or {}
    results: List[Tuple[str, str]] = []
    visited_keys: Set[str] = set()   # anti-loop por Serial/DeviceKey
    seen_ips: Set[str] = set()       # evita re-enfileirar o mesmo IP
    queue: List[Tuple[str, int]] = [(seed_ip, 0)]

    while queue:
        ip, depth = queue.pop(0)
        if ip in seen_ips:
            continue
        seen_ips.add(ip)

        if not _in_allowed_subnets(ip, allowed_subnets):
            log.info("Ignorar %s (fora das sub-redes permitidas)", ip)
            continue

        log.info("[Crawl] A ligar a %s (depth=%d)...", ip, depth)
        try:
            # 1) Recolha “normal” (gera também pasta de outputs)
            host, ts, collected, outdir = collect(ip, username, password, port=port, commands=None)
            if out_dir_override:
                outdir = out_dir_override
                os.makedirs(outdir, exist_ok=True)
            xlsx_path = os.path.join(outdir, f"{host}_levantamento.xlsx")

            # 2) Anti-loop (serial/inventory)
            key = _extract_device_key_from_collected(collected)
            key_id = key.best()
            if key_id in visited_keys:
                log.info("Já visitado (serial=%s) em %s — a saltar geração de Excel.", key_id, ip)
            else:
                visited_keys.add(key_id)

                # 3) Snapshot JSON (com métricas) + Excel + Ideia6
                metrics = metrics_from_collected(collected)
                save_snapshot(
                    base_dir=outdir,
                    hostname=host,
                    ts=ts,
                    collected=collected,
                    meta={"host_ip": ip, "metrics": metrics},
                    max_keep=10
                )

                # gravar .txt com texto bruto e cabeçalho para ESTE host ---
                base_dir_txt = out_dir_override if out_dir_override else outdir
                save_raw_outputs(
                    hostname=host,
                    collected=collected,
                    base_dir=base_dir_txt,
                    make_timestamp_subdir=False,
                    ts=ts,
                )

                create_excel(host, ts, collected, xlsx_path)
                try:
                    ideia6_run_pipeline(host, ts, collected, xlsx_path, host_ip=ip)
                except Exception as e:
                    log.error("Ideia6 pipeline falhou para %s: %s", host, e)
                _cleanup_execucao_keep_newest(xlsx_path)

                results.append((host, xlsx_path))
                log.info("✓ Excel gerado: %s", xlsx_path)

            # 4) Descobrir vizinhos e enfileirar (se ainda dentro da profundidade)
            if depth < max_depth:
                neighs = _extract_neighbors_from_collected(collected)
                log.info("[Crawl] %s: encontrados %d vizinhos", host, len(neighs))
                for n in neighs:
                    nip = (n.mgmt_ip or "").strip()
                    source = "CDP/LLDP"
                    if not nip and dns_fallback and n.device_id:
                        resolved = _resolve_dns(n.device_id)
                        if resolved:
                            nip, source = resolved, f"DNS({n.device_id})"
                    if n.device_id in hostmap:
                        nip, source = hostmap[n.device_id], "HOSTMAP"
                    if not nip:
                        log.info("[Crawl]   - %s: sem IP de gestão -> ignorado", (n.device_id or "(sem nome)"))
                        continue
                    if not _in_allowed_subnets(nip, allowed_subnets):
                        log.info("[Crawl]   - %s: IP %s fora das sub-redes permitidas -> ignorado", n.device_id, nip)
                        continue
                    log.info("[Crawl]   - %s: IP %s via %s -> enfileirado (depth %d)", n.device_id or "(sem nome)", nip, source, depth+1)
                    if nip not in seen_ips:
                        queue.append((nip, depth + 1))

        except Exception as e:
            log.error("[Crawl] Falha em %s: %s", ip, e)

    return results
