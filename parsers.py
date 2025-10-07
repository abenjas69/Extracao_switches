# clean_switch/parsers.py
from __future__ import annotations
from typing import List, Tuple
import re
import ast

# =====================================================================
# Helpers
# =====================================================================

def _norm(s) -> str:
    return (str(s) if s is not None else "").strip()

def _collapse_spaces(s: str) -> str:
    return re.sub(r"[ \t]{2,}", " ", s.strip())

def _strip_ansi(s: str) -> str:
    # remove códigos ANSI se existirem
    return re.sub(r"\x1B\[[0-9;]*[A-Za-z]", "", s or "")

def _clean_text(raw: str) -> str:
    if not raw:
        return ""
    s = raw.replace("\r\n", "\n").replace("\r", "\n")
    s = _strip_ansi(s)
    return s

# =====================================================================
# show interfaces status  (IOS clássico)
# Cabeçalho típico:
# Port      Name               Status       Vlan       Duplex  Speed Type
# Gi1/0/1   Desc Opcional      connected    trunk      a-full a-1000 10/100/1000BaseTX
# =====================================================================

def parse_show_interfaces_status(raw_text: str) -> Tuple[List[str], List[List[str]]]:
    if not raw_text:
        return ([], [])
    t = _clean_text(raw_text)
    lines = [ln for ln in t.split("\n") if ln.strip()]

    # encontrar a linha de cabeçalho
    hdr_idx = None
    for i, ln in enumerate(lines):
        if re.search(r"\bPort\b.*\bStatus\b.*\bVlan\b", ln):
            hdr_idx = i
            break
    if hdr_idx is None:
        return ([], [])

    # determinar colunas por posições a partir do header “tabelado”
    header = lines[hdr_idx]
    # mapear nomes-alvo
    headers = ["Port", "Name", "Status", "Vlan", "Duplex", "Speed", "Type"]

    # calcular “cortes” por colunas (posições de início dos rótulos)
    def _col_starts(h: str, keys: List[str]) -> List[int]:
        starts = []
        for k in keys:
            m = re.search(rf"\b{k}\b", h)
            starts.append(m.start() if m else None)
        return starts

    starts = _col_starts(header, headers)
    # fallback bruto: se falhar deteção por posição, usa split por múltiplos espaços
    if any(s is None for s in starts):
        rows = []
        for ln in lines[hdr_idx + 1:]:
            # parar ao ver nova sessão/prompt
            if re.match(r"^\S+#\s*$", ln):
                break
            # cada linha “completa” costuma ter 7 campos; o 'Name' pode ter espaços.
            # Estratégia: capturar 1º token (Port), 3 últimos (Duplex, Speed, Type), e deduzir o meio.
            ln2 = _collapse_spaces(ln)
            parts = ln2.split(" ")
            if len(parts) < 5:
                continue
            port = parts[0]
            # últimos 3 tokens
            typ = parts[-1]
            spd = parts[-2]
            dpx = parts[-3]
            # o resto: Name, Status, Vlan — normalmente 3 tokens mínimos; se mais, o “Name” absorve o excesso
            middle = parts[1:-3]
            if len(middle) < 2:
                # sem dados suficientes
                continue
            # Status e Vlan são os 2 últimos do “middle”
            if len(middle) >= 2:
                vlan = middle[-1]
                status = middle[-2]
                name = " ".join(middle[:-2]) if len(middle) > 2 else ""
            else:
                status, vlan, name = "", "", " ".join(middle)

            rows.append([port, name, status, vlan, dpx, spd, typ])
        return (headers, rows)

    # parser por faixas (posições fixas estimadas)
    # criar “slices” seguros com base no índice do seguinte cabeçalho (ou fim)
    idxs = [i for i in starts if i is not None]
    if not idxs:
        return ([], [])
    order = sorted([(starts[i], i) for i in range(len(headers)) if starts[i] is not None])
    cuts = [p for p, _ in order] + [None]  # último “None” = até ao fim

    def _slice(line, i):
        a = cuts[i]
        b = cuts[i+1]
        if a is None:
            return ""
        return line[a:b].rstrip() if b is not None else line[a:].rstrip()

    rows: List[List[str]] = []
    for ln in lines[hdr_idx + 1:]:
        if re.match(r"^\S+#\s*$", ln):
            break
        # ignorar separadores ou linhas “estranhas”
        if not ln.strip() or set(ln.strip()) == {"-"}:
            continue
        vals = [_slice(ln, i) for i in range(len(headers))]
        # validar porto (1ª coluna tem de ter algo tipo Gi1/0/.., Fa.., Te.. ou Po..)
        if not vals[0].strip():
            continue
        rows.append([_collapse_spaces(v) for v in vals])

    return (headers, rows)

# =====================================================================
# show vlan brief
# Formato clássico com possíveis quebras de linha na coluna Ports
# =====================================================================

def parse_show_vlan_brief(raw_text: str) -> Tuple[List[str], List[List[str]]]:
    if not raw_text:
        return ([], [])
    t = _clean_text(raw_text)
    lines = [ln.rstrip() for ln in t.split("\n")]

    # localizar início da grelha
    start = None
    for i, ln in enumerate(lines):
        if re.search(r"\bVLAN\s+Name\s+Status\s+Ports", ln, re.I):
            start = i + 1
            break
    if start is None:
        return ([], [])

    headers = ["vlan_id", "name", "status", "interfaces"]
    rows: List[List[str]] = []

    cur = None
    for ln in lines[start:]:
        if not ln.strip():
            continue
        if re.match(r"^-{3,}", ln):
            continue
        # linha principal de VLAN começa por dígitos ou pelos VLANs reservados 1002-1005 (act/unsup)
        m = re.match(r"^\s*(\d+)\s+(\S.*?)\s{2,}(\S+)\s*(.*)$", ln)
        if m:
            if cur:
                rows.append(cur)
            vlan_id = m.group(1).strip()
            name = m.group(2).strip()
            status = m.group(3).strip()
            ports = m.group(4).strip()
            cur = [vlan_id, name, status, ports]
        else:
            # linhas de continuação dos ports (identadas)
            if cur and re.match(r"^\s{10,}\S", ln):
                cur[3] = (cur[3] + " " + ln.strip()).strip()

    if cur:
        rows.append(cur)

    # normalizar separador de interfaces (virgulas + espaço)
    out = []
    for r in rows:
        ports = r[3]
        toks = [p.strip() for p in re.split(r"[,\s]+", ports) if p.strip()]
        r[3] = ", ".join(toks)
        out.append(r)

    return (headers, out)

# =====================================================================
# show inventory
# Blocos:
# NAME: "1", DESCR: "WS-C3750G-48TS"
# PID: WS-C3750G-48TS-S  , VID: V02  , SN: FOC1002Y3SE
# =====================================================================

def parse_show_inventory(raw_text: str) -> Tuple[List[str], List[List[str]]]:
    if not raw_text:
        return ([], [])
    t = _clean_text(raw_text)
    blocks = re.split(r"\n\s*\n", t.strip())
    headers = ["name", "descr", "pid", "vid", "serial"]
    rows: List[List[str]] = []

    for b in blocks:
        name = descr = pid = vid = serial = ""
        # linha 1
        m1 = re.search(r'NAME:\s*"([^"]*)",\s*DESCR:\s*"([^"]*)"', b, re.I)
        if m1:
            name = m1.group(1).strip()
            descr = m1.group(2).strip()
        # linha 2
        m2 = re.search(r"PID:\s*([^\s,]+).*?VID:\s*([^\s,]+).*?SN:\s*([^\s,]+)", b, re.I)
        if m2:
            pid = m2.group(1).strip()
            vid = m2.group(2).strip()
            serial = m2.group(3).strip()
        if any([name, descr, pid, vid, serial]):
            rows.append([name, descr, pid, vid, serial])

    return (headers, rows)

# =====================================================================
# show cdp neighbors detail
# Delimitado por “-------------------------”
# =====================================================================

def parse_cdp_neighbors_detail(raw_text: str) -> Tuple[List[str], List[List[str]]]:
    if not raw_text:
        return ([], [])
    t = _clean_text(raw_text)
    chunks = [c for c in re.split(r"-{5,}\s*", t) if c.strip()]

    headers = [
        "device_id", "ip", "platform", "capabilities",
        "local_interface", "port_id", "holdtime", "version"
    ]
    rows: List[List[str]] = []

    for ch in chunks:
        dev = ip = plat = caps = l_if = port_id = hold = ver = ""

        m = re.search(r"Device\s*ID\s*:\s*(.+)", ch, re.I)
        if m:
            dev = _norm(m.group(1))

        # primeiro IP
        mi = re.search(r"IP address:\s*([0-9.]+)", ch, re.I)
        if mi:
            ip = mi.group(1)

        mp = re.search(r"Platform:\s*(.+?),\s*Capabilities:\s*(.+)", ch, re.I)
        if mp:
            plat = _norm(mp.group(1))
            caps = _norm(mp.group(2))

        ml = re.search(r"Interface:\s*([^,]+),\s*Port ID.*?:\s*([^\n]+)", ch, re.I)
        if ml:
            l_if = _norm(ml.group(1))
            port_id = _norm(ml.group(2))

        mh = re.search(r"Holdtime\s*:\s*(\d+)\s*sec", ch, re.I)
        if mh:
            hold = mh.group(1)

        # versão: bloco a seguir a “Version :”
        mv = re.search(r"Version\s*:\s*(.+?)(?:\n\s*\n|$)", ch, re.I | re.S)
        if mv:
            ver = _norm(mv.group(1))

        if any([dev, ip, plat, caps, l_if, port_id, hold, ver]):
            rows.append([dev, ip, plat, caps, l_if, port_id, hold, ver])

    return (headers, rows)

# =====================================================================
# show version (subset útil e estável)
# =====================================================================

def parse_show_version(raw_text: str) -> Tuple[List[str], List[List[str]]]:
    if not raw_text:
        return ([], [])
    t = _clean_text(raw_text)

    # hostname pode não vir no 'show version' — deixamos vazio se não existir
    hostname = ""
    # versão
    mv = re.search(r"Version\s+([A-Za-z0-9.\(\)-]+)", t, re.I)
    version = mv.group(1) if mv else ""
    # plataforma/modelo
    mp = re.search(r"cisco\s+([A-Z0-9\-]+)\s*\(", t, re.I)
    platform = mp.group(1) if mp else ""
    # serial
    ms = re.search(r"System serial number\s*:\s*([A-Za-z0-9]+)", t, re.I)
    serial = ms.group(1) if ms else ""
    # uptime
    mu = re.search(r"\buptime\s+is\s+([^\n]+)", t, re.I)
    uptime = mu.group(1).strip() if mu else ""

    headers = ["hostname", "version", "platform", "serial", "uptime"]
    rows = [[hostname, version, platform, serial, uptime]]
    return (headers, rows)

# =====================================================================
# show spanning-tree (o teu já existia – deixo aqui para centralizar)
# =====================================================================

def parser_show_spanning_tree_from_text(raw_text: str) -> Tuple[List[str], List[List[str]]]:
    if not raw_text:
        return ([], [])
    t = _clean_text(raw_text)
    headers = ["vlan", "interface", "role", "state", "cost", "port_id", "port_type"]
    rows: List[List[str]] = []

    vlan = None
    vlan_hdr = re.compile(r"^\s*VLAN\s*0*(\d+)\b", re.I)
    port_re = re.compile(
        r'^(?P<intf>\S+)\s+(?P<role>\w+)\s+(?P<state>\w+)\s+(?P<cost>\d+)\s+(?P<portid>[\d\.]+)\s+(?P<ptype>.+?)\s*$',
        re.I
    )
    for line in t.split("\n"):
        m_vlan = vlan_hdr.search(line)
        if m_vlan:
            vlan = m_vlan.group(1)
            continue
        m_port = port_re.match(line.strip())
        if m_port and vlan:
            rows.append([
                vlan,
                m_port.group("intf"),
                m_port.group("role").lower(),
                m_port.group("state").lower(),
                m_port.group("cost"),
                m_port.group("portid"),
                m_port.group("ptype"),
            ])
    return (headers, rows)

# =====================================================================
# show etherchannel summary (o teu já existia – centralizo aqui)
# =====================================================================

def parse_etherchannel_summary_from_text(raw_text: str) -> Tuple[List[str], List[List[str]]]:
    if not raw_text:
        return ([], [])
    s = _clean_text(raw_text)
    s = re.sub(r"[ ]{2,}", " ", s)
    lines = [ln for ln in s.split("\n") if ln.strip()]

    # localizar header
    start = None
    for i, ln in enumerate(lines):
        if re.search(r"\bGroup\b.*\bPort-?Channel\b.*\bProtocol\b", ln, re.I):
            start = i + 1
            break
    if start is None:
        return ([], [])

    data_line_re = re.compile(
        r"^\s*(?P<group>\d+)\s+(?P<po>[A-Za-z]+[0-9]+)\((?P<flags>[^)]+)\)\s+(?P<protocol>\S+)\s*(?P<ports>.*)$"
    )
    entries = []
    cur = None
    for ln in lines[start:]:
        m = data_line_re.match(ln)
        if m:
            if cur:
                entries.append(cur)
            cur = {
                "group": m.group("group"),
                "port_channel": m.group("po"),
                "flags": m.group("flags"),
                "protocol": m.group("protocol"),
                "member_ports": (m.group("ports") or "").strip(),
            }
        else:
            if cur and not re.match(r"^\s*\d+\s+", ln):
                cur["member_ports"] = (cur["member_ports"] + " " + ln.strip()).strip()
    if cur:
        entries.append(cur)

    rows: List[List[str]] = []
    for e in entries:
        toks = [p for p in re.split(r"[,\s]+", e["member_ports"]) if p]
        status = "Up" if "U" in e["flags"].upper() else ("Down" if "D" in e["flags"].upper() else "Unknown")
        rows.append([
            int(e["group"]),
            e["port_channel"],
            e["protocol"],
            status,
            e["flags"],
            ", ".join(toks),
        ])
    headers = ["Group", "Port-Channel", "Protocol", "Status", "Flags", "Member Ports"]
    return (headers, rows)


def parse_show_interfaces_trunk(raw: str):
    """
    Fallback parser para 'show interfaces trunk' (Catalyst IOS).
    Extrai:
      - port, mode, encapsulation, status, native_vlan
      - allowed_vlans
      - allowed_active_vlans
      - stp_forwarding_not_pruned
    Retorna: (headers, rows)
    """
    if not raw:
        headers = ["port","mode","encapsulation","status","native_vlan",
                   "allowed_vlans","allowed_active_vlans","stp_forwarding_not_pruned"]
        return headers, []

    lines = [ln.rstrip() for ln in raw.splitlines() if ln.strip()]
    ports = {}  # port -> dict

    # ---------- Secção 1: tabela principal ----------
    header_idx = None
    for i, ln in enumerate(lines):
        low = ln.lower()
        # linha típica: "Port      Mode         Encapsulation  Status        Native vlan"
        if low.startswith("port") and "mode" in low and "status" in low:
            header_idx = i
            break

    if header_idx is not None:
        for ln in lines[header_idx + 1:]:
            low = ln.lower()
            # As secções seguintes começam com "Port Vlans ..."
            if low.startswith("port ") and "vlans" in low:
                break
            parts = re.split(r"\s{2,}", ln.strip())
            # formatos comuns:
            # 5 colunas: port, mode, encapsulation, status, native_vlan
            # 4 colunas: port, mode, status, native_vlan (sem encapsulation)
            if len(parts) >= 5:
                port, mode, encap, status, native = parts[:5]
            elif len(parts) == 4:
                port, mode, status, native = parts
                encap = ""
            else:
                continue
            d = ports.setdefault(port, {})
            d.update({
                "port": port,
                "mode": mode,
                "encapsulation": encap,
                "status": status,
                "native_vlan": native,
            })

    # ---------- Secções 2-4: listas de VLAN por porta ----------
    section_map = {
        "vlans allowed on trunk": "allowed_vlans",
        "vlans allowed and active in management domain": "allowed_active_vlans",
        "vlans in spanning tree forwarding state and not pruned": "stp_forwarding_not_pruned",
    }

    current_field = None
    for ln in lines:
        low = ln.lower()
        # Detecta cabeçalhos de secção
        for needle, field in section_map.items():
            if low.startswith("port ") and needle in low:
                current_field = field
                break
        else:
            # Se estivermos dentro de uma secção, esperar linhas no formato:
            # "<Port>    <lista>"
            if current_field:
                m = re.match(r"^(\S+)\s{2,}(.+)$", ln.strip())
                if m:
                    port, val = m.group(1), m.group(2).strip()
                    d = ports.setdefault(port, {"port": port})
                    d[current_field] = val

    headers = [
        "port",
        "mode",
        "encapsulation",
        "status",
        "native_vlan",
        "allowed_vlans",
        "allowed_active_vlans",
        "stp_forwarding_not_pruned",
    ]
    rows = []
    # Ordena por nome de porta para estabilidade do output
    for port in sorted(ports.keys()):
        d = ports[port]
        rows.append([
            d.get("port", ""),
            d.get("mode", ""),
            d.get("encapsulation", ""),
            d.get("status", ""),
            d.get("native_vlan", ""),
            d.get("allowed_vlans", ""),
            d.get("allowed_active_vlans", ""),
            d.get("stp_forwarding_not_pruned", ""),
        ])

    return headers, rows