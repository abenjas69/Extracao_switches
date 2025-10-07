from __future__ import annotations
import os
import datetime
import logging
from typing import List, Tuple
from netmiko import ConnectHandler

from .textfsm_utils import detect_ntc_templates_dir, parse_with_textfsm, _dictlist_to_table
from .parsers import (
    parse_show_interfaces_status,
    parse_show_vlan_brief,
    parse_show_inventory,
    parse_cdp_neighbors_detail,
    parse_show_version,
    parser_show_spanning_tree_from_text,
    parse_etherchannel_summary_from_text,
    parse_show_interfaces_trunk,
)

log = logging.getLogger("clean_switch.connector")

COMMANDS_DEFAULT = [
    "show inventory",
    "show version",
    "show interfaces status",
    "show interfaces trunk",
    "show vlan brief",
    "show cdp neighbors detail",
    "show lldp neighbors detail",
    "show etherchannel summary",
    "show spanning-tree",
]


def collect(host: str, username: str, password: str, port: int = 22, commands: List[str] | None = None):
    """SSH ao switch e recolha de outputs (raw + parsed) por comando.

    Fluxo de parsing por comando:
      1) Netmiko `use_textfsm=True`  (mais fiável no teu ambiente)
      2) CliTable (TextFSM clássico via ntc-templates)
      3) Fallbacks regex (nunca ficas com 'texto bruto')

    Returns:
        tuple[str, str, list[tuple[str, str, list[str], list[list[str]]]], str]:
            (hostname, timestamp, collected, outdir)
    """
    # Garantir NTC templates no ambiente (NET_TEXTFSM definido, etc.)
    detect_ntc_templates_dir()

    device = {
        "device_type": "cisco_ios",
        "host": host,
        "username": username,
        "password": password,
        "port": port,
        "fast_cli": True,
    }
    conn = ConnectHandler(**device)

    # "Higiene" de terminal para evitar paginação/wrap que estragam parser
    try:
        conn.send_command("terminal width 511")
        conn.send_command("terminal length 0")
    except Exception as e:
        log.debug("Ignorando erro em terminal width/length: %s", e)

    hostname = conn.find_prompt().strip("#>") or host
    ts = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    outdir = os.path.abspath(os.path.join(os.getcwd(), "outputs", hostname))
    os.makedirs(outdir, exist_ok=True)

    collected: list[tuple[str, str, list[str], list[list[str]]]] = []

    for cmd in (commands or COMMANDS_DEFAULT):
        # 0) Executar comando
        out = conn.send_command(cmd, expect_string=r"[#>]", read_timeout=60)

        headers: list[str] = []
        rows: list[list[str]] = []

        # 1) Tentar primeiro Netmiko com use_textfsm=True
        parsed = None
        try:
            parsed = conn.send_command(cmd, use_textfsm=True)
        except Exception as e:
            log.debug("use_textfsm raised for %r: %s", cmd, e)

        if parsed:
            try:
                h2, r2 = _dictlist_to_table(parsed)
                if h2 and r2:
                    headers, rows = h2, r2
                    log.debug("use_textfsm OK para %r -> %d linhas", cmd, len(rows))
            except Exception as e:
                log.debug("dictlist_to_table falhou para %r: %s", cmd, e)

        # 2) Se ainda não temos tabela, tentar CliTable (TextFSM clássico)
        if not headers or not rows:
            try:
                h3, r3 = parse_with_textfsm(cmd, out)
                if h3 and r3:
                    headers, rows = h3, r3
                    log.debug("CliTable OK para %r -> %d linhas", cmd, len(rows))
                else:
                    log.debug("CliTable MISS para %r", cmd)
            except Exception as e:
                log.debug("CliTable erro para %r: %s", cmd, e)

        # 3) Fallbacks específicos por comando (garantem tabela)
        if not headers or not rows:
            cl = cmd.strip().lower()
            if "show interfaces status" in cl:
                headers, rows = parse_show_interfaces_status(out)
            elif "show vlan brief" in cl:
                headers, rows = parse_show_vlan_brief(out)
            elif "show inventory" in cl:
                headers, rows = parse_show_inventory(out)
            elif "cdp neighbors detail" in cl:
                headers, rows = parse_cdp_neighbors_detail(out)
            elif "show version" in cl:
                headers, rows = parse_show_version(out)
            elif "show spanning-tree" in cl:
                headers, rows = parser_show_spanning_tree_from_text(out)
            elif "show etherchannel" in cl and "summary" in cl:
                headers, rows = parse_etherchannel_summary_from_text(out)
            elif "show interfaces trunk" in cl:
                headers, rows = parse_show_interfaces_trunk(out)

            if headers and rows:
                log.debug("Fallback OK para %r -> %d linhas", cmd, len(rows))
            else:
                log.debug("Fallback MISS para %r (ficará sem tabela)", cmd)

        collected.append((cmd, out, headers, rows))

    conn.disconnect()
    return hostname, ts, collected, outdir
