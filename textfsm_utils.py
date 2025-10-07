# clean_switch/textfsm_utils.py
from __future__ import annotations

import logging
import os
import re
from typing import Dict, Iterable, List, Sequence, Tuple, Any

log = logging.getLogger("clean_switch.textfsm")


# =============================================================================
# Ambiente / descoberta do diretório de templates (ntc-templates)
# =============================================================================

def detect_ntc_templates_dir() -> str | None:
    """Garante que a variável de ambiente NET_TEXTFSM aponta para os ntc-templates.

    - Se NET_TEXTFSM já estiver definida, valida a existência do ficheiro 'index'.
    - Caso contrário, tenta descobrir a pasta a partir do pacote 'ntc_templates'
      e define NET_TEXTFSM para esta sessão.
    - Escreve logs DEBUG com o caminho final e se o 'index' existe.

    Returns:
        Caminho para a pasta de templates (ou None se não encontrado).
    """
    # 1) Se já existe, validar
    env_path = os.environ.get("NET_TEXTFSM")
    if env_path:
        index_ok = os.path.exists(os.path.join(env_path, "index"))
        log.debug("NET_TEXTFSM existente: %s (index=%s)", env_path, index_ok)
        if index_ok:
            return env_path
        else:
            log.debug("NET_TEXTFSM definido mas sem 'index' — tentarei descobrir via ntc_templates.")

    # 2) Descobrir via ntc_templates
    try:
        import ntc_templates  # type: ignore
        import pathlib

        p = pathlib.Path(ntc_templates.__file__).with_name("templates")
        if p.exists() and (p / "index").exists():
            os.environ["NET_TEXTFSM"] = str(p)
            log.debug("NET_TEXTFSM definido para: %s (index True)", p)
            return str(p)
        else:
            log.debug("ntc_templates presente, mas templates/index não encontrados em: %s", p)
            return None
    except Exception as e:
        log.debug("ntc_templates não disponível: %s", e)
        return None


# =============================================================================
# Limpeza de output antes de alimentar o TextFSM (CliTable)
# =============================================================================

_PROMPT_LINE_RE = re.compile(r"(?m)^\S+[>#]\s*$")
_CMD_ECHO_RE = re.compile(r"(?i)^\s*show\s+\S.*$")

def clean_for_textfsm(raw_output: str) -> str:
    """Remove ruído comum (eco de comando, linhas só com prompt, linhas vazias extra).

    Args:
        raw_output: texto original devolvido pelo equipamento.

    Returns:
        Texto “limpo”, adequado para consumo por TextFSM/Clitable.
    """
    if not raw_output:
        return ""
    s = raw_output.replace("\r\n", "\n").replace("\r", "\n")

    # 1) remover linhas que são apenas o prompt
    s = _PROMPT_LINE_RE.sub("", s)

    # 2) remover eco de comando na 1ª linha (ex: "show vlan brief")
    lines = s.split("\n")
    if lines and _CMD_ECHO_RE.match(lines[0] or ""):
        lines = lines[1:]

    # 3) aparar linhas vazias iniciais múltiplas
    while lines and not (lines[0] or "").strip():
        lines.pop(0)

    # 4) e também finais
    while lines and not (lines[-1] or "").strip():
        lines.pop()

    return "\n".join(lines)


# =============================================================================
# CliTable (TextFSM “clássico”)
# =============================================================================

def parse_with_textfsm(command: str, raw_output: str, templates_dir: str | None = None) -> Tuple[List[str], List[List[str]]]:
    """Tenta parsear com TextFSM via CliTable (ntc-templates).

    Importante: esta função NÃO abre sessão SSH. Apenas transforma `raw_output`.

    Args:
        command: comando exato (ex.: "show vlan brief").
        raw_output: saída bruta desse comando.
        templates_dir: diretório dos templates; se None, usa NET_TEXTFSM.

    Returns:
        (headers, rows) se casar com um template; caso contrário ([], []).

    Obs.:
        - Limpa eco/prompt antes de parse (clean_for_textfsm).
        - Escreve logs DEBUG com o template usado ou com a razão do “miss”.
    """
    try:
        from textfsm import clitable  # type: ignore
    except Exception as e:
        log.debug("TextFSM/clitable indisponível: %s", e)
        return ([], [])

    tdir = templates_dir or os.environ.get("NET_TEXTFSM")
    if not tdir:
        log.debug("NET_TEXTFSM não definido — CliTable não tem onde procurar templates.")
        return ([], [])

    s = clean_for_textfsm(raw_output)
    attrs = {"Command": command.strip(), "Vendor": "cisco_ios"}

    try:
        cli = clitable.CliTable("index", tdir)
        cli.ParseCmd(s, attrs)  # pode levantar TextFSMError
        headers = list(cli.header)
        rows = [list(r) for r in cli]
        # Tentativa de log do template usado (depende da versão)
        tpl = getattr(cli, "template", None)
        tpl_name = getattr(tpl, "name", None) or getattr(cli, "TableName", None)
        log.debug("CliTable OK: cmd=%r template=%r -> %d linhas", command, tpl_name, len(rows))
        return (headers, rows)
    except Exception as e:
        # Inclui “State Error raised. Rule Line: X. Input Line: …”
        log.debug("TextFSM miss para %r: %s", command, e)
        return ([], [])


# =============================================================================
# Converter list[dict] (Netmiko use_textfsm=True) -> (headers, rows)
# =============================================================================

def _dictlist_to_table(items: Any) -> Tuple[List[str], List[List[str]]]:
    """Converte a saída típica do Netmiko com `use_textfsm=True` (list[dict]) para (headers, rows).

    Regras:
      - Preserva a ordem das chaves pelo 1º dicionário.
      - Chaves novas que surjam noutros dicionários são acrescentadas ao fim.
      - Converte valores para str; substitui None por "".

    Args:
        items: list[dict] ou outra coisa.

    Returns:
        (headers, rows) — ([],[]) se não for lista de dicts ou se estiver vazia.
    """
    if not isinstance(items, list) or not items:
        return ([], [])

    # Validar que contém dicts
    all_dicts = all(isinstance(x, dict) for x in items)
    if not all_dicts:
        return ([], [])

    # Ordem base: chaves do primeiro item
    header_order: List[str] = list(items[0].keys())

    # Acrescentar quaisquer chaves “novas” que surjam depois
    seen = set(header_order)
    for d in items[1:]:
        for k in d.keys():
            if k not in seen:
                header_order.append(k)
                seen.add(k)

    # Construir linhas, garantindo todas as colunas
    rows: List[List[str]] = []
    for d in items:
        row = []
        for k in header_order:
            v = d.get(k, "")
            if v is None:
                v = ""
            row.append(str(v))
        rows.append(row)

    return (header_order, rows)


# =============================================================================
# Utilitário simples para logar o estado do ambiente TextFSM (debug)
# =============================================================================

def log_textfsm_env():
    """Escreve logs úteis (DEBUG) sobre o estado do ambiente TextFSM."""
    path = os.environ.get("NET_TEXTFSM")
    if not path:
        log.debug("NET_TEXTFSM: (não definido)")
        return
    has_index = os.path.exists(os.path.join(path, "index"))
    log.debug("NET_TEXTFSM: %s (index=%s)", path, has_index)
