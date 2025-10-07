# clean_switch/cli.py
from __future__ import annotations
import argparse, logging, os
from .raw_outputs import save_raw_outputs
from .logging_config import setup_logging
from .network_connector import collect, COMMANDS_DEFAULT
from .excel_pipeline import create_excel, metrics_from_collected
from history_json import save_snapshot
from ideia6_features import ideia6_run_pipeline

# NOVO: modo crawl
from .topology_crawl import crawl_topology

def main():
    parser = argparse.ArgumentParser(description="Clean/refactored switch collector → Excel")
    parser.add_argument("--host", required=True, help="IP do switch semente")
    parser.add_argument("--user", required=True)
    parser.add_argument("--password", required=True)
    parser.add_argument("--port", type=int, default=22)
    parser.add_argument("--no-excel", action="store_true", help="Don't generate Excel file")
    parser.add_argument("--commands", nargs="*", default=COMMANDS_DEFAULT)
    parser.add_argument("--out", default=None, help="Output directory (defaults to ./outputs/<hostname>)")
    parser.add_argument("--log", default="INFO", help="Logging level (DEBUG, INFO, WARNING)")
    parser.add_argument("--raw-only", action="store_true",
                    help="Gerar apenas os .txt (raw outputs) e NÃO gerar Excel.")
    parser.add_argument("--raw-dir", default="outputs",
                    help="Diretório base para guardar os .txt (default: outputs).")
    parser.add_argument("--raw-ts-subdir", action="store_true",
                    help="Se definido, cria subpasta por timestamp (ts) para os .txt.")


    # ---- FLAGS NOVAS PARA CRAWL ----
    parser.add_argument("--crawl-depth", type=int, default=0,
                        help="0=apenas o semente; 1=vizinhos; 2=vizinhos dos vizinhos...")
    parser.add_argument("--allowed-subnet", action="append",
                        help="CIDR permitido (pode repetir, e.g., --allowed-subnet 192.168.99.0/24)")
    parser.add_argument("--no-dns-fallback", action="store_true",
                        help="Não tentar resolver Device ID via DNS")
    parser.add_argument("--hostmap", default=None,
                        help="YAML com mapeamento DeviceID→IP para vizinhos (fallback)")

    args = parser.parse_args()

    level = getattr(logging, (args.log or "INFO").upper(), logging.INFO)
    log = setup_logging(level)

    # --- MODO CRAWL ---
    if args.crawl_depth and args.crawl_depth > 0:
        # hostmap opcional (carregar apenas se indicado, para não obrigar a ter PyYAML instalado)
        hostmap = None
        if args.hostmap:
            try:
                import yaml  # type: ignore
            except Exception:
                log.error("Para usar --hostmap precisa de PyYAML (pip install pyyaml). A prosseguir sem hostmap.")
                hostmap = None
            else:
                try:
                    with open(args.hostmap, "r", encoding="utf-8") as f:
                        hostmap = yaml.safe_load(f) or {}
                except Exception as e:
                    log.error("Falha a ler hostmap %s: %s", args.hostmap, e)
                    hostmap = None

        outdir_override = args.out  # opcional: força todos os Excels para o mesmo diretório
        if outdir_override:
            os.makedirs(outdir_override, exist_ok=True)

        results = crawl_topology(
            seed_ip=args.host,
            username=args.user,
            password=args.password,
            enable=None,
            port=args.port,
            max_depth=args.crawl_depth,
            allowed_subnets=args.allowed_subnet,
            dns_fallback=not args.no_dns_fallback,
            hostmap=hostmap,
            out_dir_override=outdir_override,
        )
        if results:
            print("OK (crawl):")
            for host, xlsx in results:
                print(f"  - {host}: {xlsx}")
        else:
            print("OK (crawl): nenhum Excel gerado.")
        return

    # --- FLUXO “SIMPLES” (apenas o semente, como antes) ---
    host, ts, collected, outdir = collect(args.host, args.user, args.password, port=args.port, commands=args.commands)
    if args.out:
        outdir = args.out
    os.makedirs(outdir, exist_ok=True)
    xlsx_path = os.path.join(outdir, f"{host}_levantamento.xlsx")

    metrics = metrics_from_collected(collected)
    save_snapshot(base_dir=outdir, hostname=host, ts=ts, collected=collected, meta={"host_ip": args.host, "metrics": metrics}, max_keep=10)

    # --- RAW outputs (.txt)
    out_txt_dir = save_raw_outputs(
        hostname=host,
        collected=collected,
        base_dir=args.raw_dir,
        make_timestamp_subdir=args.raw_ts_subdir,
        ts=ts,
    )

    # Se pedir apenas RAW, termina já aqui (não gera Excel)
    if args.raw_only:
        print(f"OK (raw-only): .txt em {out_txt_dir}")
        return

    

    if not args.no_excel:
        
        create_excel(host, ts, collected, xlsx_path)
        try:
            ideia6_run_pipeline(host, ts, collected, xlsx_path, host_ip=args.host)
        except Exception as e:
            log.error("Ideia6 pipeline failed: %s", e)
        # --- limpeza: manter só a Execução mais recente ---
        from clean_switch.excel_utils import _list_execucao_sheets
        from openpyxl import load_workbook

        try:
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
            log.error("Falhou limpeza de sheets Execução: %s", e)

    print("OK:", xlsx_path if not args.no_excel else f"snapshots in {outdir}")

if __name__ == "__main__":
    main()
