
# clean_switch (refactor)

Refatoração modular do projeto:

- `clean_switch/network_connector.py`: ligação SSH (Netmiko) + recolha/parse.
- `clean_switch/textfsm_utils.py`: TextFSM helpers (NET_TEXTFSM auto, dictlist→tabela).
- `clean_switch/parsers.py`: parsers internos (STP, EtherChannel).
- `clean_switch/excel_utils.py`: helpers Excel (naming, autosize, CF, tabelas, gráficos).
- `clean_switch/excel_dashboard.py`: construção da Dashboard.
- `clean_switch/excel_pipeline.py`: `create_excel()` e `metrics_from_collected()`.
- `clean_switch/cli.py`: CLI simples que integra history_json + ideia6_features.

Compatibilidade: reusa `history_json.py` e `ideia6_features.py` originais sem alterações.

## Como correr
Criar Excel e recolha(.txt) + snapshots JSON
```bash
python -m clean_switch.cli --host 192.168.99.2 --user diogo --password '***'
```

Ou só recolha(.txt) + snapshots JSON:
```bash
python -m clean_switch.cli --host 192.168.99.2 --user diogo --password '***' --no-excel
```

Criar excel e recolha(txt) + snapshots JSON e modo DEBUG para logs no Terminal:
```bash
python -m clean_switch.cli --host 192.168.99.2 --user diogo --password '***' --log DEBUG
```

python -m clean_switch.cli --host 192.168.99.2 --user diogo --password "P@ssword123" --crawl-depth 0 --allowed-subnet 192.168.99.0/24