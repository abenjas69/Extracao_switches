
from clean_switch.parsers import parser_show_spanning_tree_from_text, parse_etherchannel_summary_from_text

def test_stp_parser_empty():
    assert parser_show_spanning_tree_from_text("") == ([], [])

def test_eth_parser_empty():
    assert parse_etherchannel_summary_from_text("") == ([], [])
