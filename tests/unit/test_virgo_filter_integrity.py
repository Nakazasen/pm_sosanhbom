import pytest
from src.core.bom_filter_manager import BOMFilterManager
from src.core.default_rules import DEFAULT_MODEL_RULES

def test_virgo_filter_rules_count_and_content():
    mgr = BOMFilterManager()
    virgo_rules = mgr.get_rules_for_model("Virgo", active_only=True)
    
    assert len(virgo_rules) == 62, f"Expected 62 rules for Virgo, got {len(virgo_rules)}"
    
    expected_defaults = DEFAULT_MODEL_RULES["Virgo"]
    assert len(expected_defaults) == 62
    
    for idx, (exp, act) in enumerate(zip(expected_defaults, virgo_rules), start=1):
        assert act.item_name == exp["item_name"], f"Rule #{idx} item_name mismatch: {act.item_name} != {exp['item_name']}"
        assert act.match_mode == exp["match_mode"], f"Rule #{idx} match_mode mismatch: {act.match_mode} != {exp['match_mode']}"
        assert act.part_code == exp["part_code"], f"Rule #{idx} part_code mismatch: {act.part_code} != {exp['part_code']}"

def test_virgo_model_rules_for_pruner():
    mgr = BOMFilterManager()
    pruner_rules = mgr.get_model_rules_for_pruner("Virgo")
    assert len(pruner_rules) == 62
    assert any(r.item_name == "BOTTLE WASTE" for r in pruner_rules)
    assert any(r.item_name == "FILL UP CONTAINER ASSY" and r.match_mode == "Part_name" for r in pruner_rules)
    assert any(r.item_name == "PWB DRUM ASSY" for r in pruner_rules)
    assert any(r.item_name == "PWB ID SENSOR ASSY" for r in pruner_rules)
    assert any(r.item_name == "PWB MAIN ASSY" and r.match_mode == "Part_name" for r in pruner_rules)
    # Check newly added Part_name rules (rows 51-62)
    assert any(r.item_name == "APC ASSY B" and r.match_mode == "Part_name" for r in pruner_rules)
    assert any(r.item_name == "CASE DLP" and r.match_mode == "Part_name" for r in pruner_rules)
