import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from desktop_app import (  # noqa: E402
    clean_fc_cn_lines_from_text,
    clean_id_lines_from_text,
    clean_candidate_line,
    convert_lines,
    extract_name_id_lines_from_tables,
    fc_cn_to_hex,
    hex_to_fc_cn,
    import_structured_file,
    import_result_queue_lines,
    import_unconvert_file,
    unconvert_lines,
    valid_hex8,
)


def main() -> None:
    assert hex_to_fc_cn("88984717") == (34968, 18199)
    assert fc_cn_to_hex(34968, 18199) == "88984717"
    assert valid_hex8("88984717")
    assert not valid_hex8("8898471Z")
    assert clean_candidate_line("Active Christopher Benson, 88984765")["extracted"] == "88984765"
    clipboard = clean_id_lines_from_text("Candidate Name\tColleague #\nChris Test\t88984765.0\nJordan Test\t8898-4130")
    assert clipboard["lines"] == ["88984765", "88984130"]
    assert clean_fc_cn_lines_from_text("FC 34968 CN 18199\nbad line")["lines"] == ["34968,18199"]
    assert import_result_queue_lines({"lines": ["Chris Test, 88984765"], "found_rows": 1}) == ["88984765"]
    results, invalid = convert_lines("88984717\nBAD-LINE\n8898-4765", "TEST")
    assert len(results) == 2
    assert len(invalid) == 1
    unconverted, unconvert_invalid = unconvert_lines("34968,18199\nBAD-LINE", "TEST")
    assert len(unconverted) == 1
    assert len(unconvert_invalid) == 1
    assert unconverted[0].hex_value == "88984717"
    imported = extract_name_id_lines_from_tables(
        [[["Candidate Name", "Colleague #"], ["Christopher Benson", "88984765"], ["Supervisor", "88123456"]]],
        "sample.xlsx",
    )
    assert imported["lines"] == ["Christopher Benson, 88984765"]
    sample_text = ROOT / "tests" / "_sample_import.txt"
    sample_pairs = ROOT / "tests" / "_sample_pairs.csv"
    try:
        sample_text.write_text("Worker A, 88984765\nNo usable value\nID 8898-4130\n", encoding="utf-8")
        assert import_structured_file(sample_text)["lines"] == ["88984765", "88984130"]
        sample_pairs.write_text("Facility,Card\n34968,18199\nbad,line\n34968,18277\n", encoding="utf-8")
        assert import_unconvert_file(sample_pairs)["lines"] == ["34968,18199", "34968,18277"]
    finally:
        sample_text.unlink(missing_ok=True)
        sample_pairs.unlink(missing_ok=True)
    print("Desktop app smoke checks passed.")


if __name__ == "__main__":
    main()
