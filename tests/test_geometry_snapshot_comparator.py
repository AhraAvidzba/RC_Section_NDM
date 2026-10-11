"""Provenance renaming must not hide geometric or semantic changes."""

import importlib.util
from pathlib import Path
import unittest


SCRIPT = Path(__file__).resolve().parents[1] / "tools/build_workbook/Compare-GeometrySnapshots.py"
SPEC = importlib.util.spec_from_file_location("geometry_comparator", SCRIPT)
COMPARATOR = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(COMPARATOR)


def records():
    fields = ["d:0000000000000000"] * len(COMPARATOR.FIELDS["BAR"])
    fields[COMPARATOR.FIELDS["BAR"].index("source")] = "s:Old1"
    first = ["Fixture", "BAR", "R1", *fields]
    second = first.copy()
    second[2] = "R2"
    second[3 + COMPARATOR.FIELDS["BAR"].index("source")] = "s:Old2"
    layout = ["Fixture", "LAYOUT", "Old1", *(["d:0000000000000000"] * len(COMPARATOR.FIELDS["LAYOUT"]))]
    return [first, second, layout]


class GeneratedNameMappingTests(unittest.TestCase):
    def renamed(self):
        actual = records()
        source = 3 + COMPARATOR.FIELDS["BAR"].index("source")
        actual[0][source] = "s:New1"
        actual[1][source] = "s:New2"
        actual[2][2] = "New1"
        return actual

    def test_mapping_is_explicit_and_lossless(self):
        actual = self.renamed()
        mapped, proof = COMPARATOR.map_generated_rebar_names(records(), actual)
        self.assertEqual(COMPARATOR.compare(records(), mapped), [])
        self.assertEqual(len(proof), 2)
        self.assertEqual(actual[2][2], "New1")

    def test_coordinate_mutation_is_still_detected(self):
        actual = self.renamed()
        actual[0][3] = "d:3FF0000000000000"
        mapped, _ = COMPARATOR.map_generated_rebar_names(records(), actual)
        errors = COMPARATOR.compare(records(), mapped)
        self.assertEqual(len(errors), 1)
        self.assertEqual(errors[0]["field"], "x")

    def test_ambiguous_names_are_rejected(self):
        actual = self.renamed()
        source = 3 + COMPARATOR.FIELDS["BAR"].index("source")
        actual[1][source] = "s:New1"
        with self.assertRaises(ValueError):
            COMPARATOR.map_generated_rebar_names(records(), actual)

    def test_physical_id_is_not_aliased(self):
        actual = self.renamed()
        actual[0][2] = "OtherR1"
        with self.assertRaises(ValueError):
            COMPARATOR.map_generated_rebar_names(records(), actual)

    def test_layout_requires_a_physical_bar(self):
        actual = self.renamed()
        actual[2][2] = "Unknown"
        with self.assertRaises(ValueError):
            COMPARATOR.map_generated_rebar_names(records(), actual)


if __name__ == "__main__":
    unittest.main()
