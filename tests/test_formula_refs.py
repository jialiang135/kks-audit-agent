import unittest
from kks_gen.formula_refs import remap_formula


class FormulaReferenceTests(unittest.TestCase):
    def test_row_map_is_not_a_constant_offset(self):
        self.assertEqual('=SUM($G$6:G11)+C6&F11',
                         remap_formula('=SUM($F$2:F3)+B2&E3', {2: 6, 3: 11}, 1, '主表', '主表'))

    def test_preserve_literals_names_and_other_sheets(self):
        self.assertEqual('="F2"&NamedRange&Other!F2&G6',
                         remap_formula('="F2"&NamedRange&Other!F2&F2', {2: 6}, 1, '主表', '主表'))

    def test_other_sheet_only_remaps_explicit_reference_to_moved_sheet(self):
        self.assertEqual("='主表'!G6+F2", remap_formula("='主表'!F2+F2", {2: 6}, 1, '主表', '其他'))

    def test_entire_rows_columns_and_absolute_references(self):
        self.assertEqual('=SUM($G:$H)+SUM($6:$11)',
                         remap_formula('=SUM($F:$G)+SUM($2:$3)', {2: 6, 3: 11}, 1, '主表', '主表'))
