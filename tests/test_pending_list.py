import tempfile
import unittest
from pathlib import Path

import openpyxl
from kks_gen.pending_list import write_pending_list
from tools.gen_step1 import resolve_columns


class PendingListTest(unittest.TestCase):
    def test_pending_equipment_only_and_preserves_source(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "delivered.xlsx"
            target = Path(directory) / "pending.xlsx"
            book = openpyxl.Workbook()
            sheet = book.active
            sheet.append(["原文件序号", "原KKS编码", "设备名称", "设备类型", "修编后KKS编码", "修编说明"])
            sheet.append([2, "", "压力变送器", "设备", "", "编号待核"])
            sheet.append([3, "", "给水流量", "测量设备", "", "测点留空"])
            sheet.append([4, "", "泵启动指令", "指令", "", "指令留空"])
            sheet.append([5, "=1+1", "电动阀", "设备", "", "重复编码清空"])
            sheet.append([None, "", "新增父级", "系统", "", ""])
            sheet.append([6, "", "已编码设备", "设备", "61LAA01AA001", ""])
            book.save(source)
            book.close()
            before = source.read_bytes()
            result = write_pending_list(source, target, resolve_columns)
            self.assertEqual(2, result["rows"])
            self.assertEqual(before, source.read_bytes())
            output = openpyxl.load_workbook(target)
            self.assertEqual([2, 5], [row[1].value for row in list(output.active.rows)[1:]])
            self.assertEqual("重复编码待核", output.active["G3"].value)
            self.assertEqual("s", output.active["D3"].data_type)
            output.close()
