#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""amount_unifier 单元测试（纯函数，不依赖任何真实/业务数据）。

运行：python -m unittest discover -s tests -v
"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from amount_unifier import (  # noqa: E402
    AmountNormalizer,
    UNIT_SCALE,
    normalize_amount,
    unit_rules_for,
)


class TestNormalizeAmount(unittest.TestCase):
    def test_text_yuan_to_wan(self):
        self.assertAlmostEqual(normalize_amount("2004279.6元", "s").value, 200.428, places=3)

    def test_text_wanyuan(self):
        self.assertEqual(normalize_amount("6581万元", "s").value, 6581)

    def test_text_wan(self):
        self.assertEqual(normalize_amount("777.3万", "s").value, 777.3)

    def test_text_yiyuan_to_wan(self):
        self.assertEqual(normalize_amount("2亿元", "s", target_unit="万元").value, 20000)

    def test_text_qianyuan_to_yuan(self):
        self.assertEqual(normalize_amount("5千元", "s", target_unit="元").value, 5000)

    def test_target_unit_yuan(self):
        self.assertEqual(normalize_amount("6581万元", "s", target_unit="元").value, 65810000)

    def test_target_unit_yi(self):
        self.assertEqual(normalize_amount("500万元", "s", target_unit="亿元").value, 0.05)

    def test_numeric_conservative(self):
        # 纯数值不做单位猜测，即便很大也不 ÷
        self.assertEqual(normalize_amount(5000000, "n").value, 5000000)

    def test_formula_asis(self):
        # 公式如实保留计算值，不 ÷
        self.assertEqual(normalize_amount("=a+b", "f", 13400000).value, 13400000)

    def test_formula_no_cache_manual(self):
        r = normalize_amount("=a+b", "f", None)
        self.assertTrue(r.is_manual)

    def test_range_manual(self):
        r = normalize_amount("1万-100万不等", "s")
        self.assertIsNone(r.value)
        self.assertTrue(r.is_manual)

    def test_empty_manual(self):
        self.assertTrue(normalize_amount(None, "n").is_manual)

    def test_longtext_manual(self):
        self.assertTrue(normalize_amount("累计赔偿限额200万，每次事故100万，附加若干责任", "s").is_manual)

    def test_unknown_type_manual(self):
        self.assertTrue(normalize_amount(True, "b").is_manual)


class TestUnitRules(unittest.TestCase):
    def test_default_wan_conversions(self):
        conv = {r.label: r.convert for r in unit_rules_for("万元")}
        self.assertAlmostEqual(conv["文本元"](10000), 1.0)
        self.assertEqual(conv["文本万元"](5), 5)
        self.assertEqual(conv["文本亿元"](2), 20000)
        self.assertAlmostEqual(conv["文本千元"](5), 0.5)

    def test_yuan_target_conversions(self):
        conv = {r.label: r.convert for r in unit_rules_for("元")}
        self.assertEqual(conv["文本元"](123), 123)
        self.assertEqual(conv["文本万元"](1), 10000)

    def test_bad_target_unit(self):
        with self.assertRaises(ValueError):
            unit_rules_for("海里")

    def test_custom_rules(self):
        import re
        rules = (
            __import__("amount_unifier").TextUnitRule(
                re.compile(r"^\s*([\d,\.]+)\s*桶\s*$"), "文本桶", lambda n: n * 2),
        )
        norm = AmountNormalizer(rules)
        self.assertEqual(norm.normalize("3桶", "s").value, 6)


if __name__ == "__main__":
    unittest.main()
