#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
amount_unifier.py — 金额单位统一子任务
=======================================
将收集表格中格式不统一的金额列，统一换算为标准的"万元"数值。

定位
----
本脚本是"表格清稿工具"的一个子任务模块，可两种方式使用：
  1) 单独运行（命令行）：
        python amount_unifier.py -i input.xlsx -c H I
  2) 作为模块调用：
        from amount_unifier import unify_columns, normalize_amount
        report = unify_columns("input.xlsx", ["H", "I"], mode="overwrite")
        normalize_amount("2004279.6元", "s").value      # -> 200.428

设计分层
--------
  规则层  AmountNormalizer / TextUnitRule   纯逻辑，不依赖 Excel，可单测、可扩展单位
  编排层  AmountColumnCleaner               只负责 Excel 读写、列定位、合并/上色/两种落地模式
  报告层  CleanReport                       结构化结果（统计、需人工清单）
  入口层  main(argv) / unify_columns(...)   CLI 与库 API
"""

from __future__ import annotations

import argparse
import copy
import json
import logging
import os
import re
import sys
from collections import OrderedDict
from dataclasses import dataclass, field
from datetime import datetime
from typing import Callable, List, Optional, Sequence, Tuple

try:
    import openpyxl
    from openpyxl.cell.cell import MergedCell
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import column_index_from_string, get_column_letter
except ImportError:  # pragma: no cover
    sys.stderr.write("错误：需要 openpyxl，请先安装：pip install openpyxl\n")
    raise

log = logging.getLogger("amount_unifier")


# ============================================================
# 可调常量
# ============================================================
ROUND_DIGITS = 4                 # 标准化值保留小数位
NUMERIC_FORMAT = "0.00"           # 结果单元格数字格式（数值，保留两位小数）
LEGEND_SHEET = "颜色图例"
RULE_PREFIX_MANUAL = "需人工"     # 需人工类规则名前缀
DEFAULT_TARGET_UNIT = "万元"      # 默认目标单位

# 金额单位体系：各单位相对“元”的比例（用于任意目标单位换算）
UNIT_SCALE = {
    "元": 1.0,
    "千元": 1e3,
    "万元": 1e4,
    "亿元": 1e8,
}


# ============================================================
# 配色（底色标注处理类型）
# ============================================================
FILL_MERGE   = PatternFill("solid", start_color="E4DFEC", end_color="E4DFEC")  # 浅紫（预留；合并格现按锚点类型上色）
FILL_NUMERIC = PatternFill("solid", start_color="E2EFDA", end_color="E2EFDA")  # 浅绿-数值
FILL_TEXTWAN = PatternFill("solid", start_color="DDEBF7", end_color="DDEBF7")  # 浅蓝-万元/万
FILL_TEXYUAN = PatternFill("solid", start_color="FCE4D6", end_color="FCE4D6")  # 浅橙-元
FILL_FORMULA = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")  # 浅黄-公式
FILL_MANUAL  = PatternFill("solid", start_color="FFC7CE", end_color="FFC7CE")  # 浅红-需人工

# 规则名 -> 底色
_RULE_FILL = {
    "数值": FILL_NUMERIC,
    "文本亿元": FILL_TEXTWAN,
    "文本万元": FILL_TEXTWAN,
    "文本千元": FILL_TEXTWAN,
    "文本万": FILL_TEXTWAN,
    "文本元": FILL_TEXYUAN,
    "公式": FILL_FORMULA,
}

# 图例（大类 -> 展示）
# 图例只描述“颜色↔源类型”（与目标单位无关）；换算口径由 _write_legend 顶行标注
LEGEND_ITEMS = [
    (FILL_NUMERIC, "数值",                 "原单元格为数值"),
    (FILL_TEXTWAN, "文本万/万元/千元/亿元", "原文带中文数量单位，按单位换算"),
    (FILL_TEXYUAN, "文本元",               "原文带'元'单位，按单位换算"),
    (FILL_FORMULA, "公式",                 "公式计算值，如实保留"),
    (FILL_MANUAL,  "需人工",               "区间/长文本等无法自动处理，保留原值"),
]


def fill_for_rule(rule: str) -> PatternFill:
    """按规则名返回底色；需人工类统一浅红"""
    if rule.startswith(RULE_PREFIX_MANUAL):
        return FILL_MANUAL
    return _RULE_FILL.get(rule, FILL_MANUAL)


# ============================================================
# 规则层：纯逻辑（不依赖 openpyxl 工作表对象）
# ============================================================
RE_YUAN     = re.compile(r"^\s*([\d,]+(?:\.\d+)?)\s*元\s*$")
RE_QIANYUAN = re.compile(r"^\s*([\d,]+(?:\.\d+)?)\s*千元\s*$")
RE_WANYUAN  = re.compile(r"^\s*([\d,]+(?:\.\d+)?)\s*万元\s*$")
RE_WAN      = re.compile(r"^\s*([\d,]+(?:\.\d+)?)\s*万\s*$")
RE_YIYUAN   = re.compile(r"^\s*([\d,]+(?:\.\d+)?)\s*亿元\s*$")


@dataclass(frozen=True)
class TextUnitRule:
    """一条文本单位规则：匹配模式 + 规则名 + 换算函数"""
    pattern: "re.Pattern"
    label: str
    convert: Callable[[float], float]


# 源单位识别表：(正则, 规则名label, UNIT_SCALE键)
# 注：识别 元/千元/万元/亿元；“圆”等同义词可按需再加入
_SRC_UNITS = (
    (RE_YIYUAN,   "文本亿元", "亿元"),
    (RE_WANYUAN,  "文本万元", "万元"),
    (RE_QIANYUAN, "文本千元", "千元"),
    (RE_WAN,      "文本万",   "万元"),
    (RE_YUAN,     "文本元",   "元"),
)


def _make_convert(src_key: str, target_scale: float):
    """生成换算函数：数值 × 源单位比例 ÷ 目标单位比例"""
    src_scale = UNIT_SCALE[src_key]
    return lambda n: n * src_scale / target_scale


def unit_rules_for(target_unit: str = DEFAULT_TARGET_UNIT) -> Tuple[TextUnitRule, ...]:
    """按目标单位生成文本单位规则表（target_unit='万元' 时与历史行为完全一致）。"""
    if target_unit not in UNIT_SCALE:
        raise ValueError(f"未知目标单位：{target_unit}，可选 {list(UNIT_SCALE)}")
    tscale = UNIT_SCALE[target_unit]
    return tuple(TextUnitRule(pat, label, _make_convert(key, tscale))
                 for pat, label, key in _SRC_UNITS)


# 默认文本单位规则表（目标单位=万元；保留导出名以兼容既有引用）
DEFAULT_TEXT_UNIT_RULES: Tuple[TextUnitRule, ...] = unit_rules_for(DEFAULT_TARGET_UNIT)


@dataclass(frozen=True)
class NormResult:
    """单值标准化结果"""
    value: Optional[float]   # 万元数值；None 表示需人工
    rule: str                # 规则名
    note: str = ""           # 说明（原始内容/换算过程）

    @property
    def is_manual(self) -> bool:
        return self.value is None or self.rule.startswith(RULE_PREFIX_MANUAL)


def _to_number(text) -> Optional[float]:
    """提取数值，自动去千分位逗号"""
    if text is None:
        return None
    s = str(text).replace(",", "").strip()
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


class AmountNormalizer:
    """把单元格 (值, 类型, 公式缓存值) 标准化为万元数值。不依赖工作表对象，可单测。"""

    def __init__(self, text_unit_rules: Optional[Sequence[TextUnitRule]] = None,
                 target_unit: str = DEFAULT_TARGET_UNIT):
        self.target_unit = target_unit
        self.text_unit_rules = (tuple(text_unit_rules) if text_unit_rules is not None
                                else unit_rules_for(target_unit))

    def normalize(self, value, data_type: str = "s", cached=None) -> NormResult:
        # 空值
        if value is None:
            return NormResult(None, f"{RULE_PREFIX_MANUAL}-空值", "单元格为空")

        # 公式：如实采用 data_only=True 读到的缓存计算值，不做单位换算猜测
        if data_type == "f":
            if cached is None:
                return NormResult(None, f"{RULE_PREFIX_MANUAL}-公式无缓存", "公式无计算缓存")
            num = _to_number(cached)
            if num is None:
                return NormResult(None, f"{RULE_PREFIX_MANUAL}-公式异常", f"缓存值无法解析：{cached}")
            return NormResult(round(num, ROUND_DIGITS), "公式", f"公式计算值 {num}（如实保留）")

        # 纯数值：保守保留为万元，不做单位猜测（不因数值大而 ÷10000）
        if data_type == "n":
            num = float(value)
            return NormResult(round(num, ROUND_DIGITS), "数值", f"原值 {num}（按万元保留）")

        # 文本：按单位规则表依次匹配
        if data_type == "s":
            text = str(value).strip()
            for rule in self.text_unit_rules:
                m = rule.pattern.match(text)
                if m:
                    num = _to_number(m.group(1))
                    if num is not None:
                        return NormResult(round(rule.convert(num), ROUND_DIGITS), rule.label, text)
            # 未命中单位规则
            if "不等" in text or "-" in text or "至" in text:
                return NormResult(None, f"{RULE_PREFIX_MANUAL}-区间文本", text[:50])
            if len(text) > 20:
                return NormResult(None, f"{RULE_PREFIX_MANUAL}-长文本", text[:50])
            return NormResult(None, f"{RULE_PREFIX_MANUAL}-无法解析", text[:50])

        # 其他类型（日期/布尔等）
        return NormResult(None, f"{RULE_PREFIX_MANUAL}-未知类型",
                          f"data_type={data_type}, value={value}")


# 模块级默认实例，便于低层单值复用
_DEFAULT_NORMALIZER = AmountNormalizer()


def normalize_amount(value, data_type: str = "s", cached=None,
                     text_unit_rules=None, target_unit=DEFAULT_TARGET_UNIT) -> NormResult:
    """低层单值标准化（纯函数）。
    例：normalize_amount('2004279.6元', 's').value  -> 200.428
    """
    if text_unit_rules is not None:
        normalizer = AmountNormalizer(text_unit_rules, target_unit)
    elif target_unit == DEFAULT_TARGET_UNIT:
        normalizer = _DEFAULT_NORMALIZER
    else:
        normalizer = AmountNormalizer(target_unit=target_unit)
    return normalizer.normalize(value, data_type, cached)


# ============================================================
# 报告层
# ============================================================
@dataclass
class CleanReport:
    input_path: str = ""
    output_path: str = ""
    mode: str = "append"
    target_unit: str = DEFAULT_TARGET_UNIT
    columns: List[str] = field(default_factory=list)
    total: int = 0
    success: int = 0
    by_rule: "OrderedDict[str, int]" = field(default_factory=OrderedDict)
    manual: List[dict] = field(default_factory=list)   # 需人工记录

    def add_result(self, rec: dict):
        self.total += 1
        self.by_rule[rec["rule"]] = self.by_rule.get(rec["rule"], 0) + 1
        if rec["value"] is None:
            self.manual.append({"coord": rec["coord"], "header": rec["header"],
                                "rule": rec["rule"], "note": rec["note"]})
        else:
            self.success += 1

    def to_dict(self) -> dict:
        """结构化报告，供 --json 输出或上层程序使用。"""
        return {
            "input": self.input_path,
            "output": self.output_path,
            "mode": self.mode,
            "target_unit": self.target_unit,
            "columns": list(self.columns),
            "total": self.total,
            "success": self.success,
            "manual_count": len(self.manual),
            "by_rule": dict(self.by_rule),
            "manual": list(self.manual),
        }


# ============================================================
# 编排层：Excel 读写
# ============================================================
class AmountColumnCleaner:
    """负责 Excel 读取、列定位、落地（append/overwrite）、合并与上色。"""

    def __init__(self, input_path, sheet=None, header_row=None,
                 row_start=None, row_end=None, mode="append", output_path=None,
                 text_unit_rules=None, target_unit=DEFAULT_TARGET_UNIT,
                 colorize=True, legend=True, logger=None):
        self.input_path = input_path
        self.mode = mode
        self.colorize = colorize
        self.legend = legend
        self.output_path = output_path or self._gen_output_path(input_path, mode)
        self.log = logger or log
        self.normalizer = AmountNormalizer(text_unit_rules, target_unit)

        # 双模式加载：False 读公式原文/类型；True 读公式缓存计算值
        self.wb_formula = openpyxl.load_workbook(input_path, data_only=False)
        self.wb_value = openpyxl.load_workbook(input_path, data_only=True)

        self.sheet_name = sheet or self._detect_sheet()
        self.ws_formula = self.wb_formula[self.sheet_name]
        self.ws_value = self.wb_value[self.sheet_name]

        self.header_row = header_row or self._detect_header_row()
        self.row_start = row_start or (self.header_row + 1)
        self.row_end = row_end or self.ws_formula.max_row

        self.merged_map = self._build_merged_map()
        self.report = CleanReport(input_path=input_path, output_path=self.output_path,
                                  mode=mode, target_unit=target_unit)

    # ---------- 初始化辅助 ----------
    @staticmethod
    def _gen_output_path(input_path, mode="append"):
        """默认输出名：{原文件名}_clean_{mode}_{时间戳}.xlsx"""
        base, ext = os.path.splitext(input_path)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        return f"{base}_clean_{mode}_{ts}{ext}"

    def _detect_sheet(self):
        """未指定 sheet 时，取第一张工作表"""
        return self.wb_formula.sheetnames[0]

    def _detect_header_row(self, scan=10):
        """在前 scan 行里找第一个'非空文本单元格>=3'的行作为表头行"""
        ws = self.ws_formula
        for r in range(1, min(ws.max_row, scan) + 1):
            n = sum(1 for c in range(1, ws.max_column + 1)
                    if ws.cell(r, c).data_type == "s" and ws.cell(r, c).value)
            if n >= 3:
                return r
        return 1

    def _build_merged_map(self):
        """被覆盖格(row,col) -> 锚点(row,col)"""
        mapping = {}
        for rng in self.ws_formula.merged_cells.ranges:
            anchor = (rng.min_row, rng.min_col)
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    if (r, c) != anchor:
                        mapping[(r, c)] = anchor
        return mapping

    # ---------- 列定位 ----------
    def resolve_column(self, identifier):
        """列标识（字母或列名）-> (col_idx, col_letter, header)"""
        identifier = str(identifier).strip()
        if re.match(r"^[A-Za-z]{1,3}$", identifier):
            idx = column_index_from_string(identifier.upper())
            header = self.ws_formula.cell(self.header_row, idx).value
            return idx, identifier.upper(), header
        for idx in range(1, self.ws_formula.max_column + 1):
            header = self.ws_formula.cell(self.header_row, idx).value
            if header and identifier in str(header):
                return idx, get_column_letter(idx), header
        raise ValueError(f"无法解析列标识：'{identifier}'")

    # ---------- 单格标准化（含合并继承） ----------
    def normalize_cell(self, row, col) -> NormResult:
        cell_f = self.ws_formula.cell(row, col)
        # 合并继承：取锚点格的标准化结果
        if (row, col) in self.merged_map:
            ar, ac = self.merged_map[(row, col)]
            acell = self.ws_formula.cell(ar, ac)
            a = self.normalizer.normalize(acell.value, acell.data_type,
                                          self.ws_value.cell(ar, ac).value)
            return NormResult(a.value, "合并继承",
                              f"继承自 {get_column_letter(ac)}{ar}（{a.rule}）")
        return self.normalizer.normalize(cell_f.value, cell_f.data_type,
                                         self.ws_value.cell(row, col).value)

    # ---------- 主流程 ----------
    def run(self, columns) -> CleanReport:
        self.log.info("输入：%s | 工作表：%s | 表头行：%d | 数据行：%d-%d | 模式：%s | 目标单位：%s",
                      self.input_path, self.sheet_name, self.header_row,
                      self.row_start, self.row_end, self.mode, self.report.target_unit)

        results: "OrderedDict[str, list]" = OrderedDict()
        for ident in columns:
            idx, letter, header = self.resolve_column(ident)
            col_res = []
            for row in range(self.row_start, self.row_end + 1):
                cell_f = self.ws_formula.cell(row, idx)
                nr = self.normalize_cell(row, idx)
                rec = {"row": row, "coord": f"{letter}{row}", "col_idx": idx,
                       "col_letter": letter, "header": header,
                       "value": nr.value, "rule": nr.rule, "note": nr.note,
                       "raw_value": cell_f.value}
                col_res.append(rec)
                self.report.add_result(rec)
            results[letter] = col_res
            self.report.columns.append(letter)
            self.log.info("处理列 %s（%s）：共 %d 行", letter, header, len(col_res))

        if self.mode == "overwrite":
            self._write_overwrite(results)
        else:
            self._write_append(results)
        return self.report

    # ---------- 落地公共工具 ----------
    @staticmethod
    def _shift_merges_for_insert(ws, insert_at, amount):
        """insert_cols 不会自动平移合并区，这里手动把列号>=insert_at 的合并区右移 amount 列。"""
        ranges = list(ws.merged_cells.ranges)
        ws.merged_cells.ranges = []
        for rng in ranges:
            mn, mx = rng.min_col, rng.max_col
            if mn >= insert_at:
                mn += amount
                mx += amount
            elif mx >= insert_at:
                mx += amount
            ws.merge_cells(start_row=rng.min_row, end_row=rng.max_row,
                           start_column=mn, end_column=mx)

    @staticmethod
    def _mirror_single_col_merges(ws, src_col, dst_cols):
        """把原列(src_col)上的单列纵向合并区复刻到 dst_cols，使辅助列与原列合并结构一致。"""
        to_mirror = [rng for rng in ws.merged_cells.ranges
                     if rng.min_col == src_col == rng.max_col]
        for rng in to_mirror:
            for dst in dst_cols:
                dl = get_column_letter(dst)
                rng_str = f"{dl}{rng.min_row}:{dl}{rng.max_row}"
                if not any(str(x) == rng_str for x in ws.merged_cells.ranges):
                    ws.merge_cells(start_row=rng.min_row, end_row=rng.max_row,
                                   start_column=dst, end_column=dst)
        return len(to_mirror)

    def _fill_result_cell(self, cell, rec):
        """写标准化值并上色；需人工写回原值。"""
        if rec["rule"].startswith(RULE_PREFIX_MANUAL):
            cell.value = rec["raw_value"]           # 无法处理：保留原值
        else:
            cell.value = rec["value"]
            cell.number_format = NUMERIC_FORMAT
        if self.colorize:
            cell.fill = fill_for_rule(rec["rule"])

    @staticmethod
    def _clone_text_style(src, dst):
        """把 src 的字体/对齐/边框复制到 dst（不含填充色与数字格式）。"""
        dst.font = copy.copy(src.font)
        dst.alignment = copy.copy(src.alignment)
        dst.border = copy.copy(src.border)

    @staticmethod
    def _with_bold(font):
        """在保留原字体属性基础上加粗（用于生成列表头突出显示）。"""
        return Font(name=font.name, size=font.sz, bold=True,
                    italic=font.i, color=font.color)

    def _write_legend(self, wb):
        if not (self.legend and self.colorize):
            return
        if LEGEND_SHEET in wb.sheetnames:
            del wb[LEGEND_SHEET]
        lw = wb.create_sheet(LEGEND_SHEET)
        lw.cell(1, 1, "底色标注图例").font = Font(bold=True, size=12)
        lw.cell(2, 1, f"统一换算为目标单位：{self.normalizer.target_unit}").font = \
            Font(bold=True, color="C00000")
        for c, h in enumerate(["底色示例", "处理类型", "说明"], start=1):
            lw.cell(3, c, h).font = Font(bold=True)
        for i, (fill, name, desc) in enumerate(LEGEND_ITEMS, start=4):
            lw.cell(i, 1, "").fill = fill
            lw.cell(i, 2, name)
            lw.cell(i, 3, desc)
        lw.column_dimensions["A"].width = 12
        lw.column_dimensions["B"].width = 14
        lw.column_dimensions["C"].width = 46
        self.log.info("已生成'%s'工作表", LEGEND_SHEET)

    # ---------- 落地：append（默认） ----------
    def _write_append(self, results):
        """保留原列，在原列后插入'结果列'+'说明列'，复刻原列合并结构。"""
        wb = openpyxl.load_workbook(self.input_path, data_only=False)
        ws = wb[self.sheet_name]

        # 从右往左处理，避免多列插入时索引互相影响
        items = sorted(results.items(), key=lambda kv: kv[1][0]["col_idx"], reverse=True)
        for col_letter, col_res in items:
            col_idx = col_res[0]["col_idx"]
            header_name = col_res[0]["header"] or col_letter
            insert_at = col_idx + 1

            ws.insert_cols(insert_at, 2)
            self._shift_merges_for_insert(ws, insert_at, 2)
            std_col, status_col = insert_at, insert_at + 1

            src_h = ws.cell(self.header_row, col_idx)
            for _c, _txt in ((std_col, f"{header_name}_标准化"),
                             (status_col, f"{header_name}_处理状态")):
                _h = ws.cell(self.header_row, _c, value=_txt)
                self._clone_text_style(src_h, _h)
                _h.font = self._with_bold(_h.font)

            # 原列单列合并区 -> 非锚点行（这些行不单独写值/色，交由 merge 呈现锚点）
            non_anchor = set()
            for rng in [r for r in ws.merged_cells.ranges
                        if r.min_col == col_idx == r.max_col]:
                for rr in range(rng.min_row + 1, rng.max_row + 1):
                    non_anchor.add(rr)

            for rec in col_res:
                if rec["row"] in non_anchor:
                    continue
                src = ws.cell(rec["row"], col_idx)
                std_cell = ws.cell(rec["row"], std_col)
                self._fill_result_cell(std_cell, rec)
                self._clone_text_style(src, std_cell)
                sc = ws.cell(rec["row"], status_col)
                sc.value = f"{rec['rule']} | {rec['note']}"
                if self.colorize:
                    sc.fill = fill_for_rule(rec["rule"])
                self._clone_text_style(src, sc)

            n = self._mirror_single_col_merges(ws, col_idx, [std_col, status_col])
            self.log.info("[%s] %s -> 追加 %s(结果)+%s(说明)，复刻 %d 处合并",
                          col_letter, header_name,
                          get_column_letter(std_col), get_column_letter(status_col), n)

        self._write_legend(wb)
        wb.save(self.output_path)
        self.log.info("已保存：%s", self.output_path)

    # ---------- 落地：overwrite ----------
    def _write_overwrite(self, results):
        """原地覆盖原列，不生成说明列，结果文件可直接使用。"""
        wb = openpyxl.load_workbook(self.input_path, data_only=False)
        ws = wb[self.sheet_name]

        for col_letter, col_res in results.items():
            col_idx = col_res[0]["col_idx"]
            header_name = col_res[0]["header"] or col_letter

            for rec in col_res:
                cell = ws.cell(rec["row"], col_idx)
                if isinstance(cell, MergedCell):     # 非锚点只读，跳过（显示锚点）
                    continue
                self._fill_result_cell(cell, rec)

            # 说明：合并区非锚点格已跳过；锚点按自身处理类型上色，整块合并区显示锚点颜色
            self.log.info("[%s] %s -> 已覆盖原列", col_letter, header_name)

        self._write_legend(wb)
        wb.save(self.output_path)
        self.log.info("已保存：%s", self.output_path)


# ============================================================
# 库 API（供上层清稿工具调用）
# ============================================================
def unify_columns(input_path, columns, *, sheet=None, header_row=None,
                  row_start=None, row_end=None, mode="append", output_path=None,
                  text_unit_rules=None, target_unit=DEFAULT_TARGET_UNIT,
                  colorize=True, legend=True,
                  logger=None) -> CleanReport:
    """处理输入文件中的若干金额列，统一为目标单位（默认万元），返回结构化报告。

    参数
      input_path   输入 xlsx 路径
      columns      待处理列标识列表（列字母或列名），如 ["H", "I"]
      sheet        工作表名；None 自动选数据最多的表
      header_row   表头行号；None 自动探测
      row_start    起始数据行；None = header_row + 1
      row_end      结束数据行；None = 工作表最大行
      mode         "append"（默认，保留原列+结果列+说明列）| "overwrite"（覆盖原列）
      output_path  输出路径；None 自动生成 "{输入}_clean_{mode}_{时间戳}.xlsx"
      target_unit  目标单位：元/千元/万元/亿元（默认"万元"）
      text_unit_rules 自定义文本单位规则；提供则优先于按 target_unit 生成的默认规则
      colorize     是否用底色标注；legend 是否生成颜色图例
      logger       自定义 logging.Logger
    """
    cleaner = AmountColumnCleaner(
        input_path, sheet=sheet, header_row=header_row,
        row_start=row_start, row_end=row_end, mode=mode, output_path=output_path,
        text_unit_rules=text_unit_rules, target_unit=target_unit,
        colorize=colorize, legend=legend, logger=logger)
    return cleaner.run(columns)


# ============================================================
# 入口层：CLI
# ============================================================
def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="金额单位统一子任务：将表格中格式不统一的金额列统一换算为目标单位（默认万元）。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""使用示例：
  python amount_unifier.py -i input.xlsx -c H I
  python amount_unifier.py -i input.xlsx -c 金额（万元） -m overwrite
  python amount_unifier.py -i input.xlsx -c H I -s "Sheet1 (2)" -r 3 111 --no-color
  python amount_unifier.py -i input.xlsx -c H I --target-unit 元 --json

模式说明：
  append(默认)  保留原列，在原列后生成"结果列"+"说明列"，便于核对
  overwrite     覆盖原列，不生成说明列，结果文件可直接使用
""")
    parser.add_argument("-i", "--input", required=True, help="输入 Excel 文件路径")
    parser.add_argument("-c", "--columns", nargs="+", required=True,
                        help="待处理列，支持列字母（如 H I）或列名（如 金额（万元））")
    parser.add_argument("-s", "--sheet", default=None, help="工作表名（默认取第一张表）")
    parser.add_argument("--header-row", type=int, default=None, help="表头行号（默认自动探测）")
    parser.add_argument("-r", "--rows", nargs=2, type=int, default=None,
                        metavar=("START", "END"), help="数据行范围（默认表头行+1 到 末行）")
    parser.add_argument("-o", "--output", default=None, help="输出路径（默认自动生成）")
    parser.add_argument("-m", "--mode", choices=["append", "overwrite"], default="append",
                        help="append=追加结果列(默认) | overwrite=覆盖原列")
    parser.add_argument("--target-unit", choices=list(UNIT_SCALE), default=DEFAULT_TARGET_UNIT,
                        help="目标单位（默认万元）：元/千元/万元/亿元")
    parser.add_argument("--json", action="store_true",
                        help="以 JSON 输出机器可读报告（而非人眼文本）")
    parser.add_argument("--no-color", action="store_true", help="不使用底色标注")
    parser.add_argument("--no-legend", action="store_true", help="不生成颜色图例")
    parser.add_argument("--log-level", default="INFO",
                        choices=["DEBUG", "INFO", "WARNING", "ERROR"], help="日志级别")
    return parser


def _render_report(report: CleanReport):
    """人眼输出（仅在 CLI 使用，库调用不打印）。"""
    print("=" * 60)
    print(f"输出文件：{report.output_path}")
    print(f"处理模式：{report.mode} | 处理列：{', '.join(report.columns)}")
    print(f"总数 {report.total} | 成功 {report.success} | 需人工 {len(report.manual)}")
    print("规则分布：")
    for rule, n in report.by_rule.items():
        print(f"  {rule:16s}: {n}")
    if report.manual:
        print("-" * 60)
        print("需人工确认：")
        for m in report.manual:
            print(f"  [{m['coord']}] {m['rule']} | {m['note']}")
    print("=" * 60)


def main(argv=None) -> int:
    args = _build_parser().parse_args(argv)
    logging.basicConfig(level=getattr(logging, args.log_level),
                        format="%(levelname)s %(message)s")

    if not os.path.exists(args.input):
        log.error("输入文件不存在：%s", args.input)
        return 2

    row_start, row_end = args.rows if args.rows else (None, None)
    report = unify_columns(
        args.input, args.columns, sheet=args.sheet, header_row=args.header_row,
        row_start=row_start, row_end=row_end, mode=args.mode, output_path=args.output,
        target_unit=args.target_unit,
        colorize=not args.no_color, legend=not args.no_legend)

    if args.json:
        print(json.dumps(report.to_dict(), ensure_ascii=False, indent=2))
    else:
        _render_report(report)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
