"""从纯文本里用正则抓取关键字段：金额、合同号/编号、日期、人名等。

设计成可扩展的"字段规则"：每条规则有名字、正则、说明。
批量汇总 100 份合同时，对每份文档跑一遍规则即可得到对照表。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field


@dataclass
class FieldRule:
    key: str
    label: str
    pattern: re.Pattern
    note: str = ""


# 金额：支持 ¥ / ￥ / RMB / 元 / 万元，含千分位与小数。
# 关键：裸整数（如日期里的 2024、编号里的 001）不算金额——
# 必须带货币符号、货币单位或千分位逗号之一，避免把日期/编号误当金额。
_AMOUNT = re.compile(
    r"(?P<cur>[¥￥])?\s*"
    r"(?P<num>(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d{1,2})?)"
    r"\s*(?P<unit>万元|元|块|人民币|RMB|CNY)?",
    re.IGNORECASE,
)

# 合同号 / 编号：形如 合同编号：ABC-2024-001 / No.12345 / 单号 XH2024001
_CONTRACT_NO = re.compile(
    r"(?:合同编号|合同号|协议编号|单号|编号|No\.?|No:)[:：\s]*"
    r"([A-Za-z0-9\-_/（）()#]{3,40})",
    re.IGNORECASE,
)

# 日期：2024-01-05 / 2024年1月5日 / 2024/01/05
_DATE = re.compile(
    r"(\d{4}[-/年]\s*\d{1,2}[-/月]\s*\d{1,2}\s*日?)"
)

# 甲乙方：甲方：XXX公司
_PARTY = re.compile(r"(甲\s*方|乙\s*方|丙\s*方)[:：]\s*([^\n，,。]{2,40})")

DEFAULT_RULES: list[FieldRule] = [
    FieldRule("amount", "金额", _AMOUNT, "所有形似金额的数字（含¥/元/万元）"),
    FieldRule("contract_no", "合同/编号", _CONTRACT_NO, "合同编号、单号、编号"),
    FieldRule("date", "日期", _DATE, "形如 2024-01-05 / 2024年1月5日"),
    FieldRule("party", "甲乙方", _PARTY, "甲方、乙方、丙方名称"),
]


def _normalize_amount(num: str, unit: str | None) -> str:
    """把 1,234.5 万元 统一成可读文本；保留原单位，便于人工核对。"""
    try:
        val = float(num.replace(",", ""))
    except ValueError:
        return num
    if unit in ("万元",):
        return "%.2f万元(≈%.2f元)" % (val, val * 10000)
    return ("%.2f" % val).rstrip("0").rstrip(".") + (unit or "")


def extract_fields(text: str, rules: list[FieldRule] | None = None) -> dict[str, list[str]]:
    """对一段文本按规则抓字段，返回 {key: [命中值...]}（已去重保序）。"""
    rules = rules or DEFAULT_RULES
    out: dict[str, list[str]] = {}
    for rule in rules:
        hits: list[str] = []
        seen = set()
        for m in rule.pattern.finditer(text):
            if rule.key == "amount":
                cur, num, unit = m.group("cur"), m.group("num"), m.group("unit")
                # 裸数字（无货币符号/单位、无千分位）视为日期或编号里的数字，跳过
                has_thousands = "," in num
                if not (cur or unit or has_thousands):
                    continue
                raw = _normalize_amount(num, unit)
                if cur and not unit:
                    raw = cur + raw
            elif rule.key == "party":
                raw = "%s: %s" % (m.group(1).replace(" ", ""), m.group(2).strip())
            else:
                raw = m.group(1).strip()
            if raw and raw not in seen:
                seen.add(raw)
                hits.append(raw)
        out[rule.key] = hits
    return out


def summarize_row(res, rules: list[FieldRule] | None = None) -> dict[str, str]:
    """把一份文档的字段结果压成一行（多值用 ' | ' 连接），用于汇总表。"""
    rules = rules or DEFAULT_RULES
    fields = extract_fields(res.text, rules)
    row: dict[str, str] = {"文件": res.doc.name, "路径": res.doc.path}
    if res.error:
        row["错误"] = res.error
    for rule in rules:
        row[rule.label] = " | ".join(fields.get(rule.key, []))
    return row