"""达人筛选 skill V2 —— 内容匹配 × 价格效率 双核心二次精筛。

定位：巨量星图人工初筛后的 Excel 二次精筛，输出人工建联优先名单。
不承担投放管理；不修改任何业务数据库。

V2 相对 V1 的确定性与规则修复：
    1. 外部总分与等级来源一致（有外部总分时等级由外部总分判定，并做一致性复核）；
    2. 关键门槛字段缺失 → 不静默通过，进入「待复核-数据不全」；
    3. 百分比单位显式化（比率/百分数假设随结果输出，跨阈值歧义时复核）；
    4. 风险标签分级处理（排除级 / 复核级 / 提示级，未知标签默认复核）；
    5. 星图 ID 保留、去重并导出；输入文件指纹 + 评分版本随结果存档；
    6. 互动率 3% 一票否决取消 → 转质量复核信号；
    7. 内容匹配 × 价格效率双核心矩阵；无人工核验的二级内容证据时不得输出 P1。

用法：
    python kol_filter.py --input 达人数据.xlsx --output 推荐名单.xlsx \
        [--content-tags 内容标签.csv] [--archive-dir archive] [--top 10]
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
import yaml

CONFIG_PATH = Path(__file__).with_name("config.yaml")
VALID_GRADES = ("S", "A", "B", "C")


# ════════════════ 基础工具 ════════════════

def load_config(path: Path | None = None) -> dict:
    path = path or CONFIG_PATH
    if not path.exists():
        sys.exit(f"✗ 配置文件不存在：{path}")
    return yaml.safe_load(path.read_text(encoding="utf-8")) or {}


def file_sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def read_table(path: Path) -> pd.DataFrame:
    if not path.exists():
        sys.exit(f"✗ 输入文件不存在：{path}")
    if path.suffix.lower() in (".xlsx", ".xls"):
        return pd.read_excel(path)
    for enc in ("utf-8-sig", "utf-8", "gbk", "gb18030"):
        try:
            return pd.read_csv(path, encoding=enc)
        except UnicodeDecodeError:
            continue
    sys.exit(f"✗ 无法解码 CSV：{path}")


def to_num(value):
    """'9.9%' / '0.099' / '1.2万' / '3.5亿' → 数字；无法解析返回 None。"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "").replace("，", "")
    if not text or text.lower() in ("nan", "none", "-", "null"):
        return None
    mult = 1.0
    for suffix, factor in (("亿", 1e8), ("万", 1e4), ("w", 1e4), ("W", 1e4), ("k", 1e3), ("K", 1e3)):
        if text.endswith(suffix):
            mult, text = factor, text[: -len(suffix)]
            break
    if text.endswith("%"):
        mult, text = mult / 100, text[:-1]
    try:
        return float(text) * mult
    except ValueError:
        return None


def to_ratio(value):
    """互动率/完播率等比率字段解析。

    返回 (ratio, unit_note)：
      · 带 '%' → 百分数口径，直接除以 100（'3%' → 0.03）
      · 数字 > 1 → 假定为百分数写法（'3' → 0.03），unit_note='percent-assumed'
      · 0 < 数字 ≤ 1 → 假定为小数比率（'0.03' → 0.03），unit_note='ratio-assumed'
      · 数值 ≤ 0 → 原样（负增长等）
    单位假设会随结果输出；跨阈值歧义由复核信号层检测。
    """
    num = to_num(value)
    if num is None:
        return None, None
    if isinstance(value, str) and "%" in value:
        return num, "percent"
    if num > 1:
        return num / 100, "percent-assumed"
    return num, "ratio-assumed"


def band(value, table):
    if value is None:
        return None
    for threshold, score in table:
        if value >= threshold:
            return score
    return table[-1][1] if table else None


def weighted(parts):
    valid = [(v, w) for v, w in parts if v is not None]
    if not valid:
        return None
    total_w = sum(w for _, w in valid)
    return sum(v * w for v, w in valid) / total_w


def parse_tier_cpm(text):
    """星图「预期CPM」分档文本 → {植入口径(20-60s), 定制口径(60s+)}。
    口径已用 77 条真实数据验证：20-60s 档=植入报价口径，60s+ 档=定制报价口径。
    """
    m20 = re.search(r"20-60s[:：]\s*([0-9.]+)", str(text or ""))
    m60 = re.search(r"60s\+[:：]\s*([0-9.]+)", str(text or ""))
    return {
        "植入": float(m20.group(1)) if m20 else None,
        "定制": float(m60.group(1)) if m60 else None,
    }


def cpm_bucket(cpm, cfg) -> str | None:
    if cpm is None:
        return None
    for threshold, label in cfg["CPM分档"]["档位"]:
        if cpm <= threshold:
            return label
    return cfg["CPM分档"]["档位"][-1][1]


def display_width(text: str) -> int:
    return sum(2 if ord(ch) > 0x2E80 else 1 for ch in str(text))


def pad(text, width: int) -> str:
    text = str(text)
    return text + " " * max(0, width - display_width(text))


def fmt_pct(value, digits: int = 2) -> str:
    return "-" if value is None else f"{value * 100:.{digits}f}%"


# ════════════════ 字段自适应 / 内容标签 ════════════════

def pick_columns(df: pd.DataFrame, cfg: dict):
    picked, missing = {}, []
    for field, aliases in cfg["字段别名"].items():
        candidates = [a for a in aliases if a in df.columns]
        hit = max(candidates, key=lambda a: df[a].notna().sum()) if candidates else None
        picked[field] = hit
        if hit is None:
            missing.append(f"缺「{field}」列，相关规则跳过（候选列名：{'/'.join(aliases)}）")
    return picked, missing


def load_content_tags(path: Path | None, cfg: dict):
    """人工内容标签文件：按星图 ID 提供二级内容方向与证据。
    返回 {星图id(str): {一级, 二级, 证据, 来源, 核验}}；文件缺失返回空表。
    """
    if path is None or not Path(path).exists():
        return {}
    ext = Path(path).suffix.lower()
    if ext in (".xlsx", ".xls"):
        df = pd.read_excel(path, dtype={c: str for c in cfg["内容标签文件"]["支持列"]})
    else:
        df = pd.read_csv(path, dtype=str, encoding="utf-8-sig")
    tags = {}
    allowed_sources = set(cfg["内容标签文件"]["允许来源"])
    for _, row in df.iterrows():
        sid = str(row.get("星图id", "") or "").strip()
        if not sid:
            continue
        source = str(row.get("标签来源", "") or "").strip()
        status = str(row.get("核验状态", "") or "").strip()
        second = str(row.get("二级细分", "") or "").strip()
        evidence = str(row.get("内容证据", "") or "").strip()
        verified = status == "已核验" and (source in allowed_sources) and bool(second) and bool(evidence)
        tags[sid] = {
            "一级": str(row.get("一级方向", "") or "").strip() or "待核验",
            "二级": second or None,
            "证据": evidence or None,
            "来源": source or None,
            "核验状态": status or "待核验",
            "二级已核验": verified,
        }
    return tags


def build_records(df: pd.DataFrame, cols: dict, cfg: dict, tags: dict):
    ratio_fields = {"互动率", "完播率", "女性占比", "粉丝增长率", "资深中产", "新锐白领"}
    num_fields = {"粉丝量", "均播", "报价", "定制报价", "近30日视频数", "星图指数", "种草指数", "外部五维总分"}
    records, seen_ids, duplicates = [], {}, []
    for _, row in df.iterrows():
        rec = {"输入行": len(records) + 1}
        for field, col in cols.items():
            value = row[col] if col else None
            if field in ratio_fields:
                ratio, unit = to_ratio(value)
                rec[field], rec[f"{field}__单位"] = ratio, unit
            elif field in num_fields:
                rec[field] = to_num(value)
            else:
                rec[field] = "" if value is None or pd.isna(value) else str(value).strip()
        rec["星图id"] = str(rec.get("星图id") or "").strip()
        key = rec["星图id"] or f"昵称:{rec.get('昵称', '')}"
        if key in seen_ids:
            duplicates.append({"星图id": rec["星图id"] or None, "昵称": rec.get("昵称", ""), "输入行": rec["输入行"]})
            continue
        seen_ids[key] = rec["输入行"]
        rec["内容CPM"] = (rec["报价"] / rec["均播"] * 1000) if rec.get("报价") and rec.get("均播") else None
        tiers = parse_tier_cpm(rec.get("预期CPM文本", ""))
        rec["星图CPM_植入档"], rec["星图CPM_定制档"] = tiers["植入"], tiers["定制"]
        # 价格效率 CPM：植入口径（报价/预期播放）；预期播放缺失时回退星图植入口径档
        if rec.get("报价") and rec.get("均播"):
            rec["CPM_价格效率"] = round(rec["报价"] / rec["均播"] * 1000, 2)
            rec["CPM口径"] = "植入报价/预期播放"
        elif rec["星图CPM_植入档"] is not None:
            rec["CPM_价格效率"] = rec["星图CPM_植入档"]
            rec["CPM口径"] = "星图20-60s档"
        elif rec["星图CPM_定制档"] is not None:
            rec["CPM_价格效率"] = rec["星图CPM_定制档"]
            rec["CPM口径"] = "星图60s+档(定制)"
        else:
            rec["CPM_价格效率"], rec["CPM口径"] = None, None
        rec["CPM档位"] = cpm_bucket(rec["CPM_价格效率"], cfg)
        tag = tags.get(rec["星图id"]) if rec["星图id"] else None
        rec["内容标签"] = tag
        rec["二级内容证据"] = bool(tag and tag["二级已核验"])
        records.append(rec)
    return records, duplicates


# ════════════════ 五维辅助分（不再决定最终推荐） ════════════════

def score_five(rec: dict, cfg: dict):
    s = cfg["五维打分辅助"]
    w = s["维度权重"]
    fans = band(rec.get("粉丝量"), s["粉丝量分段"])
    active = band(rec.get("近30日视频数"), s["活跃度分段"])
    if active is None:
        active = band(rec.get("均播"), s["预期播放量分段"])
    health = weighted([(fans, w["账号健康度"]["粉丝量"]), (active, w["账号健康度"]["活跃度"])])
    completion = None if rec.get("完播率") is None else min(100.0, rec["完播率"] / s["完播率满分阈值"] * 100)
    quality = weighted([(rec.get("星图指数"), w["内容质量分"]["星图指数"]),
                        (completion, w["内容质量分"]["完播率"]),
                        (rec.get("种草指数"), w["内容质量分"]["种草指数"])])
    track = f"{rec.get('赛道', '')}{rec.get('标题', '')}"
    knowledge = s["知识赛道分"] if any(k in track for k in s["知识友好赛道"]) else s["垂类基线分"]
    if any(k in track for k in s["知识友好赛道"]):
        knowledge += min(s["知识赛道加分上限"], s["二级知识词加分"] * sum(1 for k in s["二级知识词"] if k in track))
    high_end = None
    if rec.get("资深中产") is not None or rec.get("新锐白领") is not None:
        high_end = min(100.0, ((rec.get("资深中产") or 0) + (rec.get("新锐白领") or 0)) / s["高知画像满分占比"] * 100)
    tone = weighted([(float(knowledge), w["知乎调性匹配分"]["垂类知识性"]),
                     (high_end, w["知乎调性匹配分"]["高知画像"])])
    rate_score = None if rec.get("互动率") is None else min(100.0, rec["互动率"] / s["互动率满分阈值"] * 100)
    interaction = weighted([(rate_score, w["互动质量分"]["互动率"]), (completion, w["互动质量分"]["完播率"])])
    lo, hi = s["互动率合理区间"]
    rate_ok = None if rec.get("互动率") is None else (
        100.0 if lo <= rec["互动率"] <= hi else s["高于上限分"] if rec["互动率"] > hi else s["不合理分"])
    g_lo, g_hi = s["粉丝增长率合理区间"]
    growth_ok = None if rec.get("粉丝增长率") is None else (
        100.0 if g_lo <= rec["粉丝增长率"] <= g_hi else s["不合理分"])
    plays_ok = None
    if rec.get("粉丝量") is not None and rec.get("均播") is not None:
        plays_ok = s["刷量分"] if any(rec["粉丝量"] >= f and rec["均播"] < p for f, p in s["均播刷量阈值"]) else 100.0
    authenticity = weighted([(rate_ok, w["数据真实度分"]["互动率合理性"]),
                             (growth_ok, w["数据真实度分"]["粉丝增长率合理性"]),
                             (plays_ok, w["数据真实度分"]["均播合理性"])])
    dims = {"账号健康度": health, "内容质量分": quality, "知乎调性匹配分": tone,
            "互动质量分": interaction, "数据真实度分": authenticity}
    total = weighted([(v, s["五维权重"][k]) for k, v in dims.items()])
    return (total or 0.0), dims


def grade_of(score: float, cfg: dict) -> str:
    for grade, threshold in cfg["五维打分辅助"]["辅助等级阈值"].items():
        if score >= threshold:
            return grade
    return "C"


def ensure_grade(rec: dict, cfg: dict):
    """修复 V1 问题 1：外部总分与等级来源一致。
    有外部总分 → 等级必须由外部总分判定；外部等级与外部总分矛盾 → 复核。
    无外部总分 → 使用内部辅助分（标记为辅助证据）。
    """
    internal_total, dims = score_five(rec, cfg)
    rec["_辅助五维分"] = round(internal_total, 2)
    rec["_辅助等级"] = grade_of(internal_total, cfg)
    rec["复核原因"] = rec.get("复核原因", [])
    external_total = rec.get("外部五维总分")
    external_grade = rec.get("外部等级")
    if external_total is not None:
        rec["五维总分_来源"] = "外部输入"
        derived = grade_of(external_total, cfg)
        if external_grade in VALID_GRADES and external_grade != derived:
            rec["复核原因"].append(f"外部总分({external_total})推导等级({derived})与外部等级({external_grade})不一致")
        rec["等级"] = external_grade if external_grade in VALID_GRADES else derived
        rec["等级_依据"] = "外部总分"
    else:
        rec["五维总分_来源"] = "内部辅助分"
        rec["等级"] = rec["_辅助等级"]
        rec["等级_依据"] = "内部辅助分"


# ════════════════ 排除 / 复核 / 分级 ════════════════

def classify_risk_tag(tag_text: str, cfg: dict) -> str:
    level_cfg = cfg["风险标签分级"]
    invalid = {v.lower() for v in level_cfg["无效值"]}
    tags = [t.strip() for t in re.split(r"[,，;；/|]", str(tag_text)) if t.strip()]
    tags = [t for t in tags if t.lower() not in invalid]
    if not tags:
        return "无"
    levels = {"排除级": level_cfg["排除级"], "复核级": level_cfg["复核级"], "提示级": level_cfg["提示级"]}
    for tag in tags:
        for level, words in levels.items():
            if any(word in tag or tag in word for word in words):
                return level if level != "提示级" else "提示"
    return level_cfg["未知标签默认级别"]


def hard_exclusions(rec: dict, cfg: dict) -> list[str]:
    """V2 仅保留品牌安全类硬排除：避雷词 + 排除级风险标签 + 基础粉丝/报价底线。"""
    g, fails = cfg["硬排除"], []
    if rec.get("粉丝量") is not None and not (rec["粉丝量"] > g["粉丝量下限"]):
        fails.append(f"粉丝量不足（{rec['粉丝量']:,.0f} ≤ {g['粉丝量下限']:,}）")
    price_lo, price_hi = g["报价区间"]
    if rec.get("报价") is not None and not (price_lo <= rec["报价"] <= price_hi):
        fails.append(f"报价超范围（{rec['报价']:,.0f} 元，要求 {price_lo}-{price_hi}）")
    if g["检查避雷词"]:
        text = f"{rec.get('昵称', '')} {rec.get('赛道', '')} {rec.get('标题', '')}"
        hits = [word for word in cfg["避雷词表"] if word in text]
        if hits:
            fails.append(f"命中避雷词（{'、'.join(hits)}）")
    risk_level = classify_risk_tag(rec.get("风险标签", ""), cfg)
    if risk_level == "排除级":
        fails.append(f"排除级风险标签（{rec['风险标签']}）")
    rec["风险级别"] = "无" if risk_level == "无" else risk_level
    return fails


def review_flags(rec: dict, cfg: dict, median_rate, cols_present):
    f, flags = cfg["复核信号"], []
    # V2：门槛字段缺失 → 数据不全（原 V1 静默通过问题）
    missing_gates = [name for name in ("粉丝量", "互动率", "报价", "均播") if rec.get(name) is None]
    if missing_gates:
        flags.append("门槛数据不全（缺：" + "、".join(missing_gates) + "）")
    rate = rec.get("互动率")
    if rate is not None and rate < f["互动率下限（转复核）"]:
        flags.append(f"互动率偏低（{fmt_pct(rate)} < {fmt_pct(f['互动率下限（转复核）'], 1)}，转复核不排除）")
    rule = f["高互动低完播"]
    if rate is not None and rec.get("完播率") is not None:
        if rate > rule["互动率下限"] and rec["完播率"] < rule["完播率上限"]:
            flags.append(f"高互动低完播（互动率 {fmt_pct(rate)} / 完播率 {fmt_pct(rec['完播率'])}）")
    if f["粉丝负增长"] and rec.get("粉丝增长率") is not None and rec["粉丝增长率"] < 0:
        flags.append(f"粉丝负增长（{rec['粉丝增长率'] * 100:.2%}）")
    if rate is not None and rate > f["互动率上限（异常）"]:
        flags.append(f"互动率异常偏高（{fmt_pct(rate)} > {fmt_pct(f['互动率上限（异常）'], 0)}，口径待核验）")
    if rate is not None and rec.get("完播率") is not None and rec["完播率"] > f["完播率上限（异常）"]:
        flags.append(f"完播率异常偏高（{fmt_pct(rec['完播率'])}，口径待核验）")
    cpm = rec.get("CPM_价格效率")
    if cpm is not None and cpm < f["重算CPM下限（异常）"]:
        flags.append(f"重算CPM异常偏低（{cpm}，预期播放疑含自然流量）")
    if median_rate and rate is not None and rate > median_rate * f["互动率超批内中位数倍数"]:
        flags.append(f"互动率超批内中位数 {f['互动率超批内中位数倍数']} 倍")
    video_count = rec.get("近30日视频数")
    if f["视频数缺失或为零"] and video_count == 0:
        flags.append("近30日星图视频数为 0（停更信号待核验）")
    if not rec.get("二级内容证据"):
        flags.append("内容细分待核验（无人工核验的二级内容证据）")
    if rec.get("_内外评分不一致"):
        flags.append("外部总分与外部等级不一致")
    return flags


def content_match_score(rec: dict, cfg: dict) -> float:
    w = cfg["双维矩阵"]["内容匹配权重"]
    star = rec.get("星图指数")
    seed = rec.get("种草指数")
    star_n = min(100.0, star / 100 * 100) if star is not None else None
    seed_n = min(100.0, seed / 100 * 100) if seed is not None else None
    tag = rec.get("内容标签")
    base = cfg["双维矩阵"]["垂类基础分"]
    if tag and tag.get("二级"):
        base = cfg["双维矩阵"]["垂类基础分"] + cfg["双维矩阵"]["知识友好垂类加成"] \
            if any(k in str(tag["二级"]) for k in cfg["双维矩阵"]["知识友好二级词"]) else base
    base = min(100.0, base)
    parts = []
    if star_n is not None:
        parts.append((star_n, w["星图指数"]))
    if seed_n is not None:
        parts.append((seed_n, w["种草指数"]))
    if tag and tag.get("二级"):
        parts.append((base, w["垂类基础分"]))
    else:
        parts.append((base * 0.8, w["垂类基础分"]))   # 无二级证据 → 垂类维度打折（不进 P1 的机制之一）
    return round(weighted(parts) or 0.0, 1)


def assign_tier(rec: dict, cfg: dict) -> tuple[str, list[str]]:
    m = cfg["双维矩阵"]
    reasons = []
    match = rec["内容匹配分"]
    bucket = rec["CPM档位"]
    evidence = rec["二级内容证据"]
    high, mid = m["高匹配阈值"], m["中匹配阈值"]
    high_bucket = bucket in ("高效", "良好")
    mid_bucket = bucket == "可谈"
    low_bucket = bucket == "低效"
    if evidence and match >= high and high_bucket:
        return "P1-优先建联", reasons
    if match >= high and (high_bucket or mid_bucket):
        tier = "P2-推荐建联"
        if not evidence:
            reasons.append("内容细分待核验（无二级证据，P1 需人工补标后重新运行）")
        return tier, reasons
    if match >= mid and high_bucket:
        tier = "P2-推荐建联"
        if not evidence:
            reasons.append("内容细分待核验（无二级证据）")
        return tier, reasons
    if match >= mid and mid_bucket:
        return "P3-观察候选", reasons
    if low_bucket and match >= m["议价候选匹配阈值"]:
        return "议价候选-评估折扣空间", reasons
    if match < mid:
        return "待复核-内容匹配不足", reasons
    return "P3-观察候选", reasons


# ════════════════ 输出 ════════════════

def export_excel(rows, excluded, duplicates, stats, output: Path, cfg: dict):
    output.parent.mkdir(parents=True, exist_ok=True)
    rec_cols = ["达人编号", "星图id", "昵称", "建联优先级", "内容匹配分", "CPM_价格效率", "CPM口径", "CPM档位",
                "价格效率分档", "二级内容方向", "内容证据", "内容细分待核验", "互动率(%)", "完播率(%)",
                "辅助五维分", "辅助等级", "五维总分_来源", "风险级别", "复核原因", "建联建议",
                "报价", "定制报价", "粉丝量", "预期播放量", "女性占比(%)", "主页链接", "排除原因"]
    rec_df = pd.DataFrame([{c: r.get(c) for c in rec_cols} for r in rows])
    exc_df = pd.DataFrame(excluded)
    stat_df = pd.DataFrame(stats, columns=["指标", "值"])
    dup_df = pd.DataFrame(duplicates) if duplicates else pd.DataFrame(columns=["星图id", "昵称", "输入行"])
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        rec_df.to_excel(writer, sheet_name="建联优先名单", index=False)
        exc_df.to_excel(writer, sheet_name="排除清单", index=False)
        dup_df.to_excel(writer, sheet_name="去重剔除", index=False)
        stat_df.to_excel(writer, sheet_name="筛选统计", index=False)


def build_reason(rec: dict) -> str:
    parts = []
    if rec.get("二级内容方向"):
        parts.append(f"内容方向：{rec['二级内容方向']}（{rec.get('内容证据') or '无证据记录'}）")
    elif rec.get("内容细分待核验"):
        parts.append("内容细分待核验")
    if rec.get("CPM_价格效率") is not None:
        parts.append(f"{rec.get('CPM口径')} CPM {rec['CPM_价格效率']}（{rec.get('CPM档位')}档）")
    if rec.get("报价") is not None:
        parts.append(f"植入报价 {rec['报价']:,.0f} 元")
    if rec.get("互动率") is not None:
        parts.append(f"互动率 {fmt_pct(rec['互动率'])}")
    if rec.get("复核原因"):
        parts.append("复核：" + "；".join(rec["复核原因"]))
    return "；".join(parts)


# ════════════════ 主流程 ════════════════

def run_filter(input_path: Path, output_path: Path, cfg: dict,
               content_tags_path: Path | None = None, archive_dir: Path | None = None):
    fingerprint = file_sha256(input_path)
    df = read_table(input_path)
    cols, missing = pick_columns(df, cfg)
    tags = load_content_tags(content_tags_path, cfg)
    records, duplicates = build_records(df, cols, cfg, tags)

    rates = [r["互动率"] for r in records if r.get("互动率") is not None]
    median_rate = float(pd.Series(rates).median()) if rates else None
    cols_present = {f for f, c in cols.items() if c}

    excluded, kept = [], []
    for rec in records:
        rec["复核原因"] = []
        ensure_grade(rec, cfg)
        rec["内容匹配分"] = content_match_score(rec, cfg)
        rec["风险级别"] = "无"
        rec["排除原因"] = "；".join(hard_exclusions(rec, cfg))
        rec["复核原因"] = review_flags(rec, cfg, median_rate, cols_present)
        rec["内容细分待核验"] = not rec["二级内容证据"]
        rec["二级内容方向"] = (rec.get("内容标签") or {}).get("二级") or None
        rec["内容证据"] = (rec.get("内容标签") or {}).get("证据") or None
        rec["建联优先级"], tier_notes = assign_tier(rec, cfg)
        rec["复核原因"] = rec["复核原因"] + tier_notes
        rec["价格效率分档"] = rec["CPM档位"] or "-"
        rec["互动率(%)"] = round(rec["互动率"] * 100, 2) if rec.get("互动率") is not None else None
        rec["完播率(%)"] = round(rec["完播率"] * 100, 2) if rec.get("完播率") is not None else None
        rec["女性占比(%)"] = round(rec["女性占比"] * 100, 1) if rec.get("女性占比") is not None else None
        rec["内容细分待核验"] = not rec["二级内容证据"]
        rec["建联建议"] = build_reason(rec)
        rec["达人编号"] = f"R{rec['输入行']:03d}"
        (excluded if rec["排除原因"] else kept).append(rec)

    tier_order = ["P1-优先建联", "P2-推荐建联", "议价候选-评估折扣空间", "P3-观察候选",
                  "待复核-数据不全", "待复核-内容匹配不足", "待复核-低匹配"]
    kept.sort(key=lambda r: (tier_order.index(r["建联优先级"]) if r["建联优先级"] in tier_order else 99,
                             -r["内容匹配分"], -(r.get("CPM_价格效率") or 0)))
    for i, rec in enumerate(kept, 1):
        rec["排名"] = i

    tier_counts = {}
    for rec in kept + excluded:
        tier_counts[rec["建联优先级"]] = tier_counts.get(rec["建联优先级"], 0) + 1
    stats = [
        ("配置版本", cfg["version"]), ("评分版本", cfg["评分版本"]),
        ("输入文件指纹(SHA256)", fingerprint[:16] + "…"),
        ("原始达人数", len(records)), ("去重剔除", len(duplicates)),
        ("硬排除", len(excluded)), ("进入分级", len(kept)),
        ("二级内容已核验", sum(1 for r in records if r["二级内容证据"])),
        ("内容细分待核验", sum(1 for r in records if not r["二级内容证据"])),
        ("待复核人数", sum(1 for r in kept if str(r["建联优先级"]).startswith("待复核"))),
        ("批内互动率中位数", "-" if median_rate is None else round(median_rate, 4)),
        ("内容标签文件", str(content_tags_path) if content_tags_path else "未提供"),
    ]
    export_excel(kept + [dict(r, 排除原因=r["排除原因"]) for r in excluded], excluded, duplicates, stats, output_path, cfg)

    archive_dir = archive_dir or output_path.parent / cfg["存档"]["目录"]
    archive_dir.mkdir(parents=True, exist_ok=True)
    archive = {
        "archived_at": datetime.now().isoformat(timespec="seconds"),
        "config_version": cfg["version"], "scoring_version": cfg["评分版本"],
        "input_file": str(input_path), "input_sha256": fingerprint,
        "content_tags_file": str(content_tags_path) if content_tags_path else None,
        "tier_counts": tier_counts,
        "records": [{"达人编号": r["达人编号"], "星图id": r.get("星图id") or None,
                     "建联优先级": r["建联优先级"], "内容匹配分": r["内容匹配分"],
                     "CPM_价格效率": r.get("CPM_价格效率"), "辅助五维分": r["_辅助五维分"],
                     "二级内容证据": r["二级内容证据"], "排除原因": r.get("排除原因") or None}
                    for r in kept + excluded],
    }
    archive_path = archive_dir / f"filter_run_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{fingerprint[:8]}.json"
    archive_path.write_text(json.dumps(archive, ensure_ascii=False, indent=1), encoding="utf-8")
    return {"kept": kept, "excluded": excluded, "duplicates": duplicates, "stats": stats,
            "archive_path": archive_path, "fingerprint": fingerprint, "tier_counts": tier_counts}


def main() -> None:
    started = time.time()
    parser = argparse.ArgumentParser(description="达人筛选 skill V2：内容匹配 × 价格效率 双核心二次精筛")
    parser.add_argument("--input", required=True, help="输入 CSV/Excel 路径（星图初筛导出）")
    parser.add_argument("--output", required=True, help="输出 Excel 路径")
    parser.add_argument("--config", default=str(CONFIG_PATH), help="配置文件路径")
    parser.add_argument("--content-tags", default=None, help="人工内容标签文件（CSV/Excel，按星图id）")
    parser.add_argument("--archive-dir", default=None, help="筛选结果存档目录（默认输出目录/archive）")
    parser.add_argument("--top", type=int, default=10, help="终端打印前 N 名")
    args = parser.parse_args()

    cfg = load_config(Path(args.config))
    print(f"[筛选V2] 配置 v{cfg['version']}｜评分 v{cfg['评分版本']}｜输入：{args.input}")
    result = run_filter(Path(args.input), Path(args.output), cfg,
                        Path(args.content_tags) if args.content_tags else None,
                        Path(args.archive_dir) if args.archive_dir else None)

    for msg in result["stats"]:
        print(f"[筛选V2] {msg[0]}：{msg[1]}")
    print("[筛选V2] Top 建联优先名单：")
    for rec in result["kept"][: args.top]:
        print(f"  {rec['排名']:>3} {pad(rec['昵称'], 16)} {pad(rec['建联优先级'], 12)} "
              f"匹配 {rec['内容匹配分']:>5}｜{pad(str(rec.get('CPM档位')), 4)} CPM {rec.get('CPM_价格效率') or '-'}")
    flagged = [r for r in result["kept"] if r["复核原因"]]
    if flagged:
        print(f"[筛选V2] 建议人工复核 {len(flagged)} 人（仅标记，不排除）")
    print(f"[筛选V2] 存档：{result['archive_path']}")
    print(f"[筛选V2] 产物：{Path(args.output).resolve()}｜耗时 {time.time() - started:.2f} 秒")


if __name__ == "__main__":
    main()
