"""达人筛选 skill v1.0 —— 一条命令：筛人 + 排序 + 导出。

用法：
    python kol_filter.py --input 达人数据.xlsx --output 推荐清单.xlsx

本次只修 3 个问题（其余逻辑沿用 v2 已验证标准）：
    1. 排序口径：梯队由五维分等级决定（S/A→第一，B→第二，C→第三），加分项只在同梯队内排序；
    2. 异常复核标记：命中 4 类信号的人标「建议人工复核」，只标记不排除；
    3. 阈值全部来自同目录 config.yaml，本文件不含业务阈值。

字段自适应：缺列时打印「缺 XX 列，相关规则跳过」，不中断执行。
"""
from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

import pandas as pd
import yaml

CONFIG_PATH = Path(__file__).with_name("config.yaml")
TIER_ORDER = ["第一梯队-强烈推荐", "第二梯队-推荐", "第三梯队-可考虑"]


# ════════════════ 基础工具 ════════════════

def load_config(path: Path | None = None) -> dict:
    path = path or CONFIG_PATH
    if not path.exists():
        sys.exit(f"✗ 配置文件不存在：{path}")
    return yaml.safe_load(path.read_text(encoding="utf-8")) or {}


def read_table(path: Path) -> pd.DataFrame:
    """CSV / Excel 自适应读取（CSV 依次尝试 utf-8-sig / gbk / gb18030）。"""
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
    """把互动率/完播率统一成小数：0.03 与 '3%' 都 → 0.03；'3' → 0.03（星图有时导出为百分数）。"""
    num = to_num(value)
    if num is None:
        return None
    if isinstance(value, str) and "%" in value:
        return num
    return num / 100 if num > 1 else num


def band(value, table) -> int | None:
    """分段打分：table = [[阈值, 分值], ...]，取第一个 value ≥ 阈值 的分值。"""
    if value is None:
        return None
    for threshold, score in table:
        if value >= threshold:
            return score
    return table[-1][1] if table else None


def weighted(parts: list[tuple[float | None, float]]) -> float | None:
    """按权重合并；缺失子项跳过并重新归一化权重。"""
    valid = [(v, w) for v, w in parts if v is not None]
    if not valid:
        return None
    total_w = sum(w for _, w in valid)
    return sum(v * w for v, w in valid) / total_w


def display_width(text: str) -> int:
    return sum(2 if ord(ch) > 0x2E80 else 1 for ch in str(text))


def pad(text, width: int) -> str:
    text = str(text)
    return text + " " * max(0, width - display_width(text))


def fmt_pct(value, digits: int = 2) -> str:
    return "-" if value is None else f"{value * 100:.{digits}f}%"


def fmt_int(value) -> str:
    return "-" if value is None else f"{value:,.0f}"


# ════════════════ 字段自适应 ════════════════

def pick_columns(df: pd.DataFrame, cfg: dict) -> tuple[dict, list[str]]:
    """按 config 的别名表匹配列：同一字段有多个候选列同时存在时，取非空值最多的那一列。"""
    picked, missing = {}, []
    for field, aliases in cfg["字段别名"].items():
        candidates = [a for a in aliases if a in df.columns]
        hit = max(candidates, key=lambda a: df[a].notna().sum()) if candidates else None
        picked[field] = hit
        if hit is None:
            missing.append(f"缺「{field}」列，相关规则跳过（候选列名：{'/'.join(aliases)}）")
    return picked, missing


def build_records(df: pd.DataFrame, cols: dict, cfg: dict) -> list[dict]:
    """把原始 DataFrame 转成标准字段记录（每行一位达人）。"""
    ratio_fields = {"互动率", "完播率", "女性占比", "粉丝增长率", "资深中产", "新锐白领"}
    num_fields = {"粉丝量", "均播", "报价", "近30日视频数", "星图指数", "种草指数", "五维总分"}
    records = []
    for _, row in df.iterrows():
        rec = {}
        for field, col in cols.items():
            value = row[col] if col else None
            rec[field] = to_ratio(value) if field in ratio_fields else \
                to_num(value) if field in num_fields else \
                ("" if value is None or pd.isna(value) else str(value).strip())
        rec["内容形态"] = content_form(f"{rec.get('赛道', '')} {rec.get('标题', '')}", cfg)
        rec["CPM"] = rec["报价"] / rec["均播"] * 1000 if rec.get("报价") and rec.get("均播") else None
        records.append(rec)
    return records


def content_form(text: str, cfg: dict) -> str:
    for name, keywords in cfg["内容形态词表"].items():
        if any(k in text for k in keywords):
            return name
    return "其他"


# ════════════════ 五维打分（问题 1 的依据） ════════════════

def score_five(rec: dict, cfg: dict) -> tuple[float, dict, list[str]]:
    """五维分（0-100）→ 加权总分。打分口径移植自主项目星图模式，改为绝对分档。"""
    s = cfg["五维打分"]
    w = s["维度权重"]
    notes: list[str] = []

    # 账号健康度：粉丝量 + 近30日视频数（缺失时用预期播放量代理）
    fans = band(rec.get("粉丝量"), s["粉丝量分段"])
    active = band(rec.get("近30日视频数"), s["活跃度分段"])
    if active is None:
        active = band(rec.get("均播"), s["预期播放量分段"])
        if active is not None:
            notes.append("近30日视频数缺失，用预期播放量代理健康度")
    health = weighted([(fans, w["账号健康度"]["粉丝量"]), (active, w["账号健康度"]["活跃度"])])

    # 内容质量分：星图指数 + 完播率 + 种草指数
    completion = None if rec.get("完播率") is None else \
        min(100.0, rec["完播率"] / s["完播率满分阈值"] * 100)
    quality = weighted([(rec.get("星图指数"), w["内容质量分"]["星图指数"]),
                        (completion, w["内容质量分"]["完播率"]),
                        (rec.get("种草指数"), w["内容质量分"]["种草指数"])])

    # 知乎调性匹配分：垂类知识性 + 高知画像
    track = rec.get("赛道", "") + rec.get("标题", "")
    knowledge = s["垂类基线分"]
    if any(k in track for k in s["知识友好赛道"]):
        knowledge = s["知识赛道分"] + min(s["知识赛道加分上限"],
                                           s["二级知识词加分"] * sum(1 for k in s["二级知识词"] if k in track))
    high_end = None
    if rec.get("资深中产") is not None or rec.get("新锐白领") is not None:
        ratio = (rec.get("资深中产") or 0) + (rec.get("新锐白领") or 0)
        high_end = min(100.0, ratio / s["高知画像满分占比"] * 100)
    tone = weighted([(float(knowledge), w["知乎调性匹配分"]["垂类知识性"]),
                     (high_end, w["知乎调性匹配分"]["高知画像"])])

    # 互动质量分：互动率 + 完播率
    rate_score = None if rec.get("互动率") is None else \
        min(100.0, rec["互动率"] / s["互动率满分阈值"] * 100)
    interaction = weighted([(rate_score, w["互动质量分"]["互动率"]),
                            (completion, w["互动质量分"]["完播率"])])

    # 数据真实度分：互动率合理性 + 粉丝增长率 + 均播合理性
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
    total = weighted([(v, cfg["五维打分"]["五维权重"][k]) for k, v in dims.items()])
    return (total or 0.0), dims, notes


def grade_of(score: float, cfg: dict) -> str:
    for grade, threshold in cfg["五维打分"]["等级阈值"].items():
        if score >= threshold:
            return grade
    return "C"


def ensure_grade(rec: dict, cfg: dict) -> None:
    """优先用输入自带的五维分/等级；没有则内部打分。"""
    total, dims, notes = score_five(rec, cfg)
    rec["_dims"], rec["_score_notes"] = dims, notes
    rec["五维总分"] = rec["五维总分"] if rec.get("五维总分") is not None else round(total, 2)
    rec["等级"] = rec["等级"] if rec.get("等级") in ("S", "A", "B", "C") else grade_of(total, cfg)


# ════════════════ 规则：门槛 / 加分 / 复核 ════════════════

def hard_gates(rec: dict, cfg: dict) -> list[str]:
    g = cfg["硬门槛"]
    fails = []
    if rec.get("粉丝量") is not None and not (rec["粉丝量"] > g["粉丝量下限"]):
        fails.append(f"粉丝量不足（{rec['粉丝量']:,.0f} ≤ {g['粉丝量下限']:,}）")
    if rec.get("互动率") is not None and not (rec["互动率"] > g["互动率下限"]):
        fails.append(f"互动率低（{fmt_pct(rec['互动率'])} ≤ {fmt_pct(g['互动率下限'], 0)}）")
    if rec.get("CPM") is not None and not (rec["CPM"] < g["CPM上限"]):
        fails.append(f"CPM 过高（{rec['CPM']:.2f} ≥ {g['CPM上限']}）")
    if rec.get("均播") is not None and not (rec["均播"] > g["均播下限"]):
        fails.append(f"均播低（{rec['均播']:,.0f} ≤ {g['均播下限']:,}）")
    price_lo, price_hi = g["报价区间"]
    if rec.get("报价") is not None and not (price_lo <= rec["报价"] <= price_hi):
        fails.append(f"报价超范围（{rec['报价']:,.0f} 元，要求 {price_lo}-{price_hi}）")
    if g["检查避雷词"]:
        text = f"{rec.get('昵称', '')} {rec.get('赛道', '')} {rec.get('标题', '')}"
        hits = [word for word in cfg["避雷词表"] if word in text]
        if hits:
            fails.append(f"命中避雷词（{'、'.join(hits)}）")
    if g["检查风险标签"] and rec.get("风险标签"):
        if rec["风险标签"] not in ("无", "未知", "none", "NaN"):
            fails.append(f"有风险标签（{rec['风险标签']}）")
    return fails


def bonus_points(rec: dict, cfg: dict) -> tuple[int, list[str]]:
    """加分项：只影响同梯队内排序。"""
    b = cfg["加分项"]
    score, notes = 0, []
    female = rec.get("女性占比")
    if female is not None:
        f = b["女性占比"]
        if female * 100 > f["高分阈值"]:
            score += f["高分加分"]; notes.append(f"女性占比 {female * 100:.0f}%（>70%）+{f['高分加分']}")
        elif female * 100 > f["低分阈值"]:
            score += f["低分加分"]; notes.append(f"女性占比 {female * 100:.0f}%（>60%）+{f['低分加分']}")
    if rec.get("内容形态") in b["内容形态加分形态"]:
        score += b["内容形态加分"]; notes.append(f"内容形态={rec['内容形态']} +{b['内容形态加分']}")
    hit_words = [w for w in b["标题场景词"] if w in rec.get("标题", "")]
    if hit_words:
        score += b["标题词加分"]; notes.append(f"标题命中 {'、'.join(hit_words)} +{b['标题词加分']}")
    rate = rec.get("互动率")
    if rate is not None:
        r = b["互动率"]
        if rate * 100 > r["第一档阈值"]:
            score += r["第一档加分"]; notes.append(f"互动率 {rate * 100:.1f}%（>10%）+{r['第一档加分']}")
        elif rate * 100 > r["第二档阈值"]:
            score += r["第二档加分"]; notes.append(f"互动率 {rate * 100:.1f}%（>8%）+{r['第二档加分']}")
    if rec.get("完播率") is not None and rec["完播率"] * 100 > b["完播率阈值"]:
        score += b["完播率加分"]; notes.append(f"完播率 {rec['完播率'] * 100:.1f}%（>10%）+{b['完播率加分']}")
    price = rec.get("报价")
    if price is not None and any(lo <= price <= hi for lo, hi in b["价格带"]):
        score += b["价格带加分"]; notes.append(f"报价 {price:,.0f} 元（性价比价带）+{b['价格带加分']}")
    if any(k in rec.get("赛道", "") for k in b["赛道关键词"]):
        score += b["赛道加分"]; notes.append(f"赛道加分 +{b['赛道加分']}")
    return score, notes


def review_flags(rec: dict, cfg: dict, median_rate: float | None,
                 cols_present: set[str], skip_video_missing: bool = False) -> list[str]:
    """异常复核信号：命中即标「建议人工复核」，不排除、不降权。

    cols_present 记录哪些标准字段在输入里真的存在，用于区分「整列缺失」（已全局提示）
    与「单行单元格缺失」（逐行标记）；skip_video_missing 表示视频数空缺率过高、
    属于数据源问题，已由全局提示覆盖，不再逐行标记。
    """
    f, flags = cfg["复核信号"], []
    rule = f["高互动低完播"]
    if rec.get("互动率") is not None and rec.get("完播率") is not None:
        if rec["互动率"] > rule["互动率下限"] and rec["完播率"] < rule["完播率上限"]:
            flags.append(f"高互动低完播（互动率 {fmt_pct(rec['互动率'])} / 完播率 {fmt_pct(rec['完播率'])}）")
    if f["粉丝负增长"] and rec.get("粉丝增长率") is not None and rec["粉丝增长率"] < 0:
        flags.append(f"粉丝负增长（{rec['粉丝增长率'] * 100:.2f}%）")
    if f["视频数缺失或为零"]:
        video_count = rec.get("近30日视频数")
        if video_count == 0:                      # 明确为 0（停更信号）始终标记
            flags.append("近30日视频数为 0")
        elif video_count is None and not skip_video_missing:
            flags.append("近30日视频数缺失")
    times = f["互动率超批内中位数倍数"]
    if median_rate and rec.get("互动率") is not None and rec["互动率"] > median_rate * times:
        flags.append(f"互动率超批内中位数 {times} 倍（{fmt_pct(rec['互动率'])} vs 中位 {fmt_pct(median_rate)}）")
    for field in f["关键指标单元格缺失"]:
        if field in cols_present and rec.get(field) is None:
            flags.append(f"{field}缺失，无法核验该门槛")
    return flags


def tier_of(rec: dict, cfg: dict) -> str:
    return cfg["梯队规则"]["等级到梯队"].get(rec["等级"], TIER_ORDER[-1])


def interaction_band(rec: dict, cfg: dict) -> str:
    """互动率档位：仅展示用，不参与梯队划分。"""
    rate = rec.get("互动率")
    if rate is None:
        return "-"
    b = cfg["梯队规则"]["互动率区间"]
    if rate * 100 > b["第一梯队"]:
        return ">8%"
    if rate * 100 > b["第二梯队"]:
        return "5-8%"
    return "3-5%"


# ════════════════ 主流程 ════════════════

def build_reason(rec: dict, cfg: dict, bonus_notes: list[str], flags: list[str]) -> str:
    dims = "、".join(f"{k} {v:.1f}" if v is not None else f"{k} -" for k, v in rec["_dims"].items())
    text = (f"五维总分 {rec['五维总分']:.1f}（{rec['等级']} 级 → {rec['梯队']}）；"
            f"分维：{dims}；加分 {rec['加分项']} 分："
            + ("；".join(bonus_notes) if bonus_notes else "无")
            + f"；互动率 {fmt_pct(rec['互动率'])}、女性占比 "
            + ("-" if rec.get("女性占比") is None else f"{rec['女性占比'] * 100:.0f}%")
            + f"、均播 {fmt_int(rec.get('均播'))}、报价 {fmt_int(rec.get('报价'))} 元、CPM "
            + ("-" if rec.get("CPM") is None else f"{rec['CPM']:.1f}")
            + "。")
    if flags:
        text += f" 复核提示：{'；'.join(flags)}。"
    return text


def print_top(passed: list[dict], top_n: int = 10) -> None:
    headers = ["排名", "昵称", "梯队", "等级", "互动率", "女性占比", "均播", "报价", "CPM", "五维分", "复核标记"]
    widths = [4, 22, 18, 4, 8, 9, 11, 9, 7, 7, 40]
    line = "─" * (sum(widths) + len(widths) - 1)
    print("\n" + line)
    print(" ".join(pad(h, w) for h, w in zip(headers, widths)))
    print(line)
    for i, rec in enumerate(passed[:top_n], 1):
        cells = [i, rec["昵称"], rec["梯队"], rec["等级"], fmt_pct(rec["互动率"]),
                 "-" if rec.get("女性占比") is None else f"{rec['女性占比'] * 100:.0f}%",
                 fmt_int(rec.get("均播")), fmt_int(rec.get("报价")),
                 "-" if rec.get("CPM") is None else f"{rec['CPM']:.1f}",
                 f"{rec['五维总分']:.1f}",
                 "建议人工复核" if rec["复核标记"] else ""]
        print(" ".join(pad(c, w) for c, w in zip(cells, widths)))
    print(line)


def export_excel(passed: list[dict], excluded: list[dict], stats: list[tuple],
                 output: Path, cfg: dict) -> None:
    def rows(source: list[dict], with_reason: bool) -> list[dict]:
        out = []
        for i, rec in enumerate(source, 1):
            row = {"排名": i, "昵称": rec["昵称"], "梯队": rec.get("梯队", "-"), "等级": rec["等级"],
                   "五维总分": rec["五维总分"], "加分项": rec.get("加分项"),
                   "互动率": fmt_pct(rec.get("互动率")), "互动率档位": interaction_band(rec, cfg),
                   "女性占比": "-" if rec.get("女性占比") is None else f"{rec['女性占比'] * 100:.0f}%",
                   "均播": rec.get("均播"), "报价": rec.get("报价"),
                   "CPM": None if rec.get("CPM") is None else round(rec["CPM"], 2),
                   "粉丝量": rec.get("粉丝量"), "赛道": rec.get("赛道"), "内容形态": rec.get("内容形态"),
                   "粉丝增长率": "-" if rec.get("粉丝增长率") is None else f"{rec['粉丝增长率'] * 100:.2f}%",
                   "近30日视频数": rec.get("近30日视频数"),
                   "复核标记": "建议人工复核" if rec.get("复核标记") else "",
                   "复核原因": "；".join(rec.get("复核原因", [])),
                   "主页链接": rec.get("主页链接")}
            if with_reason:
                row["推荐理由"] = rec.get("推荐理由", "")
            else:
                row["排除原因"] = rec.get("排除原因", "")
            out.append(row)
        return out

    rec_cols = ["排名", "昵称", "梯队", "等级", "五维总分", "加分项", "互动率", "互动率档位", "女性占比",
                "均播", "报价", "CPM", "粉丝量", "赛道", "内容形态", "粉丝增长率", "近30日视频数",
                "复核标记", "复核原因", "推荐理由", "主页链接"]
    exc_cols = ["昵称", "排除原因", "等级", "五维总分", "粉丝量", "互动率", "均播", "报价", "CPM",
                "女性占比", "赛道", "粉丝增长率", "近30日视频数", "主页链接"]
    rec_df = pd.DataFrame(rows(passed, True), columns=rec_cols)
    if not len(rec_df):
        rec_df = pd.DataFrame(columns=rec_cols)
    exc_df = pd.DataFrame(rows(excluded, False), columns=exc_cols)
    if not len(exc_df):
        exc_df = pd.DataFrame(columns=exc_cols)
    output.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        rec_df.to_excel(writer, sheet_name="推荐清单", index=False)
        exc_df.to_excel(writer, sheet_name="排除清单", index=False)
        pd.DataFrame(stats, columns=["指标", "值"]).to_excel(writer, sheet_name="筛选统计", index=False)


def main() -> None:
    started = time.time()
    parser = argparse.ArgumentParser(description="达人筛选 skill v1.0：筛人 + 排序 + 导出")
    parser.add_argument("--input", required=True, help="输入 CSV/Excel 路径")
    parser.add_argument("--output", required=True, help="输出 Excel 路径")
    parser.add_argument("--config", default=str(CONFIG_PATH), help="配置文件路径（默认同目录 config.yaml）")
    parser.add_argument("--top", type=int, default=10, help="终端打印前 N 名（默认 10）")
    args = parser.parse_args()

    cfg = load_config(Path(args.config))
    print(f"[筛选] 配置版本 v{cfg['version']}｜输入：{args.input}")

    df = read_table(Path(args.input))
    cols, missing = pick_columns(df, cfg)
    for msg in missing:
        print(f"[筛选] ⚠ {msg}")

    records = build_records(df, cols, cfg)
    for rec in records:
        ensure_grade(rec, cfg)
        rec["加分项"], bonus_notes = bonus_points(rec, cfg)
        rec["_bonus_notes"] = bonus_notes
        rec["复核原因"] = []          # 先占位，取到批内中位数后再算
        rec["梯队"] = tier_of(rec, cfg)

    rates = [r["互动率"] for r in records if r.get("互动率") is not None]
    median_rate = pd.Series(rates).median() if rates else None
    cols_present = {field for field, col in cols.items() if col}
    video_missing = sum(1 for r in records if r.get("近30日视频数") is None)
    skip_video_missing = "近30日视频数" not in cols_present   # 整列不存在时由缺列提示覆盖
    if "近30日视频数" in cols_present and records:
        ratio = video_missing / len(records)
        if ratio > cfg["复核信号"]["视频数缺失比例上限"]:
            skip_video_missing = True
            print(f"[筛选] ⚠ 「近30日视频数」空缺率 {ratio:.0%}（数据源未提供），"
                  f"该复核信号仅作为全局提示，不逐行标记")
    for rec in records:
        rec["复核原因"] = review_flags(rec, cfg, median_rate, cols_present, skip_video_missing)
        rec["复核标记"] = bool(rec["复核原因"])
        rec["排除原因"] = "；".join(hard_gates(rec, cfg))

    passed = [r for r in records if not r["排除原因"]]
    excluded = [r for r in records if r["排除原因"]]
    passed.sort(key=lambda r: (TIER_ORDER.index(r["梯队"]), -r["加分项"], -r["五维总分"],
                               -(r.get("互动率") or 0)))
    for rec in passed:
        rec["推荐理由"] = build_reason(rec, cfg, rec["_bonus_notes"], rec["复核原因"])

    tier_counts = {tier: sum(1 for r in passed if r["梯队"] == tier) for tier in TIER_ORDER}
    stats = [
        ("配置版本", f"v{cfg['version']}"), ("输入文件", str(args.input)), ("原始达人数", len(records)),
        ("硬门槛通过", len(passed)), ("排除人数", len(excluded)),
        ("第一梯队", tier_counts[TIER_ORDER[0]]), ("第二梯队", tier_counts[TIER_ORDER[1]]),
        ("第三梯队", tier_counts[TIER_ORDER[2]]),
        ("建议人工复核", sum(1 for r in passed if r["复核标记"])),
        ("互动率门槛", cfg["硬门槛"]["互动率下限"]), ("报价区间", str(cfg["硬门槛"]["报价区间"])),
        ("CPM 上限", cfg["硬门槛"]["CPM上限"]), ("梯队依据", cfg["梯队规则"]["依据"]),
        ("批内互动率中位数", "-" if median_rate is None else round(float(median_rate), 4)),
    ]
    export_excel(passed, excluded, stats, Path(args.output), cfg)

    print(f"[筛选] 原始 {len(records)} 人 → 通过硬门槛 {len(passed)} 人 → 排除 {len(excluded)} 人")
    print(f"[筛选] 梯队（由五维等级决定）：第一 {tier_counts[TIER_ORDER[0]]} / "
          f"第二 {tier_counts[TIER_ORDER[1]]} / 第三 {tier_counts[TIER_ORDER[2]]}")
    if median_rate is not None:
        print(f"[筛选] 批内互动率中位数 {median_rate * 100:.2f}%（用于复核信号判定）")
    print_top(passed, args.top)

    flagged = [r for r in passed if r["复核标记"]]
    if flagged:
        print(f"[筛选] 建议人工复核 {len(flagged)} 人（仅标记，未排除、未降权）：")
        for rec in flagged:
            print(f"        · {rec['昵称']}（{rec['梯队']}）：{'；'.join(rec['复核原因'])}")

    elapsed = time.time() - started
    print(f"[筛选] 产物：{Path(args.output).resolve()}")
    print(f"[筛选] 耗时 {elapsed:.2f} 秒" + ("" if elapsed <= 10 else "（超过 10 秒，请检查输入文件大小）"))


if __name__ == "__main__":
    main()
