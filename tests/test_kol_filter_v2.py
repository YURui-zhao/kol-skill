"""kol-skill V2 确定性与双核心规则回归测试（纯函数 + 隔离临时文件，无网络）。"""
import json
import sys
from pathlib import Path

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import kol_filter as kf


@pytest.fixture()
def cfg():
    return kf.load_config()


def base_rec(**overrides):
    rec = {
        "昵称": "样板达人", "星图id": "9001", "粉丝量": 500000, "互动率": 0.02,
        "完播率": 0.05, "女性占比": 0.65, "均播": 200000, "报价": 3800,
        "赛道": "艺术文化", "标题": "", "风险标签": "", "近30日视频数": 6,
        "星图指数": 75, "种草指数": 70, "粉丝增长率": 0.05,
        "预期CPM文本": "短视频1-20s:,短视频20-60s:20.0短视频60s+:30.0",
        "外部五维总分": None, "外部等级": None,
    }
    rec.update(overrides)
    return rec


def processed(cfg, **overrides):
    rec = base_rec(**overrides)
    rec["复核原因"] = []
    kf.ensure_grade(rec, cfg)
    rec["内容匹配分"] = kf.content_match_score(rec, cfg)
    return rec


# ---------- 1. 外部总分与等级来源一致 ----------

def test_external_total_drives_grade(cfg):
    rec = processed(cfg, 外部五维总分=90)          # 内部辅助分约 60-70，外部 90 → 必须 S
    assert rec["等级"] == "S" and rec["等级_依据"] == "外部总分"


def test_external_grade_conflict_becomes_review(cfg):
    rec = base_rec(外部五维总分=90, 外部等级="C")
    rec["复核原因"] = []
    kf.ensure_grade(rec, cfg)
    assert any("不一致" in reason for reason in rec["复核原因"])


def test_internal_fallback_marked_as_auxiliary(cfg):
    rec = processed(cfg)
    assert rec["五维总分_来源"] == "内部辅助分" and rec["等级_依据"] == "内部辅助分"


# ---------- 2. 门槛字段缺失不静默通过 ----------

def test_missing_gate_fields_flagged_for_review(cfg):
    rec = base_rec(互动率=None, 报价=None)
    rec["复核原因"] = []
    kf.ensure_grade(rec, cfg)
    flags = kf.review_flags(rec, cfg, median_rate=0.02, cols_present={"互动率", "报价"})
    assert any("门槛数据不全" in f for f in flags)
    assert "互动率" in flags[-1] or any("报价" in f for f in flags)


def test_missing_gate_does_not_exclude_but_blocks_recommendation(cfg):
    rec = base_rec(互动率=None)
    rec["复核原因"] = []
    kf.ensure_grade(rec, cfg)
    rec["内容匹配分"] = kf.content_match_score(rec, cfg)
    assert kf.hard_exclusions(rec, cfg) == []          # 不排除
    assert any("门槛数据不全" in f for f in kf.review_flags(rec, cfg, 0.02, {"互动率"}))


# ---------- 3. 百分比单位解析 ----------

def test_ratio_unit_parsing_percent_string():
    ratio, unit = kf.to_ratio("3%")
    assert ratio == 0.03 and unit == "percent"


def test_ratio_unit_parsing_number_above_one_assumed_percent():
    ratio, unit = kf.to_ratio(3)
    assert ratio == 0.03 and unit == "percent-assumed"


def test_ratio_unit_parsing_decimal_assumed_ratio():
    ratio, unit = kf.to_ratio(0.03)
    assert ratio == 0.03 and unit == "ratio-assumed"


def test_low_engagement_now_review_not_excluded(cfg):
    rec = base_rec(互动率=0.012)                        # V1 会以 3% 硬门槛排除；V2 只复核
    rec["复核原因"] = []
    kf.ensure_grade(rec, cfg)
    rec["内容匹配分"] = kf.content_match_score(rec, cfg)
    assert kf.hard_exclusions(rec, cfg) == []
    assert any("互动率偏低" in f for f in kf.review_flags(rec, cfg, 0.02, set()))


# ---------- 4. 风险标签分级 ----------

def test_risk_tag_exclude_level(cfg):
    rec = base_rec(风险标签="疑似刷量")
    reasons = kf.hard_exclusions(rec, cfg)
    assert any("排除级风险标签" in r for r in reasons)


def test_risk_tag_review_level_not_excluded(cfg):
    rec = base_rec(风险标签="资料不全")
    assert kf.hard_exclusions(rec, cfg) == []
    assert kf.classify_risk_tag("资料不全", cfg) == "复核级"


def test_risk_tag_unknown_defaults_to_review(cfg):
    assert kf.classify_risk_tag("神秘新标签", cfg) == "复核"
    rec = base_rec(风险标签="神秘新标签")
    assert kf.hard_exclusions(rec, cfg) == []


def test_risk_tag_empty_values_ignored(cfg):
    assert kf.classify_risk_tag("无", cfg) == "无"
    assert kf.classify_risk_tag("", cfg) == "无"


# ---------- 5. 星图预期 CPM 分档解析与价格效率 ----------

def test_tier_cpm_parsing():
    tiers = kf.parse_tier_cpm("短视频1-20s:,短视频20-60s:20.5短视频60s+:31.2")
    assert tiers["植入"] == 20.5 and tiers["定制"] == 31.2


def test_price_efficiency_uses_plant_quote_first(cfg):
    rec = base_rec(报价=3800, 均播=200000)             # 重算 = 19.0（良好档）
    rec["复核原因"] = []
    kf.ensure_grade(rec, cfg)
    rec["内容匹配分"] = kf.content_match_score(rec, cfg)
    rec["CPM_价格效率"] = round(rec["报价"] / rec["均播"] * 1000, 2)
    rec["CPM口径"] = "植入报价/预期播放"
    rec["CPM档位"] = kf.cpm_bucket(rec["CPM_价格效率"], cfg)
    rec["二级内容证据"] = False
    tier, _ = kf.assign_tier(rec, cfg)
    assert rec["CPM档位"] == "良好" and rec["CPM_价格效率"] == 19.0
    assert tier == "P2-推荐建联"                       # 无二级证据 → 最高 P2


def test_build_records_computes_price_efficiency(cfg):
    df = pd.DataFrame([
        {"昵称": "样板达人", "星图id": "9001", "粉丝量": 500000, "互动率": 0.02, "完播率": 0.05,
         "均播": 200000, "报价": 3800, "赛道": "艺术文化", "近30日星图视频数": 5,
         "星图指数": 75, "种草指数": 70, "粉丝增长率": 0.05,
         "预期CPM": "短视频1-20s:,短视频20-60s:20.0短视频60s+:30.0"},
    ])
    cols, _ = kf.pick_columns(df, cfg)
    records, _ = kf.build_records(df, cols, cfg, {})
    rec = records[0]
    assert rec["CPM_价格效率"] == 19.0 and rec["CPM口径"] == "植入报价/预期播放"
    assert rec["CPM档位"] == "良好"
    assert rec["星图CPM_植入档"] == 20.0 and rec["星图CPM_定制档"] == 30.0


def test_cpm_bucket_boundaries_no_overlap(cfg):
    assert kf.cpm_bucket(15, cfg) == "高效"
    assert kf.cpm_bucket(15.01, cfg) == "良好"
    assert kf.cpm_bucket(20, cfg) == "良好"
    assert kf.cpm_bucket(30, cfg) == "可谈"
    assert kf.cpm_bucket(30.01, cfg) == "低效"


# ---------- 6. 内容标签与 P1 证据门槛 ----------

def test_p1_requires_verified_second_level_evidence(cfg):
    rec = base_rec()
    rec["复核原因"] = []
    kf.ensure_grade(rec, cfg)
    rec["内容匹配分"] = 90                              # 高匹配
    rec["CPM档位"] = "高效"
    rec["CPM_价格效率"] = 12
    rec["二级内容证据"] = False
    tier, reasons = kf.assign_tier(rec, cfg)
    assert tier == "P2-推荐建联"                        # 无证据不得 P1
    assert any("内容细分待核验" in r for r in reasons)


def test_p1_granted_with_verified_evidence(cfg):
    rec = base_rec()
    rec["复核原因"] = []
    kf.ensure_grade(rec, cfg)
    rec["内容匹配分"] = 90
    rec["CPM档位"] = "高效"
    rec["CPM_价格效率"] = 12
    rec["二级内容证据"] = True
    tier, _ = kf.assign_tier(rec, cfg)
    assert tier == "P1-优先建联"


def test_unverified_tag_file_not_treated_as_evidence(cfg, tmp_path):
    tags = pd.DataFrame([{"星图id": "9001", "一级方向": "历史人文", "二级细分": "历史知识",
                          "内容证据": "代表作标题", "标签来源": "自动推断", "核验状态": "待核验"}])
    tags_path = tmp_path / "tags.csv"
    tags.to_csv(tags_path, index=False, encoding="utf-8-sig")
    loaded = kf.load_content_tags(tags_path, cfg)
    assert loaded["9001"]["二级已核验"] is False        # 自动来源/未核验 → 不作为 P1 证据


def test_verified_tag_file_enables_evidence(cfg, tmp_path):
    tags = pd.DataFrame([{"星图id": "9001", "一级方向": "历史人文", "二级细分": "历史知识",
                          "内容证据": "代表作：王朝兴衰系列", "标签来源": "人工核验", "核验状态": "已核验"}])
    tags_path = tmp_path / "tags.csv"
    tags.to_csv(tags_path, index=False, encoding="utf-8-sig")
    loaded = kf.load_content_tags(tags_path, cfg)
    assert loaded["9001"]["二级已核验"] is True


# ---------- 7. 星图 ID 去重 ----------

def test_dedupe_by_star_id(cfg):
    df = pd.DataFrame([
        {"昵称": "样板达人", "星图id": "9001", "粉丝量": 500000, "互动率": 0.02, "完播率": 0.05,
         "均播": 200000, "报价": 3800, "赛道": "艺术文化", "近30日星图视频数": 5,
         "星图指数": 75, "种草指数": 70, "粉丝增长率": 0.05, "预期CPM": ""},
        {"昵称": "样板达人（重复行）", "星图id": "9001", "粉丝量": 100, "互动率": 0.01, "完播率": 0.01,
         "均播": 1000, "报价": 500, "赛道": "艺术文化", "近30日星图视频数": 0,
         "星图指数": 40, "种草指数": 40, "粉丝增长率": 0, "预期CPM": ""},
    ])
    cols, _ = kf.pick_columns(df, cfg)
    records, duplicates = kf.build_records(df, cols, cfg, {})
    assert len(records) == 1 and len(duplicates) == 1
    assert records[0]["昵称"] == "样板达人" and duplicates[0]["昵称"] == "样板达人（重复行）"


# ---------- 8. 存档（评分版本 + 输入指纹） ----------

def test_run_filter_archives_fingerprint_and_version(cfg, tmp_path):
    df = pd.DataFrame([
        {"昵称": "样板达人", "星图id": "9001", "粉丝量": 500000, "互动率": 0.02, "完播率": 0.05,
         "均播": 200000, "报价": 3800, "赛道": "艺术文化", "近30日星图视频数": 5,
         "星图指数": 75, "种草指数": 70, "粉丝增长率": 0.05, "预期CPM": "短视频20-60s:19.0"},
    ])
    src = tmp_path / "input.xlsx"
    df.to_excel(src, index=False)
    out = tmp_path / "out.xlsx"
    result = kf.run_filter(src, out, cfg, archive_dir=tmp_path / "archive")
    assert result["fingerprint"] == kf.file_sha256(src)
    archive = json.loads(sorted((tmp_path / "archive").glob("*.json"))[0].read_text(encoding="utf-8"))
    assert archive["scoring_version"] == cfg["评分版本"]
    assert archive["input_sha256"] == result["fingerprint"]
    assert archive["tier_counts"]


def test_run_filter_keeps_star_id_in_output(cfg, tmp_path):
    df = pd.DataFrame([
        {"昵称": "样板达人", "星图id": "9001", "粉丝量": 500000, "互动率": 0.02, "完播率": 0.05,
          "均播": 200000, "报价": 3800, "赛道": "艺术文化", "近30日星图视频数": 5,
         "星图指数": 75, "种草指数": 70, "粉丝增长率": 0.05, "预期CPM": "短视频20-60s:19.0"},
    ])
    src = tmp_path / "input.xlsx"
    df.to_excel(src, index=False)
    result = kf.run_filter(src, tmp_path / "out.xlsx", cfg, archive_dir=tmp_path / "archive")
    assert result["kept"][0]["星图id"] == "9001"
