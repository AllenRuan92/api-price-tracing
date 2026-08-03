# -*- coding: utf-8 -*-
"""
API 价格抓取脚本（tokenpricing 版，全量 + 自建历史）

数据源：https://github.com/Atena-IT/tokenpricing
  - 每 6 小时从 OpenRouter + LiteLLM 同步 3200+ 模型价格，归一化为 JSON 存 GitHub 仓库
  - 通过 raw.githubusercontent.com 拉取，零 key 零限流，几乎不会宕机
  - 价格单位直接是 per-million，无需量级判断

存储设计（每日 9:00 追加一次，自建历史，不依赖数据源的历史目录）：
  1. prices_full.db  (SQLite) — 全量 3200+ 模型每日快照，可长期累积百万行
  2. prices.xlsx     — 主推档位模型每日快照，供 visualize.py 直接读取展示

口径（2026-08 改版）：不再按"旗舰/次旗舰（相邻两代）"，改为跟踪每家
**当前主推的产品线档位**——旗舰 / 主力 / 轻量三档，每档各取该档最新版本。
每家有几档就展示几档（Kimi/MiniMax 只有旗舰一档，DeepSeek 只有旗舰+轻量）。
"""

import os
import sys
import re
import json
import sqlite3
import time
from datetime import datetime
from collections import Counter

import requests
import pandas as pd

# ---------- 配置 ----------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(SCRIPT_DIR, "prices_full.db")
EXCEL_PATH = os.path.join(SCRIPT_DIR, "prices.xlsx")
LOG_PATH = os.path.join(SCRIPT_DIR, "fetch_log.txt")

# tokenpricing 当前价格数据库（GitHub raw，CDN 加速，不限流）
PRICES_URL = "https://raw.githubusercontent.com/Atena-IT/tokenpricing/main/database/current/prices.json"


# 厂商显示名映射（tokenpricing 里智谱前缀是 z-ai/）
PROVIDER_DISPLAY = {
    "openai": "OpenAI",
    "anthropic": "Anthropic",
    "google": "Google",
    "z-ai": "智谱 (Z.ai)",
    "minimax": "MiniMax",
    "qwen": "Qwen",
    "moonshotai": "Kimi (Moonshot)",
    "moonshot": "Kimi (Moonshot)",
    "deepseek": "DeepSeek",
}

# ---------- 主推档位动态识别规则 ----------
# 口径：跟踪每家「当前主推产品线」的价格，按档位划分——
#   旗舰(顶配大杯) / 主力(中杯) / 轻量(小杯)。
# 不写死型号，而是为「每家每档」定义一条正则；每天在该档命中的型号里
# 按版本号取最新一个。这样型号换代（gpt-5.5-mini → gpt-5.6-mini）自动跟上，
# 趋势线按「厂商+档位」连续不断。
#
# 正则第 1 个捕获组提取版本号（用于同档内选最新），匹配 model_id 去掉厂商
# 前缀后的「短名」（小写）。每家有几档就配几档，缺档不报错只告警。
#
# 说明：
#  - OpenAI 旗舰用标准版 gpt-5.5（非 pro 增强版）：pro 出$180 价格畸高会拉变形
#    横向对比，标准版才与 opus 同档可比。mini=主力、nano=轻量。
#  - Kimi(K 系列)、MiniMax(M 系列) 只有单一产品线，仅旗舰一档。
#  - DeepSeek 用 vX-pro=旗舰 / vX-flash=轻量（无中杯主力）。
#  - 豆包(Doubao) 真旗舰在本数据源无价，整系列只有 lite/mini，不跟踪。
_V = r"(\d+(?:\.\d+)?)"          # 版本号捕获组
MODEL_TIERS = {
    "OpenAI": {
        "旗舰": rf"^gpt-{_V}$",
        "主力": rf"^gpt-{_V}-mini$",
        "轻量": rf"^gpt-{_V}-nano$",
    },
    "Anthropic": {
        "旗舰": rf"^claude-opus-{_V}$",
        "主力": rf"^claude-sonnet-{_V}$",
        "轻量": rf"^claude-haiku-{_V}$",
    },
    "Google": {
        "旗舰": rf"^gemini-{_V}-pro(?:-preview)?$",
        "主力": rf"^gemini-{_V}-flash(?:-preview)?$",
        "轻量": rf"^gemini-{_V}-flash-lite(?:-preview)?$",
    },
    "智谱 (Z.ai)": {
        "旗舰": rf"^glm-{_V}$",
        "主力": rf"^glm-{_V}-air$",
        "轻量": rf"^glm-{_V}-flash$",
    },
    "Qwen": {
        "旗舰": rf"^qwen{_V}-max(?:-preview)?$",
        "主力": rf"^qwen{_V}-plus(?:-[\d-]+)?$",
        "轻量": rf"^qwen{_V}-flash(?:-[\d-]+)?$",
    },
    "Kimi (Moonshot)": {
        "旗舰": rf"^kimi-k{_V}$",
    },
    "MiniMax": {
        "旗舰": rf"^minimax-m{_V}$",
    },
    "DeepSeek": {
        "旗舰": rf"^deepseek-v{_V}-pro$",
        "轻量": rf"^deepseek-v{_V}-flash$",
    },
}

# 档位展示顺序（旗舰在前）
TIER_ORDER = ["旗舰", "主力", "轻量"]

# 跟踪厂商顺序
TRACKED_PROVIDERS = list(MODEL_TIERS.keys())

# 取价来源优先级：同一型号在多个渠道命名空间下都存在时，优先取国际站
# OpenRouter 报价（与项目"美元国际站统一口径"一致），保证选中项稳定、可比。
# 仅认第一方命名空间（model_id 前缀在 PROVIDER_DISPLAY 里），排除 bedrock/
# fireworks/azure 等转售渠道的畸高/畸低价。
def _source_rank(src):
    s = src or ""
    if "openrouter" in s:
        return 2
    return 1

# 旗舰 Excel 列定义（保持与旧版兼容，可视化脚本无需改）
EXCEL_COLUMNS = [
    "抓取日期",
    "厂商",
    "定位",
    "模型ID",
    "模型名称",
    "输入价格(美元/百万token)",
    "输出价格(美元/百万token)",
    "上下文长度",
    "是否免费",
    "数据源",
]

REQUEST_TIMEOUT = 60
MAX_RETRIES = 3
RETRY_DELAY = 5


def log(message):
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}"
    print(line)
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def fetch_full_prices():
    """从 GitHub raw 拉取 tokenpricing 全量价格 JSON"""
    headers = {"User-Agent": "API-Price-Tracer/1.0", "Accept": "application/json"}
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            log(f"拉取 tokenpricing 全量价格（第 {attempt} 次）...")
            r = requests.get(PRICES_URL, headers=headers, timeout=REQUEST_TIMEOUT)
            r.raise_for_status()
            data = r.json()
            models = data.get("models", {})
            generated_at = data.get("generated_at", "")
            log(f"成功获取 {len(models)} 个模型（generated_at={generated_at}）")
            return models, generated_at
        except Exception as e:
            log(f"拉取失败: {e}")
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_DELAY)
    return None, ""


def _to_float(val, default=0.0):
    try:
        if val is None:
            return default
        return float(val)
    except (TypeError, ValueError):
        return default


def _to_int(val, default=None):
    try:
        if val is None:
            return default
        return int(val)
    except (TypeError, ValueError):
        return default


def _to_bool(val, default=False):
    if isinstance(val, bool):
        return val
    if val is None:
        return default
    return bool(val)


def init_db(conn):
    """初始化 SQLite 表（全量历史数据）"""
    conn.execute("""
        CREATE TABLE IF NOT EXISTS price_snapshots (
            抓取日期 TEXT NOT NULL,
            厂商 TEXT,
            模型ID TEXT NOT NULL,
            模型名称 TEXT,
            输入价格_per_million REAL,
            输出价格_per_million REAL,
            缓存读价格_per_million REAL,
            缓存创建价格_per_million REAL,
            上下文长度 INTEGER,
            最大输出token INTEGER,
            模型类型 TEXT,
            分类 TEXT,
            支持视觉 INTEGER,
            支持函数调用 INTEGER,
            支持流式 INTEGER,
            数据来源 TEXT,
            数据生成时间 TEXT,
            PRIMARY KEY (抓取日期, 模型ID)
        )
    """)
    conn.execute("""
        CREATE INDEX IF NOT EXISTS idx_snapshots_date ON price_snapshots(抓取日期)
    """)
    conn.execute("""
        CREATE INDEX IF NOT EXISTS idx_snapshots_model ON price_snapshots(模型ID)
    """)
    conn.commit()


def save_full_to_db(models, generated_at, today):
    """全量写入 SQLite（同日重复抓取则覆盖当日）"""
    rows = []
    for model_id, m in models.items():
        pricing = m.get("pricing", {}) or {}
        sources = m.get("sources", {}) or {}
        provider_raw = m.get("provider", "")
        # 厂商显示名：优先用映射，否则用原始 provider
        provider_display = PROVIDER_DISPLAY.get(provider_raw, provider_raw or "未知")

        rows.append((
            today,
            provider_display,
            model_id,
            m.get("display_name", model_id),
            _to_float(pricing.get("input_per_million")),
            _to_float(pricing.get("output_per_million")),
            _to_float(pricing.get("cache_read_per_million")) if pricing.get("cache_read_per_million") is not None else None,
            _to_float(pricing.get("cache_creation_per_million")) if pricing.get("cache_creation_per_million") is not None else None,
            _to_int(m.get("context_window")),
            _to_int(m.get("max_output_tokens")),
            m.get("model_type", ""),
            m.get("category", ""),
            1 if _to_bool(m.get("supports_vision")) else 0,
            1 if _to_bool(m.get("supports_function_calling")) else 0,
            1 if _to_bool(m.get("supports_streaming")) else 0,
            ",".join(sources.keys()) if sources else "",
            generated_at,
        ))

    conn = sqlite3.connect(DB_PATH)
    try:
        init_db(conn)
        # 同日覆盖：先删当日已有数据
        conn.execute("DELETE FROM price_snapshots WHERE 抓取日期 = ?", (today,))
        conn.executemany("""
            INSERT INTO price_snapshots (
                抓取日期, 厂商, 模型ID, 模型名称,
                输入价格_per_million, 输出价格_per_million,
                缓存读价格_per_million, 缓存创建价格_per_million,
                上下文长度, 最大输出token, 模型类型, 分类,
                支持视觉, 支持函数调用, 支持流式,
                数据来源, 数据生成时间
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, rows)
        conn.commit()
        log(f"全量数据已写入 {DB_PATH}：{len(rows)} 行（当日 {today}）")
    finally:
        conn.close()
    return len(rows)


def _parse_version(verstr):
    """把版本号字符串解析为可比较的元组。'5.5'->(5,5)，'4'->(4,0)，'4-1'->(4,1)"""
    s = verstr.replace("-", ".")
    parts = s.split(".")
    try:
        major = int(parts[0])
        minor = int(parts[1]) if len(parts) > 1 else 0
        return (major, minor)
    except (ValueError, IndexError):
        return (0, 0)


def select_tier_models(models):
    """
    按 MODEL_TIERS 规则，为每家每档动态挑出当前主推型号。
    规则：每档在命中该档正则的有价型号里，按 (版本号, 来源优先级, model_id) 降序取第一个
    —— 版本最新优先；同版本优先 OpenRouter 国际站报价；再退化到 model_id 字典序保证确定性。
    仅认第一方命名空间（前缀在 PROVIDER_DISPLAY），排除 bedrock/fireworks 等转售渠道。
    返回 {model_id: (厂商, 档位)}。
    """
    # 收集候选：provider -> tier -> [(version, source_rank, model_id), ...]
    cand = {p: {t: [] for t in MODEL_TIERS[p]} for p in MODEL_TIERS}

    for model_id, m in models.items():
        ns = model_id.split("/")[0]
        provider = PROVIDER_DISPLAY.get(ns, None)
        if provider not in MODEL_TIERS:
            continue  # 只认第一方命名空间

        pricing = m.get("pricing", {}) or {}
        out_price = _to_float(pricing.get("output_per_million"))
        if out_price <= 0:  # 无价型号不参与
            continue

        short = model_id.split("/")[-1].lower()
        sources = m.get("sources", {}) or {}
        src = ",".join(sources.keys()) if sources else ""
        is_preview = ("preview" in short)
        for tier, pattern in MODEL_TIERS[provider].items():
            mt = re.match(pattern, short, re.IGNORECASE)
            if not mt:
                continue
            ver = _parse_version(mt.group(1))
            cand[provider][tier].append((ver, _source_rank(src), is_preview, model_id))

    # 每档取最优候选
    result = {}
    for provider, tiers in cand.items():
        for tier, items in tiers.items():
            if not items:
                continue
            # 多级稳定排序（按优先级从低到高依次排，利用 sort 稳定性）：
            #   4) model_id 字典序小者优先（正式版通常短于带后缀的变体）
            items.sort(key=lambda x: x[3])
            #   3) 正式版优先于 -preview
            items.sort(key=lambda x: x[2])
            #   2) 来源优先级高者优先（OpenRouter 国际站）
            items.sort(key=lambda x: x[1], reverse=True)
            #   1) 版本号最新者优先
            items.sort(key=lambda x: x[0], reverse=True)
            result[items[0][3]] = (provider, tier)
    return result


def build_flagship_rows(models, today):
    """按档位规则动态挑出每家主推档位型号，构建 Excel 行"""
    model_to_tier = select_tier_models(models)
    rows = []
    for model_id, (provider, tier) in model_to_tier.items():
        m = models[model_id]
        pricing = m.get("pricing", {}) or {}
        in_price = _to_float(pricing.get("input_per_million"))
        out_price = _to_float(pricing.get("output_per_million"))
        is_free = "是" if (in_price == 0 and out_price == 0) else "否"

        rows.append({
            "抓取日期": today,
            "厂商": provider,
            "定位": tier,
            "模型ID": model_id,
            "模型名称": m.get("display_name", model_id),
            "输入价格(美元/百万token)": round(in_price, 6),
            "输出价格(美元/百万token)": round(out_price, 6),
            "上下文长度": m.get("context_length"),
            "是否免费": is_free,
            "数据源": "tokenpricing",
        })

    # 检查每家配置的档位是否都命中（缺档只告警，不报错——有些家本就没那档）
    for provider in TRACKED_PROVIDERS:
        expected = set(MODEL_TIERS[provider].keys())
        found = {tier for _, (p, tier) in model_to_tier.items() if p == provider}
        missing = expected - found
        if missing:
            log(f"⚠ {provider} 缺少档位 {'/'.join(missing)}（该档正则本次无有价命中），请检查 MODEL_TIERS 正则")

    # 排序：按 TRACKED_PROVIDERS 顺序、档位按 TIER_ORDER
    tier_rank = {t: i for i, t in enumerate(TIER_ORDER)}
    rows.sort(key=lambda r: (TRACKED_PROVIDERS.index(r["厂商"]), tier_rank.get(r["定位"], 9)))
    return rows


def save_flagship_to_excel(rows, today):
    """旗舰模型写入 Excel（同日覆盖）"""
    new_df = pd.DataFrame(rows, columns=EXCEL_COLUMNS)

    if os.path.exists(EXCEL_PATH):
        try:
            existing = pd.read_excel(EXCEL_PATH, sheet_name="价格数据")
            if "定位" not in existing.columns:
                log("检测到旧版数据结构，重置为新结构。")
                combined = new_df
            else:
                existing = existing[existing["抓取日期"] != today]
                combined = pd.concat([existing, new_df], ignore_index=True)
        except Exception as e:
            log(f"读取已有 Excel 失败，将新建: {e}")
            combined = new_df
    else:
        combined = new_df

    combined = combined.sort_values(["抓取日期", "厂商", "定位"]).reset_index(drop=True)
    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
        combined.to_excel(writer, sheet_name="价格数据", index=False)
    log(f"旗舰数据已写入 {EXCEL_PATH}：本次 {len(rows)} 行，累计 {len(combined)} 行")


def main():
    log("=" * 60)
    log("开始抓取 API 价格（tokenpricing 全量 + 自建历史）")

    models, generated_at = fetch_full_prices()
    if not models:
        log("未获取到模型数据，退出")
        sys.exit(1)

    today = datetime.now().strftime("%Y-%m-%d")  # GitHub Actions 在 UTC 1:17 跑，对应北京时间 9:17，日期一致

    # 1) 全量写入 SQLite
    total = save_full_to_db(models, generated_at, today)

    # 2) 按系列规则动态挑出旗舰写入 Excel（供可视化）
    flagship_rows = build_flagship_rows(models, today)
    if not flagship_rows:
        log("未识别到任何旗舰模型，退出")
        sys.exit(1)

    counts = Counter(r["厂商"] for r in flagship_rows)
    summary = " / ".join(f"{k}: {v}" for k, v in counts.items())
    log(f"档位动态识别 {len(flagship_rows)} 个（{len(TRACKED_PROVIDERS)} 家 × 旗舰/主力/轻量）— {summary}")

    save_flagship_to_excel(flagship_rows, today)

    log(f"完成：全量 {total} 模型 -> DB；旗舰 {len(flagship_rows)} -> Excel")
    log("=" * 60)


if __name__ == "__main__":
    main()
