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

口径（2026-08-04 版）：只跟踪**各家当前主推的产品线**，不做强/中/弱档位划分。
白名单里每条 = 厂商的一条真实产品线（如 OpenAI 的 GPT 主线、GPT mini 线），
每条取该线最新版本；每家有几条主推线就几条（1~3 条，不强凑数量一致）。
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

# ---------- 主推产品线白名单 ----------
# 口径：只跟踪各家**当前主推的产品线**，不做"强/中/弱三档"划分。
# 三档口径已废弃——它会强行把各家塞进三个格子，导致同一家三档跨代（旗舰
# gpt-5.5 配主力 gpt-5.4-mini）、甚至"主力"比"旗舰"更新更贵（Google
# gemini-3.5-flash 出$9 vs gemini-3.1-pro 出$12），定位名实不符。
#
# 白名单每条 = 厂商的一条真实产品线，键是**稳定的产品线名**（厂商自己的命名，
# 非主观分级），值是匹配该线的正则。正则第 1 个捕获组提版本号，每天在该线
# 命中的型号里取版本最新的一个——所以 gpt-5.5→gpt-5.6 换代自动跟上，
# 趋势线按"厂商+产品线"连续不断。
#
# 每家跟几条按实际主推情况定，不强求数量一致：
#  - Anthropic 3 条（Opus 主线 + Sonnet 量产线 + Fable 长文本线）
#  - OpenAI / Google / Qwen 各 2 条（主线 + 廉价量产线）
#  - 智谱 / Kimi / MiniMax / DeepSeek 各 1 条（只有单一主推线）
# 增删产品线：直接改这个字典即可，visualize.py 会自动跟着变。
#
# 已刻意排除的型号（避免污染横向对比）：
#  - 增强版：gpt-*-pro（出$180）、claude-*-fast、glm-*-turbo、*-max-thinking
#  - 特供版：*-codex / *-chat / *-image / *-audio / *-search / *-vl / *-v（视觉）
#  - 开源权重系列：gpt-oss-*、gemma-*、qwen3-235b 等按参数量命名的开放模型
#  - 极廉价小杯：gpt-*-nano、gemini-*-flash-lite、glm-*-flash（非主推，噪声大）
#  - 豆包(Doubao)：数据源对其真旗舰无报价，整家不跟踪
_V = r"(\d+(?:\.\d+)?)"          # 版本号捕获组
WATCHLIST = {
    "OpenAI": {
        "GPT": rf"^gpt-{_V}$",
        "GPT mini": rf"^gpt-{_V}-mini$",
    },
    "Anthropic": {
        "Claude Opus": rf"^claude-opus-{_V}$",
        "Claude Sonnet": rf"^claude-sonnet-{_V}$",
        "Claude Fable": rf"^claude-fable-{_V}$",
    },
    "Google": {
        "Gemini Pro": rf"^gemini-{_V}-pro(?:-preview)?$",
        "Gemini Flash": rf"^gemini-{_V}-flash(?:-preview)?$",
    },
    "智谱 (Z.ai)": {
        "GLM": rf"^glm-{_V}$",
    },
    "Qwen": {
        "Qwen Max": rf"^qwen{_V}-max(?:-preview)?$",
        "Qwen Plus": rf"^qwen{_V}-plus(?:-[\d-]+)?$",
    },
    "Kimi (Moonshot)": {
        "Kimi K": rf"^kimi-k{_V}$",
    },
    "MiniMax": {
        "MiniMax M": rf"^minimax-m{_V}$",
    },
    "DeepSeek": {
        "DeepSeek V": rf"^deepseek-v{_V}(?:-pro)?$",
    },
}

# 跟踪厂商顺序（= 白名单书写顺序）
TRACKED_PROVIDERS = list(WATCHLIST.keys())

# 产品线展示顺序：厂商内按白名单书写顺序（主线在前）
SERIES_ORDER = {p: {s: i for i, s in enumerate(WATCHLIST[p])} for p in WATCHLIST}

# 取价来源优先级：同一型号在多个渠道命名空间下都存在时，优先取国际站
# OpenRouter 报价（与项目"美元国际站统一口径"一致），保证选中项稳定、可比。
# 仅认第一方命名空间（model_id 前缀在 PROVIDER_DISPLAY 里），排除 bedrock/
# fireworks/azure 等转售渠道的畸高/畸低价。
def _source_rank(src):
    s = src or ""
    if "openrouter" in s:
        return 2
    return 1

# Excel 列定义（「定位」已改为「产品线」——不再是主观分级，而是厂商自己的产品线名）
EXCEL_COLUMNS = [
    "抓取日期",
    "厂商",
    "产品线",
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


def select_watchlist_models(models):
    """
    按 WATCHLIST 规则，为每家每条主推产品线动态挑出当前型号。
    规则：每条线在命中该线正则的有价型号里，按 (版本号, 来源优先级, 非preview, model_id)
    降序取第一个 —— 版本最新优先；同版本优先 OpenRouter 国际站报价；正式版优先于
    preview；再退化到 model_id 字典序保证确定性。
    仅认第一方命名空间（前缀在 PROVIDER_DISPLAY），排除 bedrock/fireworks 等转售渠道。
    返回 {model_id: (厂商, 产品线)}。
    """
    # 收集候选：provider -> series -> [(version, source_rank, is_preview, model_id), ...]
    cand = {p: {s: [] for s in WATCHLIST[p]} for p in WATCHLIST}

    for model_id, m in models.items():
        ns = model_id.split("/")[0]
        provider = PROVIDER_DISPLAY.get(ns, None)
        if provider not in WATCHLIST:
            continue  # 只认第一方命名空间

        pricing = m.get("pricing", {}) or {}
        out_price = _to_float(pricing.get("output_per_million"))
        if out_price <= 0:  # 无价型号不参与
            continue

        short = model_id.split("/")[-1].lower()
        sources = m.get("sources", {}) or {}
        src = ",".join(sources.keys()) if sources else ""
        is_preview = ("preview" in short)
        for series, pattern in WATCHLIST[provider].items():
            mt = re.match(pattern, short, re.IGNORECASE)
            if not mt:
                continue
            ver = _parse_version(mt.group(1))
            cand[provider][series].append((ver, _source_rank(src), is_preview, model_id))

    # 每条产品线取最优候选
    result = {}
    for provider, series_map in cand.items():
        for series, items in series_map.items():
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
            result[items[0][3]] = (provider, series)
    return result


def build_watchlist_rows(models, today):
    """按白名单规则动态挑出各家主推产品线型号，构建 Excel 行"""
    model_to_series = select_watchlist_models(models)
    rows = []
    for model_id, (provider, series) in model_to_series.items():
        m = models[model_id]
        pricing = m.get("pricing", {}) or {}
        in_price = _to_float(pricing.get("input_per_million"))
        out_price = _to_float(pricing.get("output_per_million"))
        is_free = "是" if (in_price == 0 and out_price == 0) else "否"

        rows.append({
            "抓取日期": today,
            "厂商": provider,
            "产品线": series,
            "模型ID": model_id,
            "模型名称": m.get("display_name", model_id),
            "输入价格(美元/百万token)": round(in_price, 6),
            "输出价格(美元/百万token)": round(out_price, 6),
            "上下文长度": m.get("context_length"),
            "是否免费": is_free,
            "数据源": "tokenpricing",
        })

    # 检查白名单里每条产品线是否都命中（缺失只告警不报错——可能该线尚未发布/暂无报价）
    for provider in TRACKED_PROVIDERS:
        expected = set(WATCHLIST[provider].keys())
        found = {s for _, (p, s) in model_to_series.items() if p == provider}
        missing = expected - found
        if missing:
            log(f"⚠ {provider} 缺少产品线 {'/'.join(missing)}（该线正则本次无有价命中），请检查 WATCHLIST 正则")

    # 排序：按 TRACKED_PROVIDERS 顺序，厂商内按白名单书写顺序
    rows.sort(key=lambda r: (TRACKED_PROVIDERS.index(r["厂商"]),
                             SERIES_ORDER.get(r["厂商"], {}).get(r["产品线"], 9)))
    return rows


def save_watchlist_to_excel(rows, today):
    """主推产品线数据写入 Excel（同日覆盖）"""
    new_df = pd.DataFrame(rows, columns=EXCEL_COLUMNS)

    if os.path.exists(EXCEL_PATH):
        try:
            existing = pd.read_excel(EXCEL_PATH, sheet_name="价格数据")
            if "产品线" not in existing.columns:
                # 旧口径（含"定位"=旗舰/主力/轻量 或更早的旗舰/次旗舰）与新口径不可比，整表重置
                log("检测到旧口径数据结构（无「产品线」列），重置为新结构。历史可用 backfill_history.py --rebuild 按新口径重算。")
                combined = new_df
            else:
                existing = existing[existing["抓取日期"] != today]
                combined = pd.concat([existing, new_df], ignore_index=True)
        except Exception as e:
            log(f"读取已有 Excel 失败，将新建: {e}")
            combined = new_df
    else:
        combined = new_df

    combined = combined.sort_values(["抓取日期", "厂商", "产品线"]).reset_index(drop=True)
    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
        combined.to_excel(writer, sheet_name="价格数据", index=False)
    log(f"主推产品线数据已写入 {EXCEL_PATH}：本次 {len(rows)} 行，累计 {len(combined)} 行")


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

    # 2) 按白名单规则动态挑出各家主推产品线型号，写入 Excel（供可视化）
    watch_rows = build_watchlist_rows(models, today)
    if not watch_rows:
        log("未识别到任何主推产品线模型，退出")
        sys.exit(1)

    counts = Counter(r["厂商"] for r in watch_rows)
    summary = " / ".join(f"{k}: {v}" for k, v in counts.items())
    log(f"主推产品线识别 {len(watch_rows)} 个型号（{len(TRACKED_PROVIDERS)} 家厂商）— {summary}")

    save_watchlist_to_excel(watch_rows, today)

    log(f"完成：全量 {total} 模型 -> DB；主推产品线 {len(watch_rows)} -> Excel")
    log("=" * 60)


if __name__ == "__main__":
    main()
