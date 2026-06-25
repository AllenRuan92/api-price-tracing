# -*- coding: utf-8 -*-
"""
API 价格抓取脚本（tokenpricing 版，全量 + 自建历史）

数据源：https://github.com/Atena-IT/tokenpricing
  - 每 6 小时从 OpenRouter + LiteLLM 同步 3200+ 模型价格，归一化为 JSON 存 GitHub 仓库
  - 通过 raw.githubusercontent.com 拉取，零 key 零限流，几乎不会宕机
  - 价格单位直接是 per-million，无需量级判断

存储设计（每日 9:00 追加一次，自建历史，不依赖数据源的历史目录）：
  1. prices_full.db  (SQLite) — 全量 3200+ 模型每日快照，可长期累积百万行
  2. prices.xlsx     — 12 个旗舰/次旗舰模型每日快照，供 visualize.py 直接读取展示

跟踪厂商：OpenAI / Anthropic / Google / 智谱 / MiniMax / Qwen
每家旗舰 + 次旗舰各 1 个，共 12 个模型（仅用于前台仪表盘展示）。
"""

import os
import sys
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
    "bytedance": "豆包 (Doubao)",
    "bytedance-seed": "豆包 (Doubao)",
    "doubao": "豆包 (Doubao)",
}

# 12 个旗舰/次旗舰白名单（仅用于 prices.xlsx 和仪表盘展示）
# 全量数据会全部入库，这里只是筛选展示用的子集
FLAGSHIP_MODELS = {
    "OpenAI": {
        "旗舰": "openai/gpt-5.5-pro",
        "次旗舰": "openai/gpt-5.5",
    },
    "Anthropic": {
        "旗舰": "anthropic/claude-opus-4.8",
        "次旗舰": "anthropic/claude-sonnet-4.6",
    },
    "Google": {
        "旗舰": "google/gemini-3.1-pro-preview",
        "次旗舰": "google/gemini-3.5-flash",
    },
    "智谱 (Z.ai)": {
        "旗舰": "z-ai/glm-5.2",
        "次旗舰": "z-ai/glm-5.1",
    },
    "MiniMax": {
        "旗舰": "minimax/minimax-m3",
        "次旗舰": "minimax/minimax-m2.7",
    },
    "Qwen": {
        "旗舰": "qwen/qwen3.7-max",
        "次旗舰": "qwen/qwen3.7-plus",
    },
    # Kimi 已确认匹配；豆包真旗舰(doubao-seed-2.0-pro)在数据源无价，改用有价的 Seed 2.0 代
    "Kimi (Moonshot)": {
        "旗舰": "moonshotai/kimi-k2.6",
        "次旗舰": "moonshotai/kimi-k2.5",
    },
    "豆包 (Doubao)": {
        "旗舰": "bytedance-seed/seed-2.0-lite",
        "次旗舰": "bytedance-seed/seed-2.0-mini",
    },
}

# 反向映射：model_id -> (厂商显示名, 定位)
MODEL_TO_TIER = {}
for provider, tiers in FLAGSHIP_MODELS.items():
    for tier, model_id in tiers.items():
        MODEL_TO_TIER[model_id] = (provider, tier)

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


def build_flagship_rows(models, today):
    """从全量数据中筛出 12 个白名单模型，构建 Excel 行"""
    rows = []
    found = set()
    for model_id, m in models.items():
        if model_id not in MODEL_TO_TIER:
            continue
        provider, tier = MODEL_TO_TIER[model_id]
        found.add(model_id)

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

    missing = set(MODEL_TO_TIER.keys()) - found
    if missing:
        log(f"⚠ 以下 {len(missing)} 个白名单模型未在全量数据中找到：")
        for mid in sorted(missing):
            p, t = MODEL_TO_TIER[mid]
            log(f"   - [{p}/{t}] {mid}")
        log("请检查并更新脚本顶部 FLAGSHIP_MODELS 配置。")
    return rows


def save_flagship_to_excel(rows, today):
    """12 个旗舰模型写入 Excel（同日覆盖）"""
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

    today = datetime.now().strftime("%Y-%m-%d")  # GitHub Actions 在 UTC 1:00 跑，对应北京时间 9:00，日期一致

    # 1) 全量写入 SQLite
    total = save_full_to_db(models, generated_at, today)

    # 2) 筛出 12 个旗舰写入 Excel（供可视化）
    flagship_rows = build_flagship_rows(models, today)
    if not flagship_rows:
        log("未匹配到任何白名单模型，退出")
        sys.exit(1)

    counts = Counter(r["厂商"] for r in flagship_rows)
    summary = " / ".join(f"{k}: {v}" for k, v in counts.items())
    log(f"旗舰匹配 {len(flagship_rows)}/{len(MODEL_TO_TIER)} 个 — {summary}")

    save_flagship_to_excel(flagship_rows, today)

    log(f"完成：全量 {total} 模型 → DB；旗舰 {len(flagship_rows)} → Excel")
    log("=" * 60)


if __name__ == "__main__":
    main()
