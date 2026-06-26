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

# OpenRouter 价格补丁：仅用于修正 z-ai/智谱 的定价偏差
# tokenpricing 聚合 z-ai 渠道时价格滞后/偏低（如 GLM-5.2 少报 ~40%）；
# OpenRouter 是 z-ai 模型的权威报价来源，直接拉取可修正偏差。
# 失败时自动降级（保留 tokenpricing 原始价格），不中断主流程。
OPENROUTER_MODELS_URL = "https://openrouter.ai/api/v1/models"
OPENROUTER_PATCH_PROVIDERS = {"z-ai"}  # 只补充这些前缀的模型，其余厂商不受影响

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
}

# ---------- 旗舰系列动态识别规则 ----------
# 不再写死具体型号，而是为每家定义"旗舰系列正则"：
# 每天在该系列内按版本号排序，取最新一代=旗舰、上一代=次旗舰。
# 这样型号换代（如 gpt-5.5-pro → gpt-6-pro）时自动跟上，趋势线连续不断。
#
# 正则用第 1 个捕获组提取"主版本号"（如 5.5、4.8、3.1），用于排序。
# 正则需匹配 model_id 去掉厂商前缀后的"短名"（小写），并排除
# fast/turbo/image/mini/lite/flash/air/coder/thinking 等变体与带日期戳的快照。
#
# 注：豆包(Doubao)真旗舰 seed-2.0-pro 在本数据源(tokenpricing)无价格，
#     整个系列只有 lite/mini 有价，无法体现真实旗舰价，故不跟踪豆包。
FLAGSHIP_SERIES = {
    # OpenAI 用标准版(gpt-5.5/5.4)而非 pro 增强版：
    # gpt-5.5-pro(出$180) 是类似 o1-pro 的特殊增强版，价格畸高，会拉变形横向对比；
    # 标准版 gpt-5.5(出$30) 才与 opus-4.8(出$25) 同档可比。排除 pro/mini/nano/codex/chat/image 变体。
    "OpenAI":          r"^gpt-(\d+(?:\.\d+)?)$",
    "Anthropic":       r"^claude-opus-(\d+(?:\.\d+)?)$",
    "Google":          r"^gemini-(\d+(?:\.\d+)?)-pro(?:-preview)?$",
    "智谱 (Z.ai)":     r"^glm-(\d+(?:\.\d+)?)$",
    "MiniMax":         r"^minimax-m(\d+(?:\.\d+)?)$",
    "Qwen":            r"^qwen(\d+(?:\.\d+)?)-max$",
    "Kimi (Moonshot)": r"^kimi-k(\d+(?:\.\d+)?)$",
}

# 跟踪厂商（与 FLAGSHIP_SERIES 一致）
TRACKED_PROVIDERS = list(FLAGSHIP_SERIES.keys())

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


def select_flagships(models):
    """
    按 FLAGSHIP_SERIES 规则，从全量数据中为每家动态挑出旗舰+次旗舰。
    规则：在该家旗舰系列内，按版本号降序排列，最新一代=旗舰、上一代=次旗舰。
    仅考虑有输出价格(>0)的型号。返回 {model_id: (厂商, 定位)}。
    """
    # 先按厂商归集候选：provider_display -> [(version_tuple, model_id, out_price), ...]
    candidates = {p: [] for p in FLAGSHIP_SERIES}
    seen_versions = {p: set() for p in FLAGSHIP_SERIES}  # 同版本去重(大小写/前缀重复)

    for model_id, m in models.items():
        provider_raw = m.get("provider", "")
        provider = PROVIDER_DISPLAY.get(provider_raw, None)
        if provider not in FLAGSHIP_SERIES:
            continue

        pricing = m.get("pricing", {}) or {}
        out_price = _to_float(pricing.get("output_per_million"))
        if out_price <= 0:  # 无价型号不参与
            continue

        short = model_id.split("/")[-1].lower()
        match = re.match(FLAGSHIP_SERIES[provider], short, re.IGNORECASE)
        if not match:
            continue

        ver = _parse_version(match.group(1))
        if ver in seen_versions[provider]:
            continue  # 同版本只保留第一个，避免大小写/前缀重复
        seen_versions[provider].add(ver)
        candidates[provider].append((ver, model_id))

    # 每家按版本降序，取前两名
    result = {}
    for provider, items in candidates.items():
        items.sort(key=lambda x: x[0], reverse=True)
        if len(items) >= 1:
            result[items[0][1]] = (provider, "旗舰")
        if len(items) >= 2:
            result[items[1][1]] = (provider, "次旗舰")
    return result


def build_flagship_rows(models, today):
    """按系列规则动态挑出每家旗舰+次旗舰，构建 Excel 行"""
    model_to_tier = select_flagships(models)
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

    # 检查每家是否都凑齐了旗舰+次旗舰
    for provider in TRACKED_PROVIDERS:
        tiers_found = {tier for _, (p, tier) in model_to_tier.items() if p == provider}
        missing = {"旗舰", "次旗舰"} - tiers_found
        if missing:
            log(f"⚠ {provider} 缺少 {'/'.join(missing)}（系列内有价型号不足两代），请检查 FLAGSHIP_SERIES 正则")

    # 排序：按 TRACKED_PROVIDERS 顺序、旗舰在前
    tier_order = {"旗舰": 0, "次旗舰": 1}
    rows.sort(key=lambda r: (TRACKED_PROVIDERS.index(r["厂商"]), tier_order.get(r["定位"], 9)))
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


def patch_prices_from_openrouter(models):
    """
    从 OpenRouter 补充 OPENROUTER_PATCH_PROVIDERS 中指定前缀的模型实时价格。
    仅覆写已存在模型的定价字段，不新增/删除模型，不影响其他厂商。
    OpenRouter 价格格式：per-token 字符串（"0.0000044"），需 x1,000,000 转为 per-million。
    任何异常均降级处理：打印警告并返回原始 models，主流程不中断。
    """
    try:
        headers = {"User-Agent": "API-Price-Tracer/1.0", "Accept": "application/json"}
        r = requests.get(OPENROUTER_MODELS_URL, headers=headers, timeout=30)
        r.raise_for_status()
        or_data = r.json().get("data", [])
    except Exception as e:
        log(f"⚠ OpenRouter 补充价格失败，继续使用 tokenpricing 原始价格: {e}")
        return models

    patched = 0
    for m in or_data:
        model_id = m.get("id", "")
        prefix = model_id.split("/")[0] if "/" in model_id else ""
        if prefix not in OPENROUTER_PATCH_PROVIDERS:
            continue
        if model_id not in models:
            continue  # 只更新已有模型，不增删条目

        pricing = m.get("pricing", {}) or {}
        try:
            # OpenRouter 价格单位是 per-token（字符串），x 1,000,000 转为 per-million
            inp = float(pricing.get("prompt") or 0) * 1_000_000
            out = float(pricing.get("completion") or 0) * 1_000_000
        except (TypeError, ValueError):
            continue
        if out <= 0:
            continue

        old_out = _to_float((models[model_id].get("pricing") or {}).get("output_per_million"))
        models[model_id].setdefault("pricing", {})
        models[model_id]["pricing"]["input_per_million"] = inp
        models[model_id]["pricing"]["output_per_million"] = out
        log(f"  GLM 价格更新: {model_id}  出 ${old_out:.3f}/M -> ${out:.3f}/M")
        patched += 1

    log(f"OpenRouter 价格补丁完成：{patched} 个 z-ai 模型已更新")
    return models


def main():
    log("=" * 60)
    log("开始抓取 API 价格（tokenpricing 全量 + 自建历史）")

    models, generated_at = fetch_full_prices()
    if not models:
        log("未获取到模型数据，退出")
        sys.exit(1)

    # 补充智谱 GLM 的实时价格（OpenRouter 为权威来源，tokenpricing 聚合偏差约 30~40%）
    models = patch_prices_from_openrouter(models)

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
    log(f"旗舰动态识别 {len(flagship_rows)} 个（{len(TRACKED_PROVIDERS)} 家 × 旗舰+次旗舰）— {summary}")

    save_flagship_to_excel(flagship_rows, today)

    log(f"完成：全量 {total} 模型 -> DB；旗舰 {len(flagship_rows)} -> Excel")
    log("=" * 60)


if __name__ == "__main__":
    main()
