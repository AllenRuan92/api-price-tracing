# -*- coding: utf-8 -*-
"""
回填历史数据脚本（档位口径版）

从 tokenpricing 的 database/history/ 目录拉取历史快照，每天取一份，
按 fetch_prices.select_tier_models() 的档位规则动态挑出各家主推档位型号，
合并进 prices.xlsx（已存在的日期不覆盖，只补缺失日期）。
"""
import os, sys, time, json, re
from collections import defaultdict
import requests
import pandas as pd

import fetch_prices as fp  # 复用配置、档位规则与列定义

H = {"User-Agent": "backfill", "Accept": "application/vnd.github+json"}
HIST_API = "https://api.github.com/repos/Atena-IT/tokenpricing/contents/database/history"
RAW_BASE = "https://raw.githubusercontent.com/Atena-IT/tokenpricing/main/database/history/"

EXCEL_PATH = fp.EXCEL_PATH


def get_with_retry(url, tries=4, delay=4, is_json=True):
    for i in range(1, tries + 1):
        try:
            r = requests.get(url, headers=H, timeout=40)
            r.raise_for_status()
            return r.json() if is_json else r.text
        except Exception as e:
            print(f"  请求失败({i}/{tries}) {type(e).__name__}: {str(e)[:80]}")
            if i < tries:
                time.sleep(delay)
    return None


def list_history_files():
    """列出 history 目录所有快照文件名"""
    items = get_with_retry(HIST_API)
    if not items:
        print("无法列出 history 目录")
        return []
    files = [it["name"] for it in items if it["type"] == "file" and it["name"].endswith(".json")]
    return sorted(files)


def pick_one_per_day(files):
    """每天只取第一个时间点的文件。文件名形如 prices-20260614T184401Z.json"""
    by_day = {}
    for f in files:
        m = re.search(r"prices-(\d{8})T(\d{6})Z", f)
        if not m:
            continue
        day = m.group(1)  # YYYYMMDD
        if day not in by_day or f < by_day[day]:
            by_day[day] = f
    return dict(sorted(by_day.items()))


def extract_tier_rows(models, date_str):
    """从一份快照的 models dict 按档位规则抽出主推型号行（复用 fetch_prices 逻辑）"""
    rows = fp.build_flagship_rows(models, date_str)
    for r in rows:
        r["数据源"] = "tokenpricing-history"
    return rows


def main():
    print("=== 回填六月历史 ===")
    # 已有日期
    existing_dates = set()
    if os.path.exists(EXCEL_PATH):
        cur = pd.read_excel(EXCEL_PATH, sheet_name="价格数据")
        cur["抓取日期"] = pd.to_datetime(cur["抓取日期"]).dt.strftime("%Y-%m-%d")
        existing_dates = set(cur["抓取日期"].unique())
        print(f"prices.xlsx 已有日期：{sorted(existing_dates)}")
    else:
        cur = pd.DataFrame(columns=fp.EXCEL_COLUMNS)

    files = list_history_files()
    print(f"history 目录共 {len(files)} 个快照文件")
    per_day = pick_one_per_day(files)
    print(f"覆盖 {len(per_day)} 个不同日期：{list(per_day.keys())}")

    all_new = []
    for day, fname in per_day.items():
        date_str = f"{day[:4]}-{day[4:6]}-{day[6:8]}"
        if date_str in existing_dates:
            print(f"  跳过 {date_str}（已存在）")
            continue
        print(f"  拉取 {date_str} ← {fname}")
        data = get_with_retry(RAW_BASE + fname)
        if not data:
            print(f"    ✗ {date_str} 拉取失败，跳过")
            continue
        models = data.get("models", {})
        rows = extract_tier_rows(models, date_str)
        matched = len(rows)
        all_new.extend(rows)
        print(f"    ✓ {date_str}: 命中 {matched} 个主推档位模型")
        time.sleep(0.5)  # 轻微限速，避免触发 GitHub 限流

    if not all_new:
        print("没有需要回填的新日期。")
        return

    new_df = pd.DataFrame(all_new, columns=fp.EXCEL_COLUMNS)
    combined = pd.concat([cur, new_df], ignore_index=True)
    combined = combined.sort_values(["抓取日期", "厂商", "定位"]).reset_index(drop=True)
    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as w:
        combined.to_excel(w, sheet_name="价格数据", index=False)
    print(f"\n回填完成：新增 {len(all_new)} 行，prices.xlsx 累计 {len(combined)} 行")
    print(f"现有日期：{sorted(set(combined['抓取日期']))}")


if __name__ == "__main__":
    main()
