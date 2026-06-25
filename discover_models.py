# -*- coding: utf-8 -*-
"""
探查脚本（一次性，本机运行）：
  1) 列出 tokenpricing 数据源里 Kimi(Moonshot) / 豆包(Doubao/ByteDance) 的真实模型 ID + 价格
  2) 检查数据源 GitHub 仓库是否有六月历史快照目录可回填

用法（在项目目录下）：
  .venv\\Scripts\\python.exe discover_models.py
把屏幕输出整段复制发回给我即可。
"""
import json
import requests

CURRENT_URL = "https://raw.githubusercontent.com/Atena-IT/tokenpricing/main/database/current/prices.json"
# GitHub API：列出 database 目录，看有没有 history / 日期快照
REPO_CONTENTS_API = "https://api.github.com/repos/Atena-IT/tokenpricing/contents/database"
HEADERS = {"User-Agent": "discover/1.0", "Accept": "application/json"}

# 我们关心的厂商关键词（provider 字段或 model_id 里出现即匹配）
KEYWORDS = ["moonshot", "kimi", "bytedance", "doubao", "seed"]


def main():
    print("=" * 70)
    print("【1】拉取当前价格，查找 Kimi / 豆包 相关模型")
    print("=" * 70)
    r = requests.get(CURRENT_URL, headers=HEADERS, timeout=60)
    r.raise_for_status()
    data = r.json()
    models = data.get("models", {})
    print(f"数据源 generated_at = {data.get('generated_at')}，模型总数 = {len(models)}\n")

    # 先看 provider 字段都有哪些值，方便确认前缀
    providers = sorted({(m.get("provider") or "") for m in models.values()})
    hit_providers = [p for p in providers if any(k in p.lower() for k in KEYWORDS)]
    print("匹配到的 provider 字段值：", hit_providers, "\n")

    found = []
    for mid, m in models.items():
        blob = (mid + " " + str(m.get("provider", "")) + " " + str(m.get("display_name", ""))).lower()
        if any(k in blob for k in KEYWORDS):
            p = m.get("pricing", {}) or {}
            found.append((
                mid,
                m.get("provider", ""),
                m.get("display_name", ""),
                m.get("category", ""),
                p.get("input_per_million"),
                p.get("output_per_million"),
                m.get("context_window"),
            ))

    found.sort(key=lambda x: (x[1], -(x[4] or 0)))
    print(f"共匹配 {len(found)} 个模型：\n")
    print(f"{'模型ID':<42} {'provider':<14} {'分类':<10} {'输入$/M':>9} {'输出$/M':>9}")
    print("-" * 95)
    for mid, prov, name, cat, inp, outp, ctx in found:
        print(f"{mid:<42} {prov:<14} {str(cat):<10} {str(inp):>9} {str(outp):>9}   {name}")

    print("\n" + "=" * 70)
    print("【2】检查数据源是否有历史快照（六月数据回填用）")
    print("=" * 70)
    try:
        rc = requests.get(REPO_CONTENTS_API, headers=HEADERS, timeout=60)
        rc.raise_for_status()
        entries = rc.json()
        print("database/ 目录下的条目：")
        for e in entries:
            print(f"  [{e.get('type')}] {e.get('name')}")
        print("\n→ 若看到 history / archive / 日期类目录，说明有历史数据可回填；")
        print("  若只有 current，则数据源不保留历史，只能从今天起逐日累积。")
    except Exception as e:
        print("查询目录失败：", e)


if __name__ == "__main__":
    main()
