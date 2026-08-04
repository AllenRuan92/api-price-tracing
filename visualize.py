# -*- coding: utf-8 -*-
"""
API 价格可视化脚本（ECharts CDN 版 · 现代浅色商务风 v3）

v3 在 v2 基础上改造价格趋势图，让"谁变了、变了多少"一眼可见：
  - 时间范围切换：近7天 / 近14天 / 全部（默认近7天）
  - 视图模式切换：绝对价（默认，对数/线性轴）/ 归一化指数（窗口首日=100）
  - 近期变动提示带：自动列出窗口内调价的模型（阈值 0.5%，涨红降跌绿）
  - 厂商简称：Kimi(Moonshot)→Kimi，避免 chips/图例过长

数据来源：prices.xlsx（fetch_prices.py 产出，每天每家每条主推产品线各 1 个型号）
生成文件：price_dashboard.html（单文件，零本地 JS 依赖，数据前端渲染）
口径：只跟踪各家当前主推的产品线，不做强/中/弱档位划分；每家有几条主推线就几条。
"""

import os
import sys
import json
from datetime import datetime

import pandas as pd

# ---------- 配置 ----------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(SCRIPT_DIR, "prices.xlsx")
OUTPUT_HTML = os.path.join(SCRIPT_DIR, "price_dashboard.html")

ECHARTS_CDN = "https://cdn.jsdelivr.net/npm/echarts@5/dist/echarts.min.js"

PROVIDER_COLORS = {
    "OpenAI": "#10A37F",
    "Anthropic": "#D97757",
    "Google": "#4285F4",
    "智谱 (Z.ai)": "#3B5BFE",
    "MiniMax": "#FF6B35",
    "Qwen": "#615CED",
    "Kimi (Moonshot)": "#1F2430",
    "DeepSeek": "#9B59B6",
}

PROVIDER_ORDER = ["OpenAI", "Anthropic", "Google", "智谱 (Z.ai)",
                  "MiniMax", "Qwen", "Kimi (Moonshot)", "DeepSeek"]


def load_data():
    """读取 Excel 数据"""
    if not os.path.exists(EXCEL_PATH):
        print(f"错误：找不到数据文件 {EXCEL_PATH}")
        sys.exit(1)
    df = pd.read_excel(EXCEL_PATH, sheet_name="价格数据")
    df["抓取日期"] = pd.to_datetime(df["抓取日期"]).dt.strftime("%Y-%m-%d")
    return df


def _series_order(df):
    """
    每家产品线的展示顺序。优先取 fetch_prices.WATCHLIST 的书写顺序（主线在前），
    白名单里没有的（历史遗留产品线）按字母序追加在后，保证顺序确定、图例样式稳定。
    """
    try:
        import fetch_prices as fp
        preset = {p: list(s.keys()) for p, s in fp.WATCHLIST.items()}
    except Exception:
        preset = {}
    out = {}
    for provider, g in df.groupby("厂商"):
        present = set(g["产品线"].dropna().astype(str))
        ordered = [s for s in preset.get(provider, []) if s in present]
        ordered += sorted(present - set(ordered))
        out[provider] = ordered
    return out


def build_records(df):
    """把 DataFrame 序列化为前端可用的记录列表（一次性注入页面）"""
    records = []
    for _, r in df.iterrows():
        model_id = str(r["模型ID"])
        model_short = model_id.split("/")[-1] if "/" in model_id else model_id
        in_price = r["输入价格(美元/百万token)"]
        out_price = r["输出价格(美元/百万token)"]
        records.append({
            "date": r["抓取日期"],
            "provider": r["厂商"],
            "series": r["产品线"],
            "model_id": model_id,
            "model_short": model_short,
            "model_name": r.get("模型名称", model_short),
            "input": None if pd.isna(in_price) else round(float(in_price), 4),
            "output": None if pd.isna(out_price) else round(float(out_price), 4),
            "is_free": (str(r.get("是否免费", "否")) == "是"),
        })
    return records


def build_meta(df):
    """汇总元信息供前端展示"""
    dates = sorted(df["抓取日期"].unique())
    return {
        "dates": dates,
        "latest": dates[-1] if dates else "",
        "prev": dates[-2] if len(dates) >= 2 else "",
        "provider_order": PROVIDER_ORDER,
        "provider_colors": PROVIDER_COLORS,
        "series_order": _series_order(df),
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }


def build_html(df):
    records = build_records(df)
    meta = build_meta(df)
    data_json = json.dumps(records, ensure_ascii=False)
    meta_json = json.dumps(meta, ensure_ascii=False)

    html = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>大模型 API 价格追踪仪表盘</title>
<script src="__ECHARTS_CDN__"></script>
<style>
  :root{
    --bg:#F7F8FA; --card:#FFFFFF; --line:#ECEEF2; --ink:#1F2430;
    --muted:#8A92A6; --muted2:#AEB4C2; --accent:#3B5BFE;
    --up:#E5484D; --down:#1FA971; --shadow:0 1px 3px rgba(20,24,40,.06),0 1px 2px rgba(20,24,40,.04);
    --radius:14px;
  }
  *{margin:0;padding:0;box-sizing:border-box}
  body{
    font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",sans-serif;
    background:var(--bg); color:var(--ink); -webkit-font-smoothing:antialiased;
  }
  .wrap{max-width:1280px;margin:0 auto;padding:0 20px 48px}

  .topbar{padding:26px 0 18px;display:flex;align-items:flex-end;justify-content:space-between;flex-wrap:wrap;gap:12px}
  .topbar h1{font-size:22px;font-weight:700;letter-spacing:-.2px}
  .topbar .sub{font-size:13px;color:var(--muted);margin-top:6px;display:flex;gap:8px;align-items:center;flex-wrap:wrap}
  .badge{display:inline-flex;align-items:center;gap:5px;background:#EEF2FF;color:#3B5BFE;
         font-size:12px;font-weight:600;padding:3px 9px;border-radius:999px}
  .badge.gray{background:#F1F3F7;color:#6B7280}
  .updated{font-size:12px;color:var(--muted2)}

  .kpis{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin:6px 0 18px}
  .kpi{background:var(--card);border:1px solid var(--line);border-radius:var(--radius);
       padding:16px 16px 14px;box-shadow:var(--shadow);position:relative;overflow:hidden}
  .kpi::before{content:"";position:absolute;left:0;top:0;height:3px;width:100%;background:var(--bar,#3B5BFE);opacity:.9}
  .kpi .label{font-size:12px;color:var(--muted);margin-bottom:8px}
  .kpi .value{font-size:22px;font-weight:700;line-height:1.2;letter-spacing:-.3px}
  .kpi .note{font-size:12px;color:var(--muted2);margin-top:4px}

  .card{background:var(--card);border:1px solid var(--line);border-radius:var(--radius);
        box-shadow:var(--shadow);padding:18px 18px 8px;margin-bottom:16px}
  .card .hd{display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;margin-bottom:6px}
  .card .hd h3{font-size:15px;font-weight:650}
  .card .hd .desc{font-size:12px;color:var(--muted2)}

  .controls{display:flex;gap:8px;flex-wrap:wrap;align-items:center}
  .chips{display:flex;gap:6px;flex-wrap:wrap}
  .chip{font-size:12px;font-weight:600;color:#5A6273;background:#F1F3F7;border:1px solid transparent;
        padding:5px 11px;border-radius:999px;cursor:pointer;transition:.15s;user-select:none}
  .chip:hover{background:#E7EAF1}
  .chip.active{background:#1F2430;color:#fff}
  .seg{display:inline-flex;background:#F1F3F7;border-radius:999px;padding:3px}
  .seg button{border:0;background:transparent;font-size:12px;font-weight:600;color:#5A6273;
        padding:5px 12px;border-radius:999px;cursor:pointer;transition:.15s}
  .seg button.active{background:#fff;color:var(--ink);box-shadow:0 1px 2px rgba(0,0,0,.08)}

  .change-bar{display:flex;align-items:center;gap:8px;flex-wrap:wrap;
        background:#FAFBFD;border:1px solid var(--line);border-radius:10px;
        padding:9px 12px;margin:2px 0 6px;font-size:12px;color:#5A6273;min-height:38px}
  .change-bar .cb-title{font-weight:650;color:var(--ink);margin-right:2px}
  .change-bar .cb-none{color:var(--muted2)}
  .cb-item{display:inline-flex;align-items:center;gap:4px;background:#fff;border:1px solid var(--line);
        border-radius:999px;padding:3px 9px;font-weight:600}
  .cb-item .cb-dot{width:7px;height:7px;border-radius:50%}
  .cb-item .cb-pct.up{color:var(--up)} .cb-item .cb-pct.down{color:var(--down)}

  .grid2{display:grid;grid-template-columns:1fr 1fr;gap:16px}

  .chart{width:100%}
  #c_trend{height:440px}
  #c_compare{height:380px}
  #c_scatter{height:380px}
  #c_rank{height:440px}

  table{width:100%;border-collapse:collapse;font-size:13px}
  thead th{text-align:left;font-size:12px;color:var(--muted);font-weight:600;
           padding:10px 10px;border-bottom:1px solid var(--line);cursor:pointer;white-space:nowrap;user-select:none}
  thead th .arrow{color:var(--accent);font-size:10px;margin-left:3px}
  tbody td{padding:11px 10px;border-bottom:1px solid #F2F4F7}
  tbody tr:hover{background:#FAFBFD}
  .dot{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:7px;vertical-align:middle}
  .num{font-variant-numeric:tabular-nums}
  .chg{font-size:12px;font-weight:600}
  .chg.up{color:var(--up)} .chg.down{color:var(--down)} .chg.flat{color:#9aa1b1}

  .footer{text-align:center;color:var(--muted2);font-size:12px;padding:22px 0 4px}

  @media(max-width:880px){
    .kpis{grid-template-columns:repeat(2,1fr)}
    .grid2{grid-template-columns:1fr}
  }
</style>
</head>
<body>
<div class="wrap">

  <div class="topbar">
    <div>
      <h1>大模型 API 价格追踪仪表盘</h1>
      <div class="sub">
        <span class="badge">tokenpricing 数据源</span>
        <span class="badge gray">OpenRouter + LiteLLM 聚合</span>
        <span class="badge gray">美元国际站报价</span>
        <span id="sub_summary">各家主推产品线</span>
      </div>
    </div>
    <div class="updated">生成于 __GENERATED_AT__</div>
  </div>

  <div class="kpis" id="kpis"></div>

  <!-- 主图：价格趋势 -->
  <div class="card">
    <div class="hd">
      <div>
        <h3>价格趋势</h3>
        <div class="desc">一条线＝一条产品线（型号换代自动接续）；同色为同一厂商，实线/虚线区分该厂商的不同产品线；单击厂商=只看该家，再点或点「全部厂商」恢复</div>
      </div>
      <div class="controls">
        <div class="seg" id="seg_range">
          <button data-range="7" class="active">近7天</button>
          <button data-range="14">近14天</button>
          <button data-range="all">全部</button>
        </div>
        <div class="seg" id="seg_mode">
          <button data-mode="abs" class="active">绝对价</button>
          <button data-mode="index">归一化指数</button>
        </div>
        <div class="seg" id="seg_scale">
          <button data-scale="log" class="active">对数轴</button>
          <button data-scale="value">线性轴</button>
        </div>
        <div class="seg" id="seg_io">
          <button data-io="input" class="active">输入价</button>
          <button data-io="output">输出价</button>
        </div>
      </div>
    </div>
    <div id="change_bar" class="change-bar"></div>
    <div class="chips" id="chips_provider" style="margin:4px 0 10px"></div>
    <div id="c_trend" class="chart"></div>
  </div>

  <div class="grid2">
    <div class="card">
      <div class="hd"><div><h3>最新输入价对比</h3><div class="desc">每根柱＝一个在跟踪的型号，按厂商配色分组</div></div></div>
      <div id="c_compare" class="chart"></div>
    </div>
    <div class="card">
      <div class="hd"><div><h3>输入 / 输出价格分布</h3><div class="desc">同色为同一厂商，形状区分该厂商的不同产品线；越靠左下越便宜</div></div></div>
      <div id="c_scatter" class="chart"></div>
    </div>
  </div>

  <div class="card">
    <div class="hd"><div><h3>综合价格排行</h3><div class="desc">综合价 = 输入 + 输出（$/百万 token），越靠上越便宜</div></div></div>
    <div id="c_rank" class="chart"></div>
  </div>

  <div class="card">
    <div class="hd">
      <div><h3>最新价格明细</h3><div class="desc">点击表头排序；涨跌为对比上一快照</div></div>
    </div>
    <div style="overflow-x:auto">
      <table id="tbl">
        <thead><tr>
          <th data-key="provider">厂商</th>
          <th data-key="model_short">模型</th>
          <th data-key="input" class="num">输入价 ($/M)</th>
          <th data-key="output" class="num">输出价 ($/M)</th>
          <th data-key="chg" class="num">输入价涨跌</th>
        </tr></thead>
        <tbody id="tbody"></tbody>
      </table>
    </div>
  </div>

  <div class="footer">由 API Price Tracing 自动生成 · 每日 9:00（北京）抓取更新 · 数据源 tokenpricing<br/>价格为各厂商美元国际站报价，便于横向对比；国内厂商（智谱/Qwen/Kimi/MiniMax/DeepSeek）的美元国际价可能与其人民币官方定价不同。</div>
</div>

<script>
const DATA = __DATA_JSON__;
const META = __META_JSON__;
const COLORS = META.provider_colors;
// ORDER 只保留数据中真实存在的厂商（按预设顺序），避免出现无数据的空 chip
const _present = new Set(DATA.map(d=>d.provider));
const ORDER = META.provider_order.filter(p=>_present.has(p));
const LATEST = META.latest, PREV = META.prev;
// 厂商简称（X轴/图例/chips 显示用，避免长名重叠）
const SHORT = {
  "智谱 (Z.ai)":"智谱", "Kimi (Moonshot)":"Kimi"
};
const shortName = p => SHORT[p] || p;

// 产品线配置：同一厂商用同一颜色，靠线型/符号区分它的不同产品线（非强弱分级，
// 只是"该厂商第 1 / 第 2 条主推线"的视觉区分）。顺序来自 META.series_order。
const SERIES_ORDER = META.series_order || {};
const LINE_VARIANT   = [{w:2.6,type:"solid"},{w:1.9,type:"dashed"},{w:1.5,type:"dotted"}];
const SYMBOL_VARIANT = ["circle","diamond","triangle","rect"];
const seriesList = p => SERIES_ORDER[p] || [];
const seriesIdx  = (p,s) => { const i = seriesList(p).indexOf(s); return i<0?0:i; };
// 产品线名（GPT / Claude Opus / …）本身已自带厂商信息，图例直接用它；
// 万一两家用了同名产品线，则加厂商前缀消歧，避免 ECharts 图例撞名。
const _nameCount = {};
DATA.forEach(d=>{ _nameCount[d.series] = (_nameCount[d.series]||new Set()).add(d.provider); });
const legendName = (p,s) => (_nameCount[s] && _nameCount[s].size>1) ? (shortName(p)+"·"+s) : s;

// ---- 状态 ----
// range: 7|14|"all"；mode: abs 绝对价 / index 归一化指数；scale: log/value（仅绝对价生效）
let state = { io:"input", range:"7", mode:"abs", scale:"log", providers:new Set(ORDER) };

function visibleDates(){
  const all = META.dates;
  if(state.range==="all") return all;
  const n = parseInt(state.range,10);
  return all.slice(Math.max(0, all.length-n));
}

// ---- 工具 ----
const fmt = v => (v==null ? "—" : "$"+Number(v).toFixed(Number(v)<1?3:2));
function paidLatest(){ return DATA.filter(d=>d.date===LATEST && !d.is_free); }

function priceAt(model_id, date, io){
  const r = DATA.find(d=>d.model_id===model_id && d.date===date);
  return r ? r[io] : null;
}
function deltaOf(model_id, io){
  if(!PREV) return null;
  const now = priceAt(model_id, LATEST, io);
  const before = priceAt(model_id, PREV, io);
  if(now==null || before==null || before===0) return null;
  const pct = (now-before)/before*100;
  return { pct, dir: pct>0.001?"up":(pct<-0.001?"down":"flat") };
}

// ================= KPI =================
function renderKPI(){
  const latest = DATA.filter(d=>d.date===LATEST);
  const providers = new Set(latest.map(d=>d.provider)).size;
  const models = new Set(latest.map(d=>d.model_id)).size;
  const span = META.dates.length>1 ? (META.dates[0]+" ~ "+LATEST) : LATEST;
  const cards = [
    {label:"跟踪厂商", value:providers, note:"主流大模型厂商", bar:"#3B5BFE"},
    {label:"跟踪模型", value:models, note:"各家主推产品线", bar:"#10A37F"},
    {label:"数据跨度", value:META.dates.length+" 天", note:span, bar:"#FF6B35", small:true},
  ];
  document.getElementById("kpis").innerHTML = cards.map(c=>`
    <div class="kpi" style="--bar:${c.bar}">
      <div class="label">${c.label}</div>
      <div class="value" style="${c.small?'font-size:15px':''}">${c.value}</div>
      <div class="note">${c.note}</div>
    </div>`).join("");
}

// ================= 厂商 chips =================
function renderChips(){
  const box = document.getElementById("chips_provider");
  const all = document.createElement("span");
  all.className = "chip"+(state.providers.size===ORDER.length?" active":"");
  all.textContent = "全部厂商";
  all.onclick = ()=>{ state.providers = new Set(ORDER); syncChips(); drawTrend(); };
  box.appendChild(all);
  ORDER.forEach(p=>{
    const c = document.createElement("span");
    c.className = "chip"+(state.providers.has(p)?" active":"");
    c.dataset.p = p;
    c.innerHTML = `<span class="dot" style="background:${COLORS[p]}"></span>${shortName(p)}`;
    c.onclick = ()=>{
      // 单击=只看这一家；若当前已是"只剩这一家"，再点则恢复全部
      if(state.providers.size===1 && state.providers.has(p)){
        state.providers = new Set(ORDER);
      } else {
        state.providers = new Set([p]);
      }
      syncChips(); drawTrend();
    };
    box.appendChild(c);
  });
}
function syncChips(){
  const box = document.getElementById("chips_provider");
  [...box.children].forEach(c=>{
    if(!c.dataset.p) c.classList.toggle("active", state.providers.size===ORDER.length);
    else c.classList.toggle("active", state.providers.has(c.dataset.p));
  });
}

// ================= 图表 =================
let trendChart, compareChart, scatterChart, rankChart;

function drawTrend(){
  const io = state.io;
  const dates = visibleDates();
  const baseDate = dates[0];
  const isIndex = state.mode==="index";

  // 按「厂商+产品线」分组（而非具体型号），这样型号换代时同一条线连续不断。
  // 每天取该组当天对应的型号价格；型号名记录在数据点上供 tooltip 显示。
  const groups = [];
  ORDER.filter(p=>state.providers.has(p)).forEach(p=>{
    seriesList(p).forEach(sname=>{
      const rowsAll = DATA.filter(d=>d.provider===p && d.series===sname && !d.is_free);
      if(rowsAll.length===0) return;
      groups.push({provider:p, series:sname, rows:rowsAll});
    });
  });

  const series = groups.map(g=>{
    const ls = LINE_VARIANT[seriesIdx(g.provider,g.series)] || {w:1.8,type:"solid"};
    const atDate = dt=>{ const r=g.rows.find(x=>x.date===dt); return r||null; };
    const baseRow = atDate(baseDate);
    const base = baseRow? baseRow[io] : null;
    const data = dates.map(dt=>{
      const r = atDate(dt);
      if(!r || r[io]==null) return {value:null};
      const raw = r[io];
      const val = isIndex ? ((base==null||base===0)? null : Number((raw/base*100).toFixed(2))) : raw;
      return {value:val, model:r.model_short, raw};  // 带型号名和原始价
    });
    return {
      name: legendName(g.provider, g.series),
      type:"line", data, connectNulls:true, smooth:false, symbolSize:6,
      lineStyle:{ width: ls.w, type: ls.type },
      itemStyle:{ color: COLORS[g.provider]||"#999" },
    };
  });

  const hint = (dates.length<=1) ? [{
    type:"text", left:"center", top:24,
    style:{ fill:"#AEB4C2", fontSize:13,
      text:"当前窗口仅 "+dates.length+" 天数据，趋势将随每日抓取逐步形成" }
  }] : [];

  const yType = isIndex ? "value" : state.scale;
  const yName = isIndex ? "指数（窗口首日=100）"
                        : (state.scale==="log"?"$/百万 token（对数）":"$/百万 token");
  // tooltip 用 formatter（能访问数据点上的 model 型号名）：显示「产品线 型号 价格」
  const tipFormatter = params=>{
    if(!params || !params.length) return "";
    const date = params[0].axisValue;
    const lines = params
      .filter(pp=>pp.seriesName!=="基准 100" && pp.data && pp.data.value!=null)
      .map(pp=>{
        const d = pp.data;
        const valTxt = isIndex
          ? d.value.toFixed(1)+"（"+(d.value>=100?"+":"")+(d.value-100).toFixed(1)+"%）"
          : "$"+Number(d.raw).toFixed(3)+"/M";
        const model = d.model? ` <span style="color:#AEB4C2">${d.model}</span>` : "";
        return `${pp.marker}${pp.seriesName}${model}　<b>${valTxt}</b>`;
      });
    return `<div style="font-weight:600;margin-bottom:4px">${date}</div>` + lines.join("<br/>");
  };

  const baseSeries = isIndex ? series.concat([{
      type:"line", data:dates.map(()=>null), silent:true, showSymbol:false,
      markLine:{ symbol:"none", silent:true,
        lineStyle:{color:"#C7CCD8",type:"dashed"},
        label:{formatter:"基准 100",color:"#AEB4C2",fontSize:10,position:"insideEndTop"},
        data:[{yAxis:100}] } }]) : series;

  trendChart.setOption({
    graphic: hint,
    tooltip:{ trigger:"axis", confine:true, enterable:true, order:"valueDesc",
      extraCssText:"max-height:340px;overflow:auto;", formatter: tipFormatter },
    legend:{ type:"scroll", bottom:0, textStyle:{fontSize:11,color:"#5A6273"}, icon:"roundRect" },
    grid:{ left:62, right:48, top:18, bottom:54 },
    xAxis:{ type:"category", boundaryGap:false, data:dates,
            axisLine:{lineStyle:{color:"#E5E8EF"}},
            axisLabel:{color:"#8A92A6",fontSize:11,interval:dates.length>10?"auto":0},
            axisTick:{show:false} },
    yAxis:{ type:yType, name:yName, nameTextStyle:{color:"#AEB4C2",fontSize:11},
            min: isIndex? null : (state.scale==="log"?0.1:null),
            splitLine:{lineStyle:{color:"#F0F2F6"}},
            axisLabel:{color:"#8A92A6",fontSize:11,formatter: isIndex?"{value}":"${value}"} },
    series: baseSeries
  }, true);
}

// ================= 近期变动提示带 =================
const CHG_THRESHOLD = 0.5; // 价格变化 ≥0.5% 才算调价（滤掉浮点噪声）
function renderChangeBar(){
  const dates = visibleDates();
  const box = document.getElementById("change_bar");
  const rangeLabel = state.range==="all" ? "全部时段" : ("近"+state.range+"天");
  if(dates.length<2){
    box.innerHTML = `<span class="cb-title">${rangeLabel}调价</span><span class="cb-none">数据不足，暂无法比较</span>`;
    return;
  }
  const first=dates[0], last=dates[dates.length-1];
  const io = state.io;
  const ids = [...new Set(DATA.filter(d=>!d.is_free).map(d=>d.model_id))];
  const changes=[];
  ids.forEach(id=>{
    const a=priceAt(id,first,io), b=priceAt(id,last,io);
    if(a==null||b==null||a===0) return;
    const pct=(b-a)/a*100;
    if(Math.abs(pct)>=CHG_THRESHOLD){
      const r=DATA.find(d=>d.model_id===id);
      changes.push({ provider:r.provider, name:r.model_short, pct, dir:pct>0?"up":"down" });
    }
  });
  changes.sort((x,y)=>Math.abs(y.pct)-Math.abs(x.pct));
  const ioLabel = io==="input"?"输入":"输出";
  if(changes.length===0){
    box.innerHTML = `<span class="cb-title">${rangeLabel}调价</span>`+
      `<span class="cb-none">${ids.length} 个模型${ioLabel}价均无变动</span>`;
    return;
  }
  box.innerHTML = `<span class="cb-title">${rangeLabel}调价</span>`+
    changes.map(c=>`<span class="cb-item">
        <span class="cb-dot" style="background:${COLORS[c.provider]||'#ccc'}"></span>
        ${shortName(c.provider)} ${c.name} ${ioLabel}
        <span class="cb-pct ${c.dir}">${c.dir==="up"?"▲":"▼"} ${Math.abs(c.pct).toFixed(1)}%</span>
      </span>`).join("");
}

function drawCompare(){
  const latest = paidLatest();
  // 一根柱=一个在跟踪的型号，按厂商顺序排列（同厂商的多条产品线相邻），厂商配色
  const rows = [];
  ORDER.forEach(p=>{
    seriesList(p).forEach(sname=>{
      const r = latest.find(d=>d.provider===p && d.series===sname);
      if(r) rows.push(r);
    });
  });
  const cats = rows.map(r=>r.model_short);
  const vals = rows.map(r=>({
    value: r.input,
    itemStyle:{ color: COLORS[r.provider]||"#9DB0FF", borderRadius:[4,4,0,0] }
  }));
  // 图例只用来标厂商配色（用不画数据的占位 series 承载）
  const legendProviders = [...new Set(rows.map(r=>r.provider))];
  const legendStubs = legendProviders.map(p=>({
    name: shortName(p), type:"bar", data:[], itemStyle:{color:COLORS[p]}
  }));
  compareChart.setOption({
    tooltip:{ trigger:"axis", axisPointer:{type:"shadow"},
      formatter:params=>{
        const i = params[0].dataIndex, r = rows[i];
        return `${shortName(r.provider)} · ${r.series}<br/>${r.model_short}<br/>输入 <b>$${Number(r.input).toFixed(3)}</b>/M`;
      } },
    legend:{ data:legendProviders.map(shortName), top:0, textStyle:{fontSize:11,color:"#5A6273"},
             icon:"roundRect", selectedMode:false },
    grid:{ left:48, right:18, top:34, bottom:74 },
    xAxis:{ type:"category", data:cats, axisLine:{lineStyle:{color:"#E5E8EF"}},
            axisTick:{show:false},
            axisLabel:{color:"#8A92A6",fontSize:10,interval:0,rotate:38} },
    yAxis:{ type:"value", splitLine:{lineStyle:{color:"#F0F2F6"}},
            axisLabel:{color:"#8A92A6",fontSize:11,formatter:"${value}"} },
    series: [{ type:"bar", data:vals, barMaxWidth:22 }].concat(legendStubs)
  }, true);
}

function drawScatter(){
  const latest = paidLatest();
  const series = ORDER.filter(p=>latest.some(d=>d.provider===p)).map(p=>{
    const rows = latest.filter(d=>d.provider===p);
    return {
      name:shortName(p), type:"scatter", symbolSize:15,
      itemStyle:{ color:COLORS[p], opacity:.9 },
      data: rows.map(r=>({ value:[r.input,r.output], name:r.model_short, series:r.series,
                           symbol: SYMBOL_VARIANT[seriesIdx(p,r.series)]||"circle" }))
    };
  });
  scatterChart.setOption({
    tooltip:{ trigger:"item", formatter:p=>`${p.data.name}<br/><span style="color:#AEB4C2">${p.data.series||""}</span><br/>输入 $${p.value[0]}/M<br/>输出 $${p.value[1]}/M` },
    legend:{ type:"scroll", bottom:0, textStyle:{fontSize:11,color:"#5A6273"}, icon:"circle" },
    grid:{ left:54, right:20, top:16, bottom:44 },
    xAxis:{ type:"log", name:"输入 $/M（对数）", nameTextStyle:{color:"#AEB4C2",fontSize:11},
            min:0.1, splitLine:{lineStyle:{color:"#F0F2F6"}}, axisLabel:{color:"#8A92A6",fontSize:11,formatter:"${value}"} },
    yAxis:{ type:"log", name:"输出 $/M（对数）", nameTextStyle:{color:"#AEB4C2",fontSize:11},
            min:0.1, splitLine:{lineStyle:{color:"#F0F2F6"}}, axisLabel:{color:"#8A92A6",fontSize:11,formatter:"${value}"} },
    series
  });
}

function drawRank(){
  const latest = paidLatest().map(d=>({ ...d, total:(d.input||0)+(d.output||0) }))
                 .sort((a,b)=>b.total-a.total);
  const cats = latest.map(d=>d.model_short);
  const vals = latest.map(d=>({
    value: Number(d.total.toFixed(3)),
    itemStyle:{ color: COLORS[d.provider]||"#9DB0FF", borderRadius:[0,4,4,0] }
  }));
  rankChart.setOption({
    tooltip:{ trigger:"axis", axisPointer:{type:"shadow"}, valueFormatter:v=>"$"+Number(v).toFixed(2)+"/M" },
    grid:{ left:130, right:48, top:8, bottom:8 },
    xAxis:{ type:"value", splitLine:{lineStyle:{color:"#F0F2F6"}},
            axisLabel:{color:"#8A92A6",fontSize:11,formatter:"${value}"} },
    yAxis:{ type:"category", data:cats, axisLine:{lineStyle:{color:"#E5E8EF"}},
            axisTick:{show:false}, axisLabel:{color:"#5A6273",fontSize:11} },
    series:[{ type:"bar", data:vals, barMaxWidth:18,
      label:{show:true,position:"right",formatter:p=>"$"+p.value,color:"#8A92A6",fontSize:11} }]
  });
}

// ================= 明细表 =================
let sortKey="input", sortAsc=true;
function renderTable(){
  const rows = DATA.filter(d=>d.date===LATEST).map(d=>{
    const dl = deltaOf(d.model_id,"input");
    return { ...d, _chg: dl };
  });
  rows.sort((a,b)=>{
    let va,vb;
    if(sortKey==="chg"){ va=a._chg?a._chg.pct:-1e9; vb=b._chg?b._chg.pct:-1e9; }
    else { va=a[sortKey]; vb=b[sortKey]; }
    if(va==null) va=-1e9; if(vb==null) vb=-1e9;
    if(typeof va==="string") return sortAsc? va.localeCompare(vb,"zh"):vb.localeCompare(va,"zh");
    return sortAsc? va-vb : vb-va;
  });
  document.getElementById("tbody").innerHTML = rows.map(d=>{
    let chg='<span class="chg flat">—</span>';
    if(d._chg){
      if(d._chg.dir==="flat") chg='<span class="chg flat">持平</span>';
      else chg=`<span class="chg ${d._chg.dir}">${d._chg.dir==="up"?"▲":"▼"} ${Math.abs(d._chg.pct).toFixed(1)}%</span>`;
    }
    return `<tr>
      <td><span class="dot" style="background:${COLORS[d.provider]||'#ccc'}"></span>${shortName(d.provider)}</td>
      <td>${d.model_short}</td>
      <td class="num">${fmt(d.input)}</td>
      <td class="num">${fmt(d.output)}</td>
      <td class="num">${chg}</td>
    </tr>`;
  }).join("");
  document.querySelectorAll("#tbl thead th").forEach(th=>{
    const k=th.dataset.key; th.querySelector(".arrow")?.remove();
    if(k===sortKey){ const s=document.createElement("span"); s.className="arrow"; s.textContent=sortAsc?"▲":"▼"; th.appendChild(s); }
  });
}

// ================= 初始化 =================
function syncScaleEnabled(){
  const disabled = state.mode==="index";
  document.querySelectorAll("#seg_scale button").forEach(b=>{
    b.disabled=disabled; b.style.opacity=disabled?0.4:1; b.style.cursor=disabled?"not-allowed":"pointer";
  });
}

function init(){
  trendChart   = echarts.init(document.getElementById("c_trend"));
  compareChart = echarts.init(document.getElementById("c_compare"));
  scatterChart = echarts.init(document.getElementById("c_scatter"));
  rankChart    = echarts.init(document.getElementById("c_rank"));

  renderKPI();
  renderChips();
  const provCount = new Set(DATA.map(d=>d.provider)).size;
  const modelCount = new Set(DATA.filter(d=>d.date===LATEST).map(d=>d.model_id)).size;
  document.getElementById("sub_summary").textContent =
    provCount+" 家厂商 · "+modelCount+" 条主推产品线";
  syncScaleEnabled();

  drawTrend(); renderChangeBar(); drawCompare(); drawScatter(); drawRank();
  renderTable();

  document.querySelectorAll("#seg_range button").forEach(b=>{
    b.onclick=()=>{
      document.querySelectorAll("#seg_range button").forEach(x=>x.classList.remove("active"));
      b.classList.add("active"); state.range=b.dataset.range;
      drawTrend(); renderChangeBar();
    };
  });
  document.querySelectorAll("#seg_mode button").forEach(b=>{
    b.onclick=()=>{
      document.querySelectorAll("#seg_mode button").forEach(x=>x.classList.remove("active"));
      b.classList.add("active"); state.mode=b.dataset.mode;
      syncScaleEnabled(); drawTrend();
    };
  });
  document.querySelectorAll("#seg_scale button").forEach(b=>{
    b.onclick=()=>{
      if(state.mode==="index") return;
      document.querySelectorAll("#seg_scale button").forEach(x=>x.classList.remove("active"));
      b.classList.add("active"); state.scale=b.dataset.scale; drawTrend();
    };
  });
  document.querySelectorAll("#seg_io button").forEach(b=>{
    b.onclick=()=>{
      document.querySelectorAll("#seg_io button").forEach(x=>x.classList.remove("active"));
      b.classList.add("active"); state.io=b.dataset.io;
      drawTrend(); renderChangeBar();
    };
  });
  document.querySelectorAll("#tbl thead th").forEach(th=>{
    th.onclick=()=>{ const k=th.dataset.key;
      if(sortKey===k) sortAsc=!sortAsc; else {sortKey=k; sortAsc=(k==="input"||k==="output");}
      renderTable();
    };
  });

  window.addEventListener("resize",()=>{
    [trendChart,compareChart,scatterChart,rankChart].forEach(c=>c&&c.resize());
  });
}
init();
</script>
</body>
</html>"""

    html = (html
            .replace("__ECHARTS_CDN__", ECHARTS_CDN)
            .replace("__GENERATED_AT__", meta["generated_at"])
            .replace("__DATA_JSON__", data_json)
            .replace("__META_JSON__", meta_json))
    return html


def main():
    print("读取数据...")
    df = load_data()
    print(f"已加载 {len(df)} 行数据，日期范围：{df['抓取日期'].min()} ~ {df['抓取日期'].max()}")

    print("生成仪表盘...")
    html = build_html(df)
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)

    size_kb = round(os.path.getsize(OUTPUT_HTML) / 1024, 1)
    print(f"仪表盘已生成：{OUTPUT_HTML}（{size_kb} KB，ECharts CDN · 浅色商务风 v3）")

    # 同时导出 JSON 供侧边栏 artifact 使用
    json_path = os.path.join(SCRIPT_DIR, "prices_data.json")
    payload = {"records": build_records(df), "meta": build_meta(df)}
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False)
    print(f"数据 JSON 已导出：{json_path}")


if __name__ == "__main__":
    main()
