#!/usr/bin/env python3
"""日股復盤儀表板生成器

抓東証 17 業種代表個股日線資料,輸出單頁靜態 HTML,包含:
- 大盤概況(日經平均、TOPIX)
- 17 業種漲跌排行(等權重、中位數)
- 每板塊證據瓦片:放量突破、強勢、弱勢 TOP 3

相依: yfinance, pandas, numpy
"""
import json
import sys
import warnings
from datetime import datetime, timezone, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
import yfinance as yf

warnings.filterwarnings('ignore')

HERE = Path(__file__).parent
SECTORS_JSON = HERE / 'sectors.json'
OUTPUT_HTML = HERE / 'jp_dashboard.html'
OUTPUT_EMBED = HERE / 'jp_dashboard_embed.html'
JST = timezone(timedelta(hours=9))


def load_sectors():
    with open(SECTORS_JSON, 'r', encoding='utf-8') as f:
        return json.load(f)


def yf_ticker(code: str) -> str:
    if code.startswith('^') or code.endswith('.T'):
        return code
    return f"{code}.T"


def build_ticker_map(sectors: dict):
    """Return {yf_ticker: (sector, name)}. Skips _meta and 指数 keys.
    Dedupes across sectors (last-write wins, which is rare in curated list)."""
    m = {}
    for sector, stocks in sectors.items():
        if sector.startswith('_') or sector == '指数':
            continue
        for stk in stocks:
            m[yf_ticker(stk['code'])] = (sector, stk['name'])
    return m


def fetch_indices(index_list):
    """Fetch index values one-by-one (small list, handles ^ tickers)."""
    out = []
    for item in index_list:
        code = item['code']
        try:
            hist = yf.Ticker(code).history(period='10d', auto_adjust=False)
            if len(hist) < 2:
                continue
            last = float(hist['Close'].iloc[-1])
            prev = float(hist['Close'].iloc[-2])
            out.append({
                'name': item['name'],
                'code': code,
                'close': last,
                'chg': (last - prev) / prev * 100,
                'date': hist.index[-1].date(),
            })
        except Exception as e:
            print(f"  [index err] {code}: {e}", file=sys.stderr)
    return out


def fetch_all(tickers):
    """Batch-download ~90d OHLCV. Returns multi-index DataFrame."""
    print(f"Downloading {len(tickers)} tickers (yfinance batch)...")
    df = yf.download(
        tickers=' '.join(tickers),
        period='120d',
        group_by='ticker',
        threads=True,
        progress=False,
        auto_adjust=False,
    )
    return df


def stock_metrics(df, ticker):
    try:
        sub = df[ticker].dropna(subset=['Close'])
    except (KeyError, ValueError):
        return None
    if len(sub) < 21:
        return None
    close = sub['Close'].astype(float)
    vol = sub['Volume'].astype(float)
    last_close = float(close.iloc[-1])
    prev_close = float(close.iloc[-2])
    if prev_close <= 0:
        return None
    chg = (last_close - prev_close) / prev_close * 100
    last_vol = float(vol.iloc[-1])
    avg_vol_20 = float(vol.iloc[-21:-1].mean())
    vol_ratio = last_vol / avg_vol_20 if avg_vol_20 > 0 else 0.0
    high_20 = float(close.iloc[-21:-1].max())
    ma20 = float(close.iloc[-20:].mean())
    # RSI14
    delta = close.diff()
    gain = delta.clip(lower=0).rolling(14).mean()
    loss = -delta.clip(upper=0).rolling(14).mean()
    rs = gain / loss.replace(0, np.nan)
    rsi_series = 100 - 100 / (1 + rs)
    rsi = float(rsi_series.iloc[-1]) if pd.notna(rsi_series.iloc[-1]) else None
    # 52w-ish low (using all available data, typically ~120d)
    low_min = float(close.min())
    near_low = (last_close - low_min) / low_min * 100 if low_min > 0 else 0
    return {
        'chg': chg,
        'close': last_close,
        'vol_ratio': vol_ratio,
        'broke_20d_high': last_close > high_20,
        'ma20': ma20,
        'above_ma20': last_close > ma20,
        'rsi': rsi,
        'near_low_pct': near_low,
        'last_vol': last_vol,
        'date': sub.index[-1].date(),
    }


def build_sector_report(ticker_map, df):
    by_sector = {}
    for ticker, (sector, name) in ticker_map.items():
        m = stock_metrics(df, ticker)
        if m is None:
            continue
        by_sector.setdefault(sector, []).append({
            'ticker': ticker,
            'code': ticker.replace('.T', ''),
            'name': name,
            **m,
        })
    report = []
    for sector, stocks in by_sector.items():
        if not stocks:
            continue
        chgs = np.array([s['chg'] for s in stocks])
        breakouts = sorted(
            [s for s in stocks if s['vol_ratio'] >= 2.0 and s['broke_20d_high']],
            key=lambda x: x['vol_ratio'], reverse=True,
        )[:3]
        bottom_volume = sorted(
            [s for s in stocks if s['vol_ratio'] >= 2.0 and s['near_low_pct'] <= 10],
            key=lambda x: x['vol_ratio'], reverse=True,
        )[:3]
        top_gainers = sorted(stocks, key=lambda x: x['chg'], reverse=True)[:3]
        top_losers = sorted(stocks, key=lambda x: x['chg'])[:3]
        report.append({
            'sector': sector,
            'avg_chg': float(chgs.mean()),
            'median_chg': float(np.median(chgs)),
            'advancers': int((chgs > 0).sum()),
            'decliners': int((chgs < 0).sum()),
            'n': len(stocks),
            'volume_breakouts': breakouts,
            'bottom_volume': bottom_volume,
            'top_gainers': top_gainers,
            'top_losers': top_losers,
        })
    report.sort(key=lambda x: x['avg_chg'], reverse=True)
    return report


# ---------- HTML rendering ----------

def _stock_row(s, highlight_key=None):
    """Render a single stock row inside an evidence tile."""
    chg_cls = 'up' if s['chg'] > 0 else ('down' if s['chg'] < 0 else '')
    chg_str = f"{s['chg']:+.2f}%"
    extra = ''
    if highlight_key == 'vol':
        extra = f"<span class='extra'>量 {s['vol_ratio']:.1f}×</span>"
    elif highlight_key == 'rsi' and s.get('rsi') is not None:
        extra = f"<span class='extra'>RSI {s['rsi']:.0f}</span>"
    return (
        f"<div class='stock-row'>"
        f"<span class='code'>{s['code']}</span>"
        f"<span class='name'>{s['name']}</span>"
        f"<span class='chg {chg_cls}'>{chg_str}</span>"
        f"{extra}"
        f"</div>"
    )


def _tile(title, stocks, key=None, empty_text='— 無 —'):
    if not stocks:
        body = f"<div class='tile-empty'>{empty_text}</div>"
    else:
        body = ''.join(_stock_row(s, highlight_key=key) for s in stocks)
    return (
        f"<div class='tile'>"
        f"<div class='tile-title'>{title}</div>"
        f"{body}"
        f"</div>"
    )


def _sector_card(s):
    chg_cls = 'up' if s['avg_chg'] > 0 else ('down' if s['avg_chg'] < 0 else '')
    width = min(abs(s['avg_chg']) * 15, 100)
    bar_side = 'bar-up' if s['avg_chg'] > 0 else 'bar-down'
    return (
        f"<details class='sector-card' open>"
        f"<summary class='sector-summary'>"
        f"<div class='sec-head'>"
        f"<span class='sec-name'>{s['sector']}</span>"
        f"<span class='sec-chg {chg_cls}'>{s['avg_chg']:+.2f}%</span>"
        f"</div>"
        f"<div class='sec-meta'>"
        f"成分 {s['n']} 檔 | 漲 {s['advancers']} 跌 {s['decliners']} | 中位 {s['median_chg']:+.2f}%"
        f"</div>"
        f"<div class='sec-bar'><div class='{bar_side}' style='width:{width}%'></div></div>"
        f"</summary>"
        f"<div class='tiles'>"
        f"{_tile('🔥 放量突破 TOP3', s['volume_breakouts'], key='vol', empty_text='今日無放量創新高')}"
        f"{_tile('📈 漲幅 TOP3', s['top_gainers'])}"
        f"{_tile('📉 跌幅 TOP3', s['top_losers'])}"
        f"{_tile('💡 底部放量 TOP3', s['bottom_volume'], key='vol', empty_text='今日無底部放量')}"
        f"</div>"
        f"</details>"
    )


def _index_card(idx):
    chg_cls = 'up' if idx['chg'] > 0 else ('down' if idx['chg'] < 0 else '')
    if idx.get('is_breadth'):
        # breadth card: 顯示等權%、漲跌家數比
        adv = idx.get('advancers', 0)
        dec = idx.get('decliners', 0)
        return (
            f"<div class='index-card breadth'>"
            f"<div class='idx-name'>{idx['name']}</div>"
            f"<div class='idx-close {chg_cls}'>{idx['chg']:+.2f}%</div>"
            f"<div class='idx-chg'>漲 <span class='up'>{adv}</span> / 跌 <span class='down'>{dec}</span></div>"
            f"</div>"
        )
    return (
        f"<div class='index-card'>"
        f"<div class='idx-name'>{idx['name']}</div>"
        f"<div class='idx-close'>{idx['close']:,.2f}</div>"
        f"<div class='idx-chg {chg_cls}'>{idx['chg']:+.2f}%</div>"
        f"</div>"
    )


EMBED_STYLE = """<style>
  .jpr-root {
    --jpr-bg: #fafaf7;
    --jpr-card: #ffffff;
    --jpr-border: #e5e2dc;
    --jpr-text: #2a2a2a;
    --jpr-muted: #888;
    --jpr-up: #c0392b;
    --jpr-down: #27874b;
    background: var(--jpr-bg);
    color: var(--jpr-text);
    font-family: -apple-system, "Hiragino Sans", "Noto Sans JP", "Noto Sans TC", sans-serif;
    font-size: 14px;
    line-height: 1.5;
    padding: 16px;
    box-sizing: border-box;
  }
  .jpr-root * { box-sizing: border-box; }
  .jpr-root header { margin-bottom: 16px; }
  .jpr-root h1 { font-size: 18px; margin: 0 0 4px; font-weight: 600; }
  .jpr-root .sub { color: var(--jpr-muted); font-size: 12px; }
  .jpr-root .indices { display: flex; gap: 10px; margin: 12px 0 20px; flex-wrap: wrap; }
  .jpr-root .index-card { flex: 1 1 140px; background: var(--jpr-card); border: 1px solid var(--jpr-border); border-radius: 6px; padding: 10px 12px; }
  .jpr-root .idx-name { font-size: 12px; color: var(--jpr-muted); }
  .jpr-root .idx-close { font-size: 18px; font-weight: 600; margin-top: 2px; }
  .jpr-root .idx-chg { font-size: 13px; margin-top: 2px; }
  .jpr-root .sector-card { background: var(--jpr-card); border: 1px solid var(--jpr-border); border-radius: 6px; margin-bottom: 10px; overflow: hidden; }
  .jpr-root .sector-summary { cursor: pointer; list-style: none; padding: 12px 14px; user-select: none; }
  .jpr-root .sector-summary::-webkit-details-marker { display: none; }
  .jpr-root .sec-head { display: flex; justify-content: space-between; align-items: baseline; gap: 12px; }
  .jpr-root .sec-name { font-size: 15px; font-weight: 600; }
  .jpr-root .sec-chg { font-size: 15px; font-weight: 600; font-variant-numeric: tabular-nums; }
  .jpr-root .sec-meta { font-size: 11px; color: var(--jpr-muted); margin-top: 2px; }
  .jpr-root .sec-bar { height: 3px; background: #f0ede8; margin-top: 8px; border-radius: 2px; overflow: hidden; }
  .jpr-root .bar-up { background: var(--jpr-up); height: 100%; border-radius: 2px; }
  .jpr-root .bar-down { background: var(--jpr-down); height: 100%; border-radius: 2px; }
  .jpr-root .tiles { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 8px; padding: 0 14px 14px; }
  .jpr-root .tile { background: #fbfaf6; border: 1px solid #ece8df; border-radius: 4px; padding: 8px 10px; }
  .jpr-root .tile-title { font-size: 12px; color: var(--jpr-muted); margin-bottom: 4px; font-weight: 500; }
  .jpr-root .tile-empty { font-size: 12px; color: #bbb; padding: 4px 0; }
  .jpr-root .stock-row { display: grid; grid-template-columns: 48px 1fr auto auto; gap: 6px; padding: 3px 0; font-size: 13px; align-items: baseline; }
  .jpr-root .stock-row .code { color: var(--jpr-muted); font-variant-numeric: tabular-nums; font-size: 12px; }
  .jpr-root .stock-row .name { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .jpr-root .stock-row .chg { font-variant-numeric: tabular-nums; font-weight: 500; }
  .jpr-root .stock-row .extra { font-size: 11px; color: var(--jpr-muted); min-width: 52px; text-align: right; }
  .jpr-root .up { color: var(--jpr-up); }
  .jpr-root .down { color: var(--jpr-down); }
  .jpr-root footer { margin-top: 24px; padding-top: 12px; border-top: 1px solid var(--jpr-border); color: var(--jpr-muted); font-size: 11px; }
</style>"""


def render_embed(indices, report):
    """嵌入 WordPress 用的版本 — iframe srcdoc 真隔離,避免主題 CSS 污染。

    WP KSES 會過濾 <script> 和 inline style,所以用 HTML width/height
    屬性 + 固定高度(over-sized,確保所有板塊展開都塞得下)。"""
    standalone = render_html(indices, report)
    escaped = (standalone
               .replace('&', '&amp;')
               .replace('"', '&quot;')
               .replace("'", '&#39;'))
    # 17 業種全展開約 6500-9500px,抓 12000 保險。iframe 多餘高度是空白,比 scroll 好。
    return f'<iframe srcdoc="{escaped}" width="100%" height="12000" frameborder="0"></iframe>'


def render_html(indices, report):
    gen_ts = datetime.now(JST).strftime('%Y-%m-%d %H:%M JST')
    data_date = indices[0]['date'].strftime('%Y-%m-%d') if indices else '—'
    idx_html = ''.join(_index_card(i) for i in indices)
    sec_html = ''.join(_sector_card(s) for s in report)

    return f"""<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>日股復盤 {data_date}</title>
<style>
  :root {{
    --bg: #fafaf7;
    --card: #ffffff;
    --border: #e5e2dc;
    --text: #2a2a2a;
    --muted: #888;
    --up: #c0392b;
    --down: #27874b;
    --accent: #1a4d8c;
  }}
  * {{ box-sizing: border-box; }}
  body {{
    margin: 0;
    padding: 16px;
    font-family: -apple-system, "Hiragino Sans", "Noto Sans JP", "Noto Sans TC", sans-serif;
    background: var(--bg);
    color: var(--text);
    font-size: 14px;
    line-height: 1.5;
  }}
  header {{ margin-bottom: 16px; }}
  h1 {{
    font-size: 18px;
    margin: 0 0 4px;
    font-weight: 600;
    letter-spacing: 0.02em;
  }}
  .sub {{ color: var(--muted); font-size: 12px; }}
  .indices {{
    display: flex;
    gap: 10px;
    margin: 12px 0 20px;
    flex-wrap: wrap;
  }}
  .index-card {{
    flex: 1 1 140px;
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 10px 12px;
  }}
  .idx-name {{ font-size: 12px; color: var(--muted); }}
  .idx-close {{ font-size: 18px; font-weight: 600; margin-top: 2px; }}
  .idx-chg {{ font-size: 13px; margin-top: 2px; }}
  .sector-card {{
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 6px;
    margin-bottom: 10px;
    overflow: hidden;
  }}
  .sector-summary {{
    cursor: pointer;
    list-style: none;
    padding: 12px 14px;
    user-select: none;
  }}
  .sector-summary::-webkit-details-marker {{ display: none; }}
  .sec-head {{
    display: flex;
    justify-content: space-between;
    align-items: baseline;
    gap: 12px;
  }}
  .sec-name {{ font-size: 15px; font-weight: 600; }}
  .sec-chg {{ font-size: 15px; font-weight: 600; font-variant-numeric: tabular-nums; }}
  .sec-meta {{ font-size: 11px; color: var(--muted); margin-top: 2px; }}
  .sec-bar {{
    height: 3px;
    background: #f0ede8;
    margin-top: 8px;
    border-radius: 2px;
    overflow: hidden;
    position: relative;
  }}
  .bar-up {{ background: var(--up); height: 100%; border-radius: 2px; }}
  .bar-down {{ background: var(--down); height: 100%; border-radius: 2px; }}
  .tiles {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
    gap: 8px;
    padding: 0 14px 14px;
  }}
  .tile {{
    background: #fbfaf6;
    border: 1px solid #ece8df;
    border-radius: 4px;
    padding: 8px 10px;
  }}
  .tile-title {{
    font-size: 12px;
    color: var(--muted);
    margin-bottom: 4px;
    font-weight: 500;
  }}
  .tile-empty {{ font-size: 12px; color: #bbb; padding: 4px 0; }}
  .stock-row {{
    display: grid;
    grid-template-columns: 48px 1fr auto auto;
    gap: 6px;
    padding: 3px 0;
    font-size: 13px;
    align-items: baseline;
  }}
  .stock-row .code {{ color: var(--muted); font-variant-numeric: tabular-nums; font-size: 12px; }}
  .stock-row .name {{ overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
  .stock-row .chg {{ font-variant-numeric: tabular-nums; font-weight: 500; }}
  .stock-row .extra {{ font-size: 11px; color: var(--muted); min-width: 52px; text-align: right; }}
  .up {{ color: var(--up); }}
  .down {{ color: var(--down); }}
  footer {{
    margin-top: 24px;
    padding-top: 12px;
    border-top: 1px solid var(--border);
    color: var(--muted);
    font-size: 11px;
  }}
</style>
</head>
<body>
<header>
  <h1>日股復盤・板塊地圖</h1>
  <div class="sub">資料日 {data_date} | 產生時間 {gen_ts}</div>
  <div class="indices">{idx_html}</div>
</header>
<main>
  {sec_html}
</main>
<footer>
  資料來源 Yahoo Finance。板塊漲跌為該板塊成分股等權重日漲幅平均。放量突破 = 成交量 ≥ 20 日均量 2 倍 且 收盤創 20 日新高。底部放量 = 放量且價位距區間低點 10% 以內。僅供參考,非投資建議。
</footer>
</body>
</html>
"""


def synthesize_breadth(report):
    """所有成分股等權平均 + 漲跌家數 — 當 N225 / TOPIX 的對照組。"""
    all_chgs = []
    advancers = decliners = unchanged = 0
    last_date = None
    for sec in report:
        for src in (sec['top_gainers'], sec['top_losers']):
            for s in src:
                last_date = s.get('date', last_date)
        # use full sector lists via re-traverse: top_gainers/losers only have 3 each
    # Better: compute from advancers/decliners totals already in report
    for sec in report:
        advancers += sec['advancers']
        decliners += sec['decliners']
        # avg_chg already weighted within sector; recompute total via n*avg_chg
        all_chgs.extend([sec['avg_chg']] * sec['n'])
    if not all_chgs:
        return None
    import numpy as np
    return {
        'name': '成分股等權平均',
        'code': 'BREADTH',
        'close': float(np.mean(all_chgs)),  # actually % change, label as such in card
        'chg': float(np.mean(all_chgs)),
        'date': last_date or datetime.now(JST).date(),
        'is_breadth': True,
        'advancers': advancers,
        'decliners': decliners,
    }


def main():
    sectors = load_sectors()
    ticker_map = build_ticker_map(sectors)
    print(f"Sectors: {len([k for k in sectors if not k.startswith('_') and k != '指数'])}, tickers: {len(ticker_map)}")
    idx_data = fetch_indices(sectors.get('指数', []))
    df = fetch_all(list(ticker_map.keys()))
    report = build_sector_report(ticker_map, df)
    ok = sum(r['n'] for r in report)
    print(f"Processed {ok}/{len(ticker_map)} tickers across {len(report)} sectors")
    breadth = synthesize_breadth(report)
    if breadth:
        idx_data.append(breadth)
    html = render_html(idx_data, report)
    OUTPUT_HTML.write_text(html, encoding='utf-8')
    print(f"✓ Wrote {OUTPUT_HTML}")
    embed = render_embed(idx_data, report)
    OUTPUT_EMBED.write_text(embed, encoding='utf-8')
    print(f"✓ Wrote {OUTPUT_EMBED}")


if __name__ == '__main__':
    main()
