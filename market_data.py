"""Dated Japanese daily bars with independent public-source failover.

Yahoo Finance supplies the long lookback; Kabutan supplies missing sessions.
Yahoo Japan's public history table is a third route. A cache is never a fresh
observation unless its actual trading date matches the required session.
"""
import gzip
import json
import math
import os
import re
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta, timezone
from pathlib import Path

import exchange_calendars as xcals
import pandas as pd
import requests
import yfinance as yf
from bs4 import BeautifulSoup

HERE = Path(__file__).resolve().parent
JST = timezone(timedelta(hours=9))
COLS = ['Open', 'High', 'Low', 'Close', 'Volume']
DELISTED_URL = 'https://www.jpx.co.jp/english/listing/stocks/delisted/'


def expected_session(now=None):
    now = pd.Timestamp(now or datetime.now(timezone.utc))
    if now.tzinfo is None:
        raise ValueError('An explicit timezone is required')
    today = now.tz_convert('Asia/Tokyo').date()
    cal = xcals.get_calendar('XTKS')
    sessions = cal.sessions_in_range(today - timedelta(days=20), today)
    # Allow the closing auction to settle; never consume an unfinished daily bar.
    closed = [s for s in sessions if cal.session_close(s) + pd.Timedelta(minutes=20) <= now]
    if not closed:
        raise ValueError('No completed Tokyo session in calendar')
    return closed[-1].date().isoformat()


def get_html(url):
    last_error = None
    for attempt in range(2):
        try:
            response = requests.get(url, timeout=(8, 25), headers={'User-Agent': 'Mozilla/5.0'})
            response.raise_for_status()
            response.encoding = 'utf-8'
            return response.text
        except requests.RequestException as exc:
            last_error = exc
            if isinstance(exc, requests.HTTPError) and exc.response is not None and exc.response.status_code in (400, 401, 403, 404, 405):
                raise
            if attempt == 0:
                time.sleep(1)
    raise last_error


def number(text):
    return float(text.strip().replace(',', ''))


def clean(frame, target, is_index=False):
    if frame is None or frame.empty:
        return pd.DataFrame(columns=COLS)
    frame = frame.copy()
    frame.index = pd.to_datetime(frame.index).tz_localize(None).normalize()
    frame = frame[~frame.index.duplicated(keep='last')].sort_index()
    frame = frame.reindex(columns=COLS).apply(pd.to_numeric, errors='coerce')
    frame = frame.loc[frame.index <= pd.Timestamp(target)]
    prices = frame[COLS[:4]]
    valid = prices.notna().all(axis=1) & (prices > 0).all(axis=1)
    valid &= prices.apply(lambda c: c.map(math.isfinite)).all(axis=1)
    valid &= frame['High'] >= prices[['Open', 'Close', 'Low']].max(axis=1)
    valid &= frame['Low'] <= prices[['Open', 'Close', 'High']].min(axis=1)
    if is_index:
        frame['Volume'] = frame['Volume'].fillna(0)
    valid &= frame['Volume'].notna() & (frame['Volume'] >= 0) & frame['Volume'].map(math.isfinite)
    return frame.loc[valid].tail(160)


def parse_kabutan(html, target, is_index=False):
    soup = BeautifulSoup(html, 'html.parser')
    rows = []
    for tr in soup.select('#stock_kabuka_table tbody tr'):
        stamp = tr.select_one('th time[datetime]')
        cells = tr.find_all('td', recursive=False)
        if not stamp or len(cells) < 7:
            continue
        try:
            values = [number(c.get_text()) for c in cells[:4]]
            volume = number(cells[6].get_text()) if cells[6].get_text().strip() not in ('－', '-', '') else (0 if is_index else None)
            rows.append([stamp['datetime'], *values, volume])
        except (ValueError, TypeError):
            continue
    frame = pd.DataFrame(rows, columns=['Date', *COLS]).set_index('Date')
    return clean(frame, target, is_index)


def parse_yahoo_japan(html, target, is_index=False):
    soup = BeautifulSoup(html, 'html.parser')
    rows = []
    for table in soup.find_all('table'):
        if not all(word in table.get_text() for word in ('始値', '高値', '安値', '終値')):
            continue
        for tr in table.select('tbody tr'):
            cells = tr.find_all(['th', 'td'], recursive=False)
            if len(cells) < (5 if is_index else 6):
                continue
            raw_date = cells[0].get_text(strip=True)
            if not re.fullmatch(r'\d{4}/\d{1,2}/\d{1,2}', raw_date):
                continue
            try:
                date = datetime.strptime(raw_date, '%Y/%m/%d').date().isoformat()
                values = [number(c.get_text()) for c in cells[1:5]]
                volume = 0 if is_index else number(cells[5].get_text())
                rows.append([date, *values, volume])
            except ValueError:
                continue
    return clean(pd.DataFrame(rows, columns=['Date', *COLS]).set_index('Date'), target, is_index)


def merge_history(older, newer):
    """Never silently splice histories with different split/price bases."""
    if older.empty:
        return newer
    if newer.empty:
        return older
    overlap = older.index.intersection(newer.index)
    if len(overlap):
        differences = ((older.loc[overlap, 'Close'] / newer.loc[overlap, 'Close']) - 1).abs()
        if (differences > 0.005).any():
            # Prefer the current provider's internally consistent history.
            return newer
    return pd.concat([older, newer]).loc[lambda f: ~f.index.duplicated(keep='last')].sort_index().tail(160)


def kabutan_history(symbol, target, need_history=False):
    code = '0000' if symbol == '^N225' else symbol.removesuffix('.T')
    if not re.fullmatch(r'[0-9A-Z]{4}', code):
        raise ValueError('Unsupported Kabutan symbol')
    url = f'https://kabutan.jp/stock/kabuka?code={code}'
    frame = parse_kabutan(get_html(url), target, symbol.startswith('^'))
    # Public pagination only; enough history when Yahoo and the local cache fail.
    if need_history:
        for page in range(2, 5):
            if len(frame) >= 60:
                break
            more = parse_kabutan(get_html(url + f'&ashi=day&page={page}'), target, symbol.startswith('^'))
            frame = merge_history(frame, more)
    return frame


def yahoo_japan_history(symbol, target, need_history=False):
    code = '998407.O' if symbol == '^N225' else symbol
    url = f'https://finance.yahoo.co.jp/quote/{code}/history'
    frame = parse_yahoo_japan(get_html(url), target, symbol.startswith('^'))
    if need_history:
        for page in range(2, 5):
            if len(frame) >= 60:
                break
            more = parse_yahoo_japan(get_html(url + f'?page={page}'), target, symbol.startswith('^'))
            frame = merge_history(frame, more)
    return frame


def load_cache():
    try:
        with gzip.open(HERE / 'market_cache.json.gz', 'rt') as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def fetch_bars(symbols, target):
    symbols = list(dict.fromkeys(symbols))
    cached = load_cache()
    forced = set(filter(None, os.getenv('JP_FORCE_BACKUP', '').split(',')))
    start = (pd.Timestamp(target) - pd.Timedelta(days=220)).date().isoformat()
    end = (pd.Timestamp(target) + pd.Timedelta(days=1)).date().isoformat()
    primary = {}
    calendar = xcals.get_calendar('XTKS')
    required = calendar.sessions_in_range(pd.Timestamp(target)-pd.Timedelta(days=60), target)[-21:]
    for offset in range(0, len(symbols), 60):
        batch = symbols[offset:offset+60]
        try:
            data = yf.download(batch, start=start, end=end, auto_adjust=False,
                               group_by='ticker', threads=6, progress=False, timeout=20)
            for symbol in batch:
                if symbol in data.columns.get_level_values(0):
                    primary[symbol] = clean(data[symbol], target, symbol.startswith('^'))
        except Exception as exc:
            print(f'Yahoo batch unavailable: {type(exc).__name__}')

    def one(symbol):
        is_index = symbol.startswith('^')
        old = pd.DataFrame(cached.get(symbol, []), columns=['Date', *COLS]).set_index('Date')
        old = clean(old, target, is_index)
        yahoo = primary.get(symbol, pd.DataFrame(columns=COLS))
        frame = merge_history(old, yahoo)
        source = 'Yahoo Finance' if not yahoo.empty and yahoo.index[-1] == pd.Timestamp(target) else 'cache'
        errors = []
        def current(f):
            needed = required[-2:] if is_index else required
            return len(f) >= len(needed) and needed.isin(f.index).all()
        if symbol in forced or not current(frame):
            for name, fetcher in [
                ('Kabutan', lambda: kabutan_history(symbol, target, len(frame) < 21)),
                ('Yahoo Japan', lambda: yahoo_japan_history(symbol, target, len(frame) < 21)),
            ]:
                try:
                    candidate = fetcher()
                    if candidate.empty or candidate.index[-1] != pd.Timestamp(target):
                        raise ValueError('Expected closing session is absent')
                    merged = merge_history(frame, candidate)
                    if not current(merged):
                        # A split-basis mismatch may have discarded the old history.
                        if name == 'Kabutan':
                            merged = kabutan_history(symbol, target, True)
                        else:
                            merged = yahoo_japan_history(symbol, target, True)
                        if not current(merged):
                            raise ValueError('Insufficient consistent history')
                    frame, source = merged, name
                    break
                except Exception as exc:
                    errors.append(f'{name}: {type(exc).__name__}: {str(exc)[:120]}')
        as_of = frame.index[-1].date().isoformat() if not frame.empty else None
        status = 'ok' if current(frame) else 'unavailable'
        if symbol in forced and source not in ('Kabutan', 'Yahoo Japan'):
            status = 'unavailable'
        detail = {'date':as_of, 'source':source, 'status':status, 'bars':len(frame), 'errors':errors}
        print(f'{symbol}: {status} {as_of} via {source}')
        return symbol, frame, detail

    with ThreadPoolExecutor(max_workers=4) as pool:
        results = list(pool.map(one, symbols))
    frames = {s:f for s,f,_ in results}
    details = {s:d for s,_,d in results}
    # Persist only dated, validated observations, including useful partial fetches.
    payload = {s:[[i.date().isoformat(), *map(float, row)] for i,row in f[COLS].iterrows()] for s,f in frames.items()}
    temporary = HERE / 'market_cache.tmp.gz'
    with gzip.open(temporary, 'wt') as f:
        json.dump(payload, f, allow_nan=False)
    temporary.replace(HERE / 'market_cache.json.gz')
    return pd.concat(frames, axis=1), details


def parse_delistings(html):
    rows = {}
    soup = BeautifulSoup(html, 'html.parser')
    for tr in soup.select('table tr'):
        cells = [c.get_text(' ', strip=True) for c in tr.find_all('td', recursive=False)]
        if len(cells) < 5 or not re.fullmatch(r'[0-9A-Z]{4}', cells[2]):
            continue
        try:
            date = datetime.strptime(cells[0].replace('.', ''), '%b %d, %Y').date().isoformat()
        except ValueError:
            continue
        rows[cells[2]+'.T'] = {'date':date,'name':cells[1],'reason':cells[4]}
    if len(rows) < 10:
        raise ValueError('JPX delisting table format was not recognized')
    return rows


def active_tickers(ticker_map, target):
    path = HERE / 'listing_status.json'
    try:
        previous = json.loads(path.read_text()).get('delisted', {}) if path.exists() else {}
        rows = {**previous, **parse_delistings(get_html(DELISTED_URL))}
        payload = {'checked_at':datetime.now(timezone.utc).isoformat(), 'source':DELISTED_URL, 'delisted':rows}
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2))
    except Exception as exc:
        print(f'JPX status fetch: {type(exc).__name__}')
        payload = json.loads(path.read_text()) if path.exists() else {'delisted':{}}
    excluded = {s:d for s,d in payload['delisted'].items() if s in ticker_map and d['date'] <= target}
    return {s:v for s,v in ticker_map.items() if s not in excluded}, excluded
