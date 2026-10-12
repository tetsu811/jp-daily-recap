import pandas as pd
import market_data as m

def test_holiday_and_closing_auction():
    assert m.expected_session('2026-10-12T18:00:00+09:00') == '2026-10-09'
    assert m.expected_session('2026-10-13T15:40:00+09:00') == '2026-10-09'
    assert m.expected_session('2026-10-13T16:00:00+09:00') == '2026-10-13'

def bars(dates, closes):
    return pd.DataFrame([[v,v,v,v,100] for v in closes],index=pd.to_datetime(dates),columns=m.COLS)

def test_no_forward_fill_or_intraday():
    f=bars(['2026-10-08','2026-10-09','2026-10-13'],[100,float('nan'),101])
    result=m.clean(f,'2026-10-09')
    assert list(result.index)==[pd.Timestamp('2026-10-08')]

def test_split_basis_not_spliced():
    old=bars(['2026-10-07','2026-10-08'],[200,202])
    new=bars(['2026-10-08','2026-10-09'],[101,102])
    assert m.merge_history(old,new).equals(new)

def test_kabutan_uses_dated_history_not_header():
    html='''<h1>2026/10/12</h1><div id="stock_kabuka_table"><table><tbody><tr><th><time datetime="2026-10-09">26/10/09</time></th><td>100</td><td>102</td><td>99</td><td>101</td><td>1</td><td>1%</td><td>1,234</td></tr></tbody></table></div>'''
    f=m.parse_kabutan(html,'2026-10-09')
    assert f.index[-1]==pd.Timestamp('2026-10-09')
    assert f.iloc[-1].Volume==1234

def test_future_delistings_remain_active(monkeypatch,tmp_path):
    monkeypatch.setattr(m,'HERE',tmp_path)
    monkeypatch.setattr(m,'get_html',lambda _: '')
    monkeypatch.setattr(m,'parse_delistings',lambda _: {'1111.T':{'date':'2026-10-09'},'2222.T':{'date':'2026-10-13'}})
    active, excluded=m.active_tickers({'1111.T':1,'2222.T':2},'2026-10-09')
    assert list(active)==['2222.T']
    assert list(excluded)==['1111.T']

def test_stale_primary_uses_real_backup(monkeypatch,tmp_path):
    monkeypatch.setattr(m,'HERE',tmp_path)
    dates=m.xcals.get_calendar('XTKS').sessions_in_range('2026-08-01','2026-10-09')
    fresh=bars(dates,[100+i for i in range(len(dates))])
    stale=fresh.iloc[:-1]
    monkeypatch.setattr(m.yf,'download',lambda *a,**kw: pd.concat({'7203.T':stale},axis=1))
    monkeypatch.setattr(m,'kabutan_history',lambda *a,**kw:fresh)
    monkeypatch.delenv('JP_FORCE_BACKUP',raising=False)
    frame,details=m.fetch_bars(['7203.T'],'2026-10-09')
    assert details['7203.T']['source']=='Kabutan'
    assert details['7203.T']['status']=='ok'
    assert frame['7203.T'].index[-1]==pd.Timestamp('2026-10-09')

def test_all_sources_stale_are_not_success(monkeypatch,tmp_path):
    monkeypatch.setattr(m,'HERE',tmp_path)
    dates=m.xcals.get_calendar('XTKS').sessions_in_range('2026-08-01','2026-10-08')
    stale=bars(dates,[100+i for i in range(len(dates))])
    monkeypatch.setattr(m.yf,'download',lambda *a,**kw: pd.concat({'7203.T':stale},axis=1))
    monkeypatch.setattr(m,'kabutan_history',lambda *a,**kw:stale)
    monkeypatch.setattr(m,'get_html',lambda *a: '')
    monkeypatch.delenv('JP_FORCE_BACKUP',raising=False)
    _,details=m.fetch_bars(['7203.T'],'2026-10-09')
    assert details['7203.T']['status']=='unavailable'
    assert details['7203.T']['date']=='2026-10-08'

def test_yahoo_japan_index_has_no_volume_column():
    html='<table><thead><tr><th>日付</th><th>始値</th><th>高値</th><th>安値</th><th>終値</th></tr></thead><tbody><tr><th>2026/10/9</th><td>100</td><td>102</td><td>99</td><td>101</td></tr></tbody></table>'
    f=m.parse_yahoo_japan(html,'2026-10-09',True)
    assert len(f)==1 and f.iloc[-1].Close==101 and f.iloc[-1].Volume==0

def test_yahoo_japan_index_symbol_mapping(monkeypatch):
    urls=[]
    monkeypatch.setattr(m,'get_html',lambda url: urls.append(url) or '')
    m.yahoo_japan_history('^N225','2026-10-09')
    assert urls==['https://finance.yahoo.co.jp/quote/998407.O/history']
