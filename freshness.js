(async function refreshHealth() {
  const node = document.getElementById('market-health');
  if (!node) return;
  try {
    const response = await fetch('https://tetsu811.github.io/jp-daily-recap/health.json?t=' + Date.now(), {cache:'no-store'});
    if (!response.ok) throw new Error('health unavailable');
    const h = await response.json();
    const overdue = Date.now() - Date.parse(h.checked_at) > 30*60*60*1000;
    const different = node.dataset.session !== h.expected_session;
    if (h.status === 'ok' && different) {
      const url = new URL(location.href);
      if (url.searchParams.get('updated') !== h.expected_session) {
        url.searchParams.set('updated', h.expected_session);
        location.replace(url.href);
        return;
      }
    }
    const failed = h.status !== 'ok' || overdue || different;
    node.style.background = failed ? '#fee2e2' : '#e8f5e9';
    node.textContent = (failed ? '更新異常／待核對，請勿視為最新行情。' : '行情日期已逐檔核對。') +
      '最新完整交易日 ' + h.expected_session + ' ｜ 符合日期 ' + h.available + '/' + h.expected +
      ' 檔 ｜ 檢查時間 ' + h.checked_at + ' ｜ 日線盤後更新' +
      (h.problems?.length ? ' ｜ 未完成：' + h.problems.join('、') : '');
  } catch (e) {
    node.style.background = '#fee2e2';
    node.textContent += ' ｜ 無法取得最新更新狀態，請勿視為已驗證。';
  }
  setTimeout(refreshHealth, 300000);
})();
