#!/usr/bin/env python3
"""把 jp_dashboard_embed.html 透過 WP REST API 推到 WordPress 頁面。

用法:
    WP_URL=https://tetsu811.com \
    WP_USER=your_username \
    WP_PASS='xxxx xxxx xxxx xxxx xxxx xxxx' \
    WP_PAGE_ID=123 \
    python wp_publish.py

第一次跑 (沒 WP_PAGE_ID):
    腳本會建立一個 draft page,印出 page id,你去 WP admin 確認後再把 id 設進
    WP_PAGE_ID 環境變數,之後就會 update 同一頁。

依賴: requests
"""
import base64
import json
import os
import sys
from pathlib import Path
from urllib import request, error

HERE = Path(__file__).parent
EMBED_FILE = HERE / 'jp_dashboard_embed.html'
PAGE_TITLE = '日股復盤 (auto)'
PAGE_SLUG = 'jp-recap'


def env(name, required=True):
    v = os.environ.get(name)
    if required and not v:
        sys.exit(f"missing env var: {name}")
    return v


def wp_request(method, url, user, password, body=None):
    """REST API call with Application Password auth."""
    auth = base64.b64encode(f"{user}:{password}".encode()).decode()
    headers = {
        'Authorization': f'Basic {auth}',
        'Content-Type': 'application/json; charset=utf-8',
        'User-Agent': 'jp-recap-publisher/1.0',
    }
    data = json.dumps(body).encode() if body else None
    req = request.Request(url, data=data, method=method, headers=headers)
    try:
        with request.urlopen(req, timeout=30) as r:
            return json.loads(r.read())
    except error.HTTPError as e:
        body_txt = e.read().decode(errors='replace')
        sys.exit(f"WP API {method} {url} failed: HTTP {e.code}\n{body_txt}")


def main():
    wp_url = env('WP_URL').rstrip('/')
    user = env('WP_USER')
    password = env('WP_PASS').replace(' ', '')  # WP shows pwd with spaces
    page_id = os.environ.get('WP_PAGE_ID')
    status = os.environ.get('WP_STATUS', 'publish')  # 'draft' / 'private' / 'publish'

    if not EMBED_FILE.exists():
        sys.exit(f"missing {EMBED_FILE} — run generate_jp.py first")
    content = EMBED_FILE.read_text(encoding='utf-8')

    payload = {
        'title': PAGE_TITLE,
        'content': content,
        'status': status,
        'slug': PAGE_SLUG,
    }

    if page_id:
        url = f"{wp_url}/wp-json/wp/v2/pages/{page_id}"
        result = wp_request('POST', url, user, password, payload)
        print(f"✓ updated page id={result['id']} link={result.get('link')}")
    else:
        url = f"{wp_url}/wp-json/wp/v2/pages"
        # First time: create as draft so user can review before going live
        payload['status'] = 'draft'
        result = wp_request('POST', url, user, password, payload)
        pid = result['id']
        print(f"✓ created draft page id={pid}")
        print(f"  link: {result.get('link')}")
        print(f"  → set env: export WP_PAGE_ID={pid}")
        print(f"  → 之後跑這個腳本會 update 此頁,並用 WP_STATUS 控制是否 publish")


if __name__ == '__main__':
    main()
