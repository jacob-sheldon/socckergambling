"""
Flask web application for soccer betting analysis tool.
Provides web UI and API for match data scraping and Excel generation.
"""
import asyncio
import json
import os
import subprocess
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path

from flask import Flask, request, jsonify, send_file

_here = Path(__file__).resolve().parent
sys.path.insert(0, str(_here))

import browser_bet_scraper as bbs

from browser_bet_scraper import (
    fetch_matches_with_browser,
    fetch_asian_handicap_data,
    fetch_euro_kelly_data,
    create_template_workbook,
    set_column_widths,
    merge_header_cells,
    style_header_rows,
    add_match_data,
    MatchData,
    BASE_URL,
    _get_default_playwright_cache_dir,
)

app = Flask(__name__, template_folder=str(_here / 'templates'))

# ---------------------------------------------------------------------------
# Global progress state (single-user local app, no locking needed)
# ---------------------------------------------------------------------------
_scrape_state: dict = {
    'running': False,
    'stage': 'idle',
    'progress': 0,
    'message': '',
    'matches': [],
    'match_count': 0,
    'result_file': None,
    'error': None,
    'log': [],
}


def _reset_state():
    _scrape_state.update({
        'running': True,
        'stage': 'idle',
        'progress': 0,
        'message': '',
        'matches': [],
        'match_count': 0,
        'result_file': None,
        'error': None,
        'log': [],
    })


def _update(stage: str, progress: int, message: str, **kw):
    _scrape_state['stage'] = stage
    _scrape_state['progress'] = progress
    _scrape_state['message'] = message
    _scrape_state['log'].append(f"[{stage}] {message}")
    for k, v in kw.items():
        _scrape_state[k] = v


# ---------------------------------------------------------------------------
# Browser install state
# ---------------------------------------------------------------------------
_install_state: dict = {
    'installing': False,
    'message': '',
    'done': False,
    'error': None,
}


def _check_browser_installed() -> bool:
    cache = _get_default_playwright_cache_dir()
    if not cache.exists():
        return False
    chromium_dirs = list(cache.glob('chromium-*'))
    return len(chromium_dirs) > 0


def _run_install_browser():
    global _install_state
    _install_state.update({'installing': True, 'message': '正在下载 Chromium 浏览器...', 'done': False, 'error': None})
    try:
        result = subprocess.run(
            [sys.executable, '-m', 'playwright', 'install', 'chromium'],
            capture_output=True, text=True, timeout=300000, cwd=str(_here),
        )
        if result.returncode == 0 and _check_browser_installed():
            _install_state.update({'installing': False, 'message': '安装完成！', 'done': True})
        else:
            _install_state.update({
                'installing': False,
                'message': '安装失败',
                'error': result.stderr or result.stdout or '未知错误',
                'done': True,
            })
    except subprocess.TimeoutExpired:
        _install_state.update({'installing': False, 'message': '安装超时', 'error': '下载超时，请检查网络后重试', 'done': True})
    except Exception as e:
        _install_state.update({'installing': False, 'message': '安装出错', 'error': str(e), 'done': True})


# ---------------------------------------------------------------------------
# Scraping pipeline (runs in background thread)
# ---------------------------------------------------------------------------
def _run_scrape(url: str, headless: bool, max_matches, fetch_asian: bool):
    try:
        _update('starting', 5, '正在启动浏览器...')

        # Step 1 — fetch matches
        _update('loading', 10, '正在加载页面...')
        matches = asyncio.run(fetch_matches_with_browser(url, headless=headless))

        if not matches:
            from browser_bet_scraper import _generate_fallback_matches
            matches = _generate_fallback_matches(5)
            err_detail = bbs.LAST_SCRAPE_ERROR or '未知错误'
            _update('extracting', 35, f'未获取到数据，使用示例数据 ({len(matches)} 场)')
            _scrape_state['log'].append(f'[warning] 爬取失败，错误详情: {err_detail[:500]}')
        else:
            _update('extracting', 40, f'已提取 {len(matches)} 场比赛数据')

        if max_matches is not None:
            matches = matches[:max_matches]

        match_dicts = [m.to_dict() for m in matches]

        # Step 2 — Asian handicap
        if fetch_asian:
            _update('asian_handicap', 50, '正在获取亚洲盘口数据...')
            matches = asyncio.run(fetch_asian_handicap_data(matches, headless=headless))
            match_dicts = [m.to_dict() for m in matches]

            _update('euro_kelly', 70, '正在获取百家欧赔即时凯利数据...')
            matches = asyncio.run(fetch_euro_kelly_data(matches, headless=headless))
            match_dicts = [m.to_dict() for m in matches]

        # Step 3 — Generate Excel
        _update('generating', 85, '正在生成 Excel 模板...')
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"live_betting_template_{ts}.xlsx"
        filepath = str(_here / filename)

        wb, ws = create_template_workbook()
        set_column_widths(ws)
        merge_header_cells(ws)
        style_header_rows(ws)

        row = 3
        for match in matches:
            row += add_match_data(ws, row, match)

        ws.freeze_panes = 'A3'
        wb.save(filepath)

        _update('done', 100, f'完成！共 {len(matches)} 场比赛',
                matches=match_dicts, match_count=len(matches),
                result_file=filepath, running=False)

    except Exception:
        _update('error', 0, '发生错误',
                error=traceback.format_exc(), running=False)


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------
@app.route('/')
def index():
    return send_file(str(_here / 'templates' / 'index.html'))


@app.route('/api/status')
def api_status():
    return jsonify({k: v for k, v in _scrape_state.items() if k != 'log'})


@app.route('/api/log')
def api_log():
    return jsonify({'log': _scrape_state.get('log', [])})


@app.route('/api/scrape', methods=['POST'])
def api_scrape():
    if _scrape_state['running']:
        return jsonify({'error': '已有正在进行的爬取任务'}), 409

    data = request.get_json(silent=True) or {}
    url = data.get('url', BASE_URL)
    headless = data.get('headless', True)
    max_matches = data.get('max_matches', None)
    fetch_asian = data.get('asian_handicap', False)

    _reset_state()

    t = threading.Thread(target=_run_scrape, args=(url, headless, max_matches, fetch_asian), daemon=True)
    t.start()

    return jsonify({'status': 'started'})


@app.route('/api/download')
def api_download():
    path = _scrape_state.get('result_file')
    if not path or not os.path.exists(path):
        return jsonify({'error': '文件不存在或尚未生成'}), 404
    name = os.path.basename(path)
    return send_file(path, as_attachment=True, download_name=name)


@app.route('/api/browser-status')
def api_browser_status():
    return jsonify({
        'installed': _check_browser_installed(),
        'install_state': _install_state,
    })


@app.route('/api/install-browser', methods=['POST'])
def api_install_browser():
    if _install_state['installing']:
        return jsonify({'error': '正在安装中'}), 409
    t = threading.Thread(target=_run_install_browser, daemon=True)
    t.start()
    return jsonify({'status': 'started'})


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main():
    import argparse
    p = argparse.ArgumentParser(description='足球彩票分析工具 - Web 服务')
    p.add_argument('--port', type=int, default=5100, help='端口号 (默认: 5100)')
    p.add_argument('--host', default='127.0.0.1', help='绑定地址 (默认: 127.0.0.1)')
    args = p.parse_args()

    print(f'足球彩票分析工具 Web 服务')
    print(f'访问地址: http://{args.host}:{args.port}')
    app.run(host=args.host, port=args.port, debug=False)


if __name__ == '__main__':
    main()
