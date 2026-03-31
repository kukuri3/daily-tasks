#!/usr/bin/env python3
"""
Daily Tasks メニューバー管理コンソール
macOSメニューバーに常駐し、タスクリストの状態表示と操作を提供する。

メール取得は現在 Claude in Chrome (Outlook Web) 経由で行っている。
Microsoft Graph API が利用可能になり次第、スキャンプロンプトを
API呼び出しに置き換える。（generate_html.py, menubar_app.py の変更のみで移行可能）
"""

import json
import re
import subprocess
import threading
import time
import webbrowser
from datetime import datetime, timedelta
from pathlib import Path

import rumps

BASE_DIR = Path(__file__).parent
CACHE_FILE = BASE_DIR / "tasks_cache.json"
LOG_DIR = BASE_DIR / "logs"
GENERATE_SCRIPT = BASE_DIR / "generate_html.py"
GITHUB_PAGES_URL = "https://kukuri3.github.io/daily-tasks/"

# --- メール取得プロンプト (現在: Outlook Web経由 → 将来: Graph API に置換) ---

SCAN_PROMPT = """Outlookの新着メールを確認してタスクリストを更新してください。

【ブラウザ操作ルール】
- navigate, get_page_text, read_page, find → OK
- computer のクリック → リンククリックのみ許可
- computer での文字入力・ペースト、form_input → 禁止

【手順】
1. tasks_cache.json を読み、last_scan 以降の新着のみ対象
2. 受信トレイ + 送信トレイを確認（read_page でメール一覧取得）
3. メルマガ・広告・自動通知を除外、自分宛メモはタスク追加
4. 送信トレイから既存タスクの完結判定を更新
5. tasks_cache.json を更新、python3 {generate_path} を実行

tasks_cache.json: {cache_path}
generate_html.py: {generate_path}
"""

RESCAN_PROMPT = """Outlookメールを過去{weeks}週間分フルスキャンして、タスクリストを作り直してください。

【ブラウザ操作ルール】
- navigate, get_page_text, read_page, find → OK
- computer のクリック → リンククリックのみ許可
- computer での文字入力・ペースト、form_input → 禁止

【手順】
1. tasks_cache.json を {{"last_scan": null, "tasks": []}} にクリア
2. 受信トレイ + 送信トレイを対象期間({since_date}以降)でスキャン
3. メルマガ・広告を除外、自分宛メモはタスク追加
4. 送信トレイから完結判定
5. tasks_cache.json に書き込み、python3 {generate_path} を実行

tasks_cache.json: {cache_path}
generate_html.py: {generate_path}
"""

# --- 進捗推定 ---

PROGRESS_MARKERS = [
    (5,  [r"claude", r"starting", r"セッション"]),
    (10, [r"outlook", r"Outlook", r"tabs_context"]),
    (20, [r"navigate", r"受信トレイ", r"inbox"]),
    (30, [r"read_page", r"メール.*確認", r"get_page_text"]),
    (50, [r"タスク.*抽出", r"extract", r"分析"]),
    (70, [r"tasks_cache", r"キャッシュ", r"json"]),
    (85, [r"generate_html", r"HTML"]),
    (95, [r"git.*push", r"Push", r"GitHub"]),
    (100, [r"完了", r"success", r"deployed"]),
]


def estimate_progress(log_path):
    if not log_path or not Path(log_path).exists():
        return 0
    try:
        with open(log_path, "r", encoding="utf-8", errors="replace") as f:
            content = f.read()
    except Exception:
        return 0
    if not content.strip():
        return 2
    max_pct = 5
    for pct, patterns in PROGRESS_MARKERS:
        if any(re.search(p, content, re.IGNORECASE) for p in patterns):
            max_pct = max(max_pct, pct)
    return min(max_pct, 99)


# --- キャッシュ読み込み ---

def load_cache():
    if not CACHE_FILE.exists():
        return {"last_scan": None, "tasks": []}
    with open(CACHE_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def get_active_count(cache):
    return len([t for t in cache.get("tasks", []) if not t.get("completed", False)])


def get_completed_count(cache):
    return len([t for t in cache.get("tasks", []) if t.get("completed", False)])


def get_last_scan(cache):
    ls = cache.get("last_scan")
    if not ls:
        return "未実行"
    try:
        return datetime.fromisoformat(ls).strftime("%m/%d %H:%M")
    except (ValueError, TypeError):
        return ls


def get_last_build():
    index = BASE_DIR / "repo" / "index.html"
    if index.exists():
        return datetime.fromtimestamp(index.stat().st_mtime).strftime("%m/%d %H:%M")
    return "未生成"


# --- Claude CLI 検出 ---

def find_claude_cli():
    base = Path.home() / "Library" / "Application Support" / "Claude" / "claude-code"
    if not base.exists():
        return None
    for v in sorted(base.iterdir(), reverse=True):
        cli = v / "claude.app" / "Contents" / "MacOS" / "claude"
        if cli.exists():
            return cli
    return None


# --- メニューバーアプリ ---

class DailyTasksApp(rumps.App):
    def __init__(self):
        super().__init__("📝", quit_button=None)

        self.status_item = rumps.MenuItem("読み込み中...")
        self.scan_item = rumps.MenuItem("最終更新: ...")
        self.build_item = rumps.MenuItem("最終ビルド: ...")

        self.update_btn = rumps.MenuItem("📬 Outlookより更新", callback=self.do_update)
        self.rescan_menu = rumps.MenuItem("🔄 再スキャン")
        self.rescan_menu["1週間"] = rumps.MenuItem("1週間", callback=lambda _: self.do_rescan(1))
        self.rescan_menu["2週間"] = rumps.MenuItem("2週間", callback=lambda _: self.do_rescan(2))
        self.rescan_menu["4週間"] = rumps.MenuItem("4週間", callback=lambda _: self.do_rescan(4))
        self.open_btn = rumps.MenuItem("🌐 ページを開く", callback=self.open_page)
        self.log_btn = rumps.MenuItem("📋 ログを表示", callback=self.show_log)
        self.quit_btn = rumps.MenuItem("❌ 終了", callback=self.quit_app)

        self.scan_process = None
        self.scan_log_path = None
        self.scan_active = False

        self.menu = [
            self.status_item,
            self.scan_item,
            self.build_item,
            None,
            self.update_btn,
            self.rescan_menu,
            self.open_btn,
            self.log_btn,
            None,
            self.quit_btn,
        ]
        self.update_status()

    # --- タイマー ---

    @rumps.timer(3)
    def refresh(self, _):
        if self.scan_active and self.scan_log_path:
            pct = estimate_progress(self.scan_log_path)
            bar = "▓" * (pct // 10) + "░" * (10 - pct // 10)
            self.title = f"📝 {pct}%"
            self.update_btn.title = f"📬 {bar} {pct}%"
        elif not self.scan_active:
            self.title = f"📝 {get_active_count(load_cache())}"

    @rumps.timer(60)
    def refresh_status(self, _):
        if not self.scan_active:
            self.update_status()

    # --- ステータス更新 ---

    def update_status(self):
        cache = load_cache()
        active = get_active_count(cache)
        completed = get_completed_count(cache)
        if not self.scan_active:
            self.title = f"📝 {active}"
        self.status_item.title = f"アクティブ: {active}件 ／ 完了: {completed}件"
        self.scan_item.title = f"📡 最終更新: {get_last_scan(cache)}"
        self.build_item.title = f"🔨 ビルド: {get_last_build()}"

    # --- Outlookより更新（差分スキャン） ---

    def do_update(self, _):
        if self.scan_active:
            rumps.notification("Daily Tasks", "実行中", "既に更新処理が実行中です")
            return
        self._start_scan(
            SCAN_PROMPT.format(cache_path=CACHE_FILE, generate_path=GENERATE_SCRIPT),
            "update",
            "Outlookメールをスキャン中...",
        )

    # --- 再スキャン ---

    def do_rescan(self, weeks):
        if self.scan_active:
            rumps.notification("Daily Tasks", "実行中", "既に更新処理が実行中です")
            return
        resp = rumps.alert(
            title=f"再スキャン（過去{weeks}週間）",
            message=f"既存タスクをクリアして再抽出します。続行しますか？",
            ok="実行", cancel="キャンセル",
        )
        if resp != 1:
            return
        since = (datetime.now() - timedelta(weeks=weeks)).strftime("%Y-%m-%d")
        self._start_scan(
            RESCAN_PROMPT.format(weeks=weeks, since_date=since,
                                 cache_path=CACHE_FILE, generate_path=GENERATE_SCRIPT),
            f"rescan_{weeks}w",
            f"過去{weeks}週間分をスキャン中...",
        )

    # --- スキャン共通ロジック ---

    def _start_scan(self, prompt, log_prefix, notify_msg):
        cli = find_claude_cli()
        if not cli:
            rumps.notification("Daily Tasks", "エラー",
                               "Claude Code CLIが見つかりません")
            return

        self.scan_active = True
        self.title = "📝 0%"
        self.update_btn.title = "📬 ░░░░░░░░░░ 0%"
        rumps.notification("Daily Tasks", "更新開始", notify_msg, sound=False)

        log_file = LOG_DIR / f"{log_prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        self.scan_log_path = str(log_file)

        try:
            self.scan_process = subprocess.Popen(
                [str(cli), "--output-format", "stream-json",
                 "--verbose", "--dangerously-skip-permissions", "-p", prompt],
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                cwd=str(BASE_DIR),
            )
            threading.Thread(target=self._stream_and_wait, args=(log_file,), daemon=True).start()
        except Exception as e:
            self._reset_scan_state()
            rumps.notification("Daily Tasks", "起動エラー", str(e)[:200])

    def _stream_and_wait(self, log_file):
        try:
            with open(log_file, "w", encoding="utf-8") as lf:
                for line in iter(self.scan_process.stdout.readline, b""):
                    lf.write(line.decode("utf-8", errors="replace"))
                    lf.flush()
            rc = self.scan_process.wait(timeout=60)
            if rc == 0:
                self.title = "📝 ✅"
                self.update_btn.title = "📬 ▓▓▓▓▓▓▓▓▓▓ 100%"
                time.sleep(1.5)
                rumps.notification("Daily Tasks", "更新完了 ✅",
                                   "タスクリストを更新しました", sound=True)
            else:
                rumps.notification("Daily Tasks", "エラー",
                                   f"終了コード: {rc}", sound=True)
        except subprocess.TimeoutExpired:
            self.scan_process.kill()
            rumps.notification("Daily Tasks", "タイムアウト",
                               "更新が完了しませんでした", sound=True)
        except Exception as e:
            rumps.notification("Daily Tasks", "エラー", str(e)[:200], sound=True)
        finally:
            self._reset_scan_state()

    def _reset_scan_state(self):
        self.scan_active = False
        self.update_btn.title = "📬 Outlookより更新"
        self.scan_process = None
        self.scan_log_path = None
        self.update_status()

    # --- その他 ---

    def open_page(self, _):
        webbrowser.open(GITHUB_PAGES_URL)

    def show_log(self, _):
        logs = sorted(LOG_DIR.glob("*.log"), reverse=True)
        if not logs:
            rumps.notification("Daily Tasks", "ログなし", "まだログがありません")
            return
        subprocess.run(["open", str(logs[0])])

    def quit_app(self, _):
        rumps.quit_application()


if __name__ == "__main__":
    DailyTasksApp().run()
