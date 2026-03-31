#!/usr/bin/env python3
"""
タスクキャッシュからHTMLを生成し、GitHub Pagesにデプロイする。

Usage:
    python generate_html.py              # HTML生成 + git push
    python generate_html.py --dry-run    # HTML生成のみ（pushしない）
    python generate_html.py --add-task   # 対話的にタスクを追加
"""

import json
import os
import subprocess
import sys
from datetime import date, datetime
from pathlib import Path

from jinja2 import Environment, FileSystemLoader

BASE_DIR = Path(__file__).parent
CACHE_FILE = BASE_DIR / "tasks_cache.json"
TEMPLATE_DIR = BASE_DIR / "templates"
REPO_DIR = BASE_DIR / "repo"
OUTPUT_FILE = REPO_DIR / "index.html"


def load_cache():
    """キャッシュファイルを読み込む。存在しなければ空のキャッシュを返す。"""
    if not CACHE_FILE.exists():
        return {"last_scan": None, "tasks": []}
    with open(CACHE_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def save_cache(cache):
    """キャッシュファイルに保存する。"""
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


def generate_html(cache, generated_date=None):
    """キャッシュからHTMLを生成する。"""
    if generated_date is None:
        generated_date = date.today().isoformat()

    # アクティブ/完了を分離
    active_tasks = [t for t in cache["tasks"] if not t.get("completed", False)]
    completed_tasks = [t for t in cache["tasks"] if t.get("completed", False)]

    def task_to_js(t):
        return {
            "id": t["id"],
            "priority": t["priority"],
            "title": t["title"],
            "from": t["from"],
            "fromEmail": t.get("fromEmail", ""),
            "mailDate": t.get("mailDate", ""),
            "mailSubject": t.get("mailSubject", ""),
            "deadline": t["deadline"],
            "urgent": t.get("urgent", False),
            "note": t["note"],
            "summary": t.get("summary", ""),
            "thread_status": t.get("thread_status", "open"),
            "thread_summary": t.get("thread_summary", ""),
            "related_threads": t.get("related_threads", []),
        }

    js_tasks = [task_to_js(t) for t in active_tasks]
    js_completed = [{"id": t["id"], "title": t["title"], "from": t["from"],
                     "thread_status": t.get("thread_status", "resolved"),
                     "thread_summary": t.get("thread_summary", ""),
                     "summary": t.get("summary", "")}
                    for t in completed_tasks]

    tasks_json = json.dumps(js_tasks, ensure_ascii=False, indent=2)
    completed_tasks_json = json.dumps(js_completed, ensure_ascii=False, indent=2)

    # last_scan表示用
    last_scan = cache.get("last_scan")
    if last_scan:
        try:
            dt = datetime.fromisoformat(last_scan)
            last_scan_display = dt.strftime("%Y/%m/%d %H:%M")
        except (ValueError, TypeError):
            last_scan_display = last_scan
    else:
        last_scan_display = "未実行"

    # Jinja2テンプレートでHTML生成
    env = Environment(
        loader=FileSystemLoader(str(TEMPLATE_DIR)),
        autoescape=False,  # JSONをそのまま埋め込むため
    )
    template = env.get_template("index.html.j2")
    build_time = datetime.now().strftime("%Y/%m/%d %H:%M")

    html = template.render(
        generated_date=generated_date,
        tasks_json=tasks_json,
        completed_tasks_json=completed_tasks_json,
        last_scan_display=last_scan_display,
        build_time=build_time,
    )

    # 出力
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"HTML generated: {OUTPUT_FILE}")
    print(f"  Active tasks: {len(active_tasks)}")
    print(f"  Generated date: {generated_date}")
    return OUTPUT_FILE


def setup_git_remote():
    """GITHUB_TOKEN環境変数があればリモートURLにトークンを埋め込む。"""
    token = os.environ.get("GITHUB_TOKEN")
    if not token:
        return
    # 現在のリモートURLを取得
    result = subprocess.run(
        ["git", "remote", "get-url", "origin"],
        capture_output=True, text=True, cwd=str(REPO_DIR)
    )
    url = result.stdout.strip()
    # 既にトークンが埋め込まれている場合はスキップ
    if "@" in url and "github.com" in url:
        return
    # https://github.com/... → https://TOKEN@github.com/...
    if url.startswith("https://github.com/"):
        new_url = url.replace("https://github.com/", f"https://{token}@github.com/")
        subprocess.run(
            ["git", "remote", "set-url", "origin", new_url],
            cwd=str(REPO_DIR)
        )


def git_push(message=None):
    """生成したHTMLをgit commit & pushする。"""
    if message is None:
        message = f"Update tasks {date.today().isoformat()}"

    os.chdir(REPO_DIR)
    setup_git_remote()

    # 変更があるか確認
    result = subprocess.run(
        ["git", "status", "--porcelain"],
        capture_output=True, text=True
    )
    if not result.stdout.strip():
        print("No changes to commit.")
        return False

    subprocess.run(["git", "add", "index.html"], check=True)
    subprocess.run(["git", "commit", "-m", message], check=True)
    subprocess.run(["git", "pull", "--rebase"], capture_output=True)
    subprocess.run(["git", "push"], check=True)
    print("Pushed to GitHub Pages.")
    return True


def next_task_id(cache):
    """次のタスクIDを生成する。"""
    existing_ids = [t["id"] for t in cache["tasks"]]
    max_num = 0
    for tid in existing_ids:
        if tid.startswith("t"):
            try:
                num = int(tid[1:])
                max_num = max(max_num, num)
            except ValueError:
                pass
    return f"t{max_num + 1}"


def add_task_interactive(cache):
    """対話的にタスクを追加する。"""
    task_id = next_task_id(cache)
    print(f"\n--- 新しいタスクを追加 (ID: {task_id}) ---")

    title = input("タイトル: ").strip()
    if not title:
        print("キャンセルしました。")
        return

    priority = input("優先度 (high/mid/low) [mid]: ").strip() or "mid"
    from_name = input("差出人: ").strip()
    deadline = input("期限: ").strip()
    urgent = input("緊急? (y/n) [n]: ").strip().lower() == "y"
    note = input("備考: ").strip()
    summary = input("要約 (複数行は\\nで区切り): ").strip().replace("\\n", "\n")

    task = {
        "id": task_id,
        "priority": priority,
        "title": title,
        "from": from_name,
        "fromEmail": "",
        "mailDate": date.today().isoformat(),
        "mailSubject": title,
        "deadline": deadline,
        "urgent": urgent,
        "note": note,
        "summary": summary,
        "completed": False,
        "created_at": datetime.now().isoformat(),
    }

    cache["tasks"].append(task)
    save_cache(cache)
    print(f"タスク '{title}' を追加しました。")


def main():
    dry_run = "--dry-run" in sys.argv
    add_task = "--add-task" in sys.argv

    cache = load_cache()

    if add_task:
        add_task_interactive(cache)

    # HTML生成
    generate_html(cache)

    # Git push
    if not dry_run:
        git_push()
    else:
        print("(dry-run: git push skipped)")


if __name__ == "__main__":
    main()
