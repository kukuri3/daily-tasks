#!/usr/bin/env python3
"""
Windows Outlook COM メール取得 → タスク抽出 → HTML生成 → GitHub Pages デプロイ

前提:
  - Windows 10/11
  - Microsoft Outlook デスクトップ版が起動中
  - pip install pywin32 jinja2

使い方:
  python extract_tasks_win.py              # 差分スキャン（last_scan以降）
  python extract_tasks_win.py --rescan 3   # 過去3週間フルスキャン
  python extract_tasks_win.py --dry-run    # メール取得のみ（cache書込・pushなし）
"""

import json
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path

try:
    import win32com.client
    import pythoncom
except ImportError:
    print("ERROR: pywin32 が必要です。 pip install pywin32")
    sys.exit(1)

BASE_DIR = Path(__file__).parent
CACHE_FILE = BASE_DIR / "tasks_cache.json"
GENERATE_SCRIPT = BASE_DIR / "generate_html.py"

# --- 設定 ---
SCAN_WEEKS = 3  # デフォルトのスキャン期間（週）
SELF_EMAIL = "mima.kazuhiro@sist.ac.jp"

# メルマガ・広告として除外する差出人パターン
EXCLUDE_SENDERS = [
    r"AliExpress",
    r"Amazon\s*Business",
    r"Amazon\.co\.jp",
    r"amazon\.co\.jp",
    r"no-?reply",
    r"noreply",
    r"ieej_office@iee\.or\.jp",  # 電気学会メルマガ
    r"ミスミ",
    r"MISUMI",
    r"meviy",
    r"オートデスク",
    r"Autodesk",
    r"NVIDIA",
    r"Zoom\s*ジャパン",
    r"SharePoint\s*Online",
    r"RSJ-NEWS",
    r"jsme_2026@mta\.co\.jp",  # 日本機械学会
    r"オリジナルマインド",
    r"アルテックスグループ",
    r"北陽電機",  # 広告メール
    r"WDB\s*人材情報",
    r"ネイチャーショップ",
    r"アットホーム",
    r"静岡県産業振興財団",
    r"南／エージェンシーアシスト",
    r"FosterLink",  # 360度診断の自動送信
]

# 情報共有のみ（タスク化するが低優先度）
INFO_ONLY_PATTERNS = [
    r"お知らせ",
    r"ご案内",
    r"メンテナンス",
    r"工事",
]


def load_cache():
    if not CACHE_FILE.exists():
        return {"last_scan": None, "tasks": []}
    with open(CACHE_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def save_cache(cache):
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


def is_excluded(sender_name, sender_email):
    """メルマガ・広告の差出人かどうか"""
    for pattern in EXCLUDE_SENDERS:
        if re.search(pattern, sender_name or "", re.IGNORECASE):
            return True
        if re.search(pattern, sender_email or "", re.IGNORECASE):
            return True
    return False


def is_self_memo(sender_email, recipients):
    """自分宛メモ（自分→自分）かどうか"""
    if not sender_email:
        return False
    sender = sender_email.lower()
    if SELF_EMAIL.lower() not in sender:
        return False
    # 宛先にも自分が含まれている
    for r in (recipients or "").lower().split(";"):
        if SELF_EMAIL.lower() in r.strip():
            return True
    return False


def get_priority(subject, body):
    """件名・本文から優先度を推定"""
    text = (subject or "") + " " + (body or "")[:500]
    if re.search(r"【重要】|【緊急】|【至急】|urgent|deadline|〆切|期限.*本日|今日中", text, re.IGNORECASE):
        return "high"
    if re.search(r"【依頼】|【確認】|【お願い】|回答.*まで|提出.*まで", text, re.IGNORECASE):
        return "mid"
    return "low"


def fetch_emails(since_date=None):
    """Outlook COM でメール取得"""
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")

        # 受信トレイ
        inbox = ns.GetDefaultFolder(6)
        inbox_msgs = []

        if since_date:
            # DASL フィルタで日付絞り込み
            filter_str = f"[ReceivedTime] >= '{since_date.strftime('%m/%d/%Y %H:%M %p')}'"
            items = inbox.Items.Restrict(filter_str)
        else:
            items = inbox.Items

        items.Sort("[ReceivedTime]", True)  # 新しい順

        for item in items:
            try:
                if not hasattr(item, "Subject"):
                    continue
                inbox_msgs.append({
                    "subject": item.Subject or "",
                    "sender_name": item.SenderName or "",
                    "sender_email": getattr(item, "SenderEmailAddress", "") or "",
                    "received_time": item.ReceivedTime.strftime("%Y-%m-%dT%H:%M:%S") if item.ReceivedTime else "",
                    "body": (item.Body or "")[:2000],  # 本文先頭2000文字
                    "conversation_id": getattr(item, "ConversationID", "") or "",
                    "unread": getattr(item, "UnRead", False),
                    "folder": "inbox",
                })
            except Exception as e:
                print(f"  WARN: メール読み取りスキップ: {e}")
                continue

        # 送信トレイ
        sent = ns.GetDefaultFolder(5)
        sent_msgs = []

        if since_date:
            filter_str = f"[SentOn] >= '{since_date.strftime('%m/%d/%Y %H:%M %p')}'"
            sent_items = sent.Items.Restrict(filter_str)
        else:
            sent_items = sent.Items

        sent_items.Sort("[SentOn]", True)

        for item in sent_items:
            try:
                if not hasattr(item, "Subject"):
                    continue
                sent_msgs.append({
                    "subject": item.Subject or "",
                    "to": getattr(item, "To", "") or "",
                    "sent_time": item.SentOn.strftime("%Y-%m-%dT%H:%M:%S") if item.SentOn else "",
                    "body": (item.Body or "")[:1000],
                    "conversation_id": getattr(item, "ConversationID", "") or "",
                    "folder": "sent",
                })
            except Exception:
                continue

        print(f"取得: 受信{len(inbox_msgs)}件, 送信{len(sent_msgs)}件")
        return inbox_msgs, sent_msgs

    finally:
        pythoncom.CoUninitialize()


def determine_thread_status(msg, sent_msgs):
    """送信トレイとの突合でスレッド完結判定"""
    conv_id = msg.get("conversation_id", "")
    if not conv_id:
        return "open"

    # 同じスレッドで自分が返信しているか
    my_replies = [s for s in sent_msgs if s.get("conversation_id") == conv_id]
    if not my_replies:
        return "open"

    # 自分が返信済み → 相手の反応があるか確認
    # （簡易判定: 相手からの最新受信が自分の最新送信より後ならresolved）
    my_latest = max(s["sent_time"] for s in my_replies)
    msg_time = msg.get("received_time", "")

    if msg_time > my_latest:
        # 相手が自分の返信の後にさらに返信 → 内容による
        body = msg.get("body", "").lower()
        if any(w in body for w in ["承知", "了解", "ありがとう", "確認しました", "お礼"]):
            return "resolved"
        return "open"  # 追加の要求かもしれない

    return "waiting"  # 自分が返信済みで相手の反応なし


def extract_tasks(inbox_msgs, sent_msgs, existing_tasks=None):
    """メールからタスクを抽出"""
    existing_ids = set()
    if existing_tasks:
        existing_ids = {t.get("conversation_id") for t in existing_tasks if t.get("conversation_id")}

    tasks = []
    next_id = 1
    if existing_tasks:
        max_num = max(
            (int(t["id"][1:]) for t in existing_tasks if t["id"].startswith("t") and t["id"][1:].isdigit()),
            default=0
        )
        next_id = max_num + 1

    for msg in inbox_msgs:
        # 除外判定
        if is_excluded(msg["sender_name"], msg["sender_email"]):
            continue

        # 重複チェック（同一スレッド）
        if msg.get("conversation_id") in existing_ids:
            continue

        # 優先度判定
        priority = get_priority(msg["subject"], msg["body"])

        # 自分宛メモ判定
        sender_email = msg.get("sender_email", "")
        self_memo = SELF_EMAIL.lower() in sender_email.lower()

        # スレッド完結判定
        thread_status = determine_thread_status(msg, sent_msgs)

        # 期限推定（メール本文からキーワード抽出）
        deadline = "随時"
        body_text = msg.get("body", "")
        deadline_match = re.search(
            r"(\d{1,2}/\d{1,2}[（(][月火水木金土日][）)]?.*?まで|"
            r"\d{1,2}月\d{1,2}日.*?まで|"
            r"本日中|今日中|至急|早急|年度末)",
            body_text
        )
        if deadline_match:
            deadline = deadline_match.group(0)[:30]

        task = {
            "id": f"t{next_id}",
            "priority": "mid" if self_memo else priority,
            "title": msg["subject"][:60],
            "from": f"{msg['sender_name']}（自分メモ）" if self_memo else msg["sender_name"],
            "fromEmail": msg.get("sender_email", ""),
            "mailDate": msg.get("received_time", "")[:10],
            "mailSubject": msg["subject"],
            "deadline": deadline,
            "urgent": priority == "high",
            "note": body_text[:200].replace("\n", " ").strip(),
            "summary": body_text[:500].strip(),
            "thread_status": thread_status,
            "thread_summary": f"{msg.get('received_time', '')[:10]} 受信。",
            "related_threads": [],
            "conversation_id": msg.get("conversation_id", ""),
            "completed": thread_status == "resolved",
        }

        tasks.append(task)
        next_id += 1
        existing_ids.add(msg.get("conversation_id"))

    return tasks


def main():
    args = sys.argv[1:]
    dry_run = "--dry-run" in args
    rescan_weeks = None

    for i, arg in enumerate(args):
        if arg == "--rescan" and i + 1 < len(args):
            rescan_weeks = int(args[i + 1])

    cache = load_cache()

    # スキャン期間の決定
    if rescan_weeks:
        since_date = datetime.now() - timedelta(weeks=rescan_weeks)
        cache = {"last_scan": None, "tasks": []}
        print(f"再スキャン: 過去{rescan_weeks}週間（{since_date.strftime('%Y-%m-%d')}以降）")
    elif cache.get("last_scan"):
        since_date = datetime.fromisoformat(cache["last_scan"])
        print(f"差分スキャン: {since_date.strftime('%Y-%m-%d %H:%M')} 以降")
    else:
        since_date = datetime.now() - timedelta(weeks=SCAN_WEEKS)
        print(f"初回スキャン: 過去{SCAN_WEEKS}週間")

    # メール取得
    inbox_msgs, sent_msgs = fetch_emails(since_date)

    # タスク抽出
    existing = cache.get("tasks", []) if not rescan_weeks else []
    new_tasks = extract_tasks(inbox_msgs, sent_msgs, existing)
    print(f"新規タスク: {len(new_tasks)}件")

    # キャッシュ更新
    if rescan_weeks:
        cache["tasks"] = new_tasks
    else:
        cache["tasks"].extend(new_tasks)

    cache["last_scan"] = datetime.now().isoformat()

    if not dry_run:
        save_cache(cache)
        print(f"キャッシュ更新: {len(cache['tasks'])}件")

        # HTML生成 + push
        import subprocess
        result = subprocess.run(
            [sys.executable, str(GENERATE_SCRIPT)],
            capture_output=True, text=True, cwd=str(BASE_DIR)
        )
        print(result.stdout)
        if result.stderr:
            print(result.stderr)
    else:
        print("(dry-run: キャッシュ書込・pushスキップ)")
        for t in new_tasks[:10]:
            print(f"  [{t['priority']}] {t['title']} ({t['thread_status']})")


if __name__ == "__main__":
    main()
