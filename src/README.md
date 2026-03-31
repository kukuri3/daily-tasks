# Daily Task Extractor

Outlookメールから未完了タスクを自動抽出し、優先度付きのインタラクティブなHTMLタスクリストとしてGitHub Pagesに公開するシステム。

**公開URL:** https://kukuri3.github.io/daily-tasks/

## 機能

- Outlookメールからアクション必要なタスクを自動抽出（メルマガ・広告は除外）
- 自分宛メモ（美馬→美馬）はタスクとして自動追加
- 送信トレイ突合によるスレッド完結判定（open / waiting / resolved）
- 優先度別＋ステータス別のカテゴリ表示（📨要対応 / ⏳返信待ち / ✅完結）
- チェックボックス（localStorage + GitHub API同期）
- チェック後5秒間の「取り消し」ボタン（誤チェック救済）
- メール要約トグル、メール本文モーダル
- スマートフォン対応レイアウト
- GitHub Pagesへの自動デプロイ

## アーキテクチャ

### 現在（macOS + Outlook Web）

```
[Outlook Web] → [Claude in Chrome MCP] → [Claude: タスク抽出]
                                                ↓
                                          [tasks_cache.json]
                                                ↓
                                    [generate_html.py + Jinja2]
                                                ↓
                                          [git push → GitHub Pages]
```

### 移植先（Windows + Outlook デスクトップ）

```
[Outlook デスクトップ] → [extract_tasks.py (COM)] → [Claude API: タスク抽出]
                                                          ↓
                                                    [tasks_cache.json]
                                                          ↓
                                              [generate_html.py + Jinja2]
                                                          ↓
                                                    [git push → GitHub Pages]
```

## ファイル構成

```
daily_task/
├── README.md                     # 本ファイル
├── MIGRATION_GUIDE.md            # Windows移植ガイド
├── generate_html.py              # HTML生成 + git push（共通）
├── menubar_app.py                # macOSメニューバーアプリ（macOS専用）
├── extract_tasks_win.py          # Outlook COM メール取得（Windows用テンプレート）
├── tasks_cache.json              # タスクキャッシュ（差分管理）
├── requirements.txt              # Python依存パッケージ
├── templates/
│   └── index.html.j2             # HTMLテンプレート（Jinja2）
├── repo/                         # GitHub Pages リポジトリ（kukuri3/daily-tasks）
│   ├── index.html                # 生成されたHTML
│   └── done.json                 # チェック状態同期用
├── docs/
│   ├── azure_ad_request.md       # Graph API申請テンプレート（将来用）
│   └── system_design.md          # 詳細設計書
├── eval/
│   └── comparison_report.md      # LLMモデル比較レポート
└── logs/                         # 実行ログ
```

## セットアップ

### 前提条件

- Python 3.10+
- git
- GitHub アカウント（GitHub Pages用）
- Outlook（デスクトップ版 or Web版）

### インストール

```bash
git clone https://github.com/kukuri3/daily-task-system.git
cd daily-task-system
pip install -r requirements.txt
```

### GitHub Pages リポジトリのクローン

```bash
cd repo
git clone https://github.com/kukuri3/daily-tasks.git .
```

### 実行

```bash
# HTMLビルド＋デプロイ
python generate_html.py

# ドライラン（pushなし）
python generate_html.py --dry-run

# Windows: Outlookからメール取得→タスク抽出→ビルド→デプロイ
python extract_tasks_win.py
```

## タスクキャッシュ仕様

`tasks_cache.json`のフォーマット:

```json
{
  "last_scan": "2026-03-31T09:45:00",
  "tasks": [
    {
      "id": "t1",
      "priority": "high",
      "title": "タスク名",
      "from": "差出人",
      "fromEmail": "email@example.com",
      "mailDate": "2026-03-19",
      "mailSubject": "メール件名",
      "deadline": "期限",
      "urgent": true,
      "note": "備考",
      "summary": "メール要約",
      "thread_status": "open",
      "thread_summary": "スレッド経緯",
      "related_threads": ["t2"],
      "completed": false
    }
  ]
}
```

### thread_status の値

| 値 | 意味 | 表示カテゴリ |
|---|---|---|
| `open` | 自分のアクションが必要 | 📨 要対応 |
| `waiting` | 相手の返答待ち | ⏳ 返信待ち |
| `resolved` | メール上は完結 | ✅ スレッド完結 |

## ライセンス

Private - 静岡理工科大学 美馬研究室
