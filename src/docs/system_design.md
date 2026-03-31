# Daily Task Extractor 詳細設計書

**バージョン:** 2.0
**作成日:** 2026-03-31
**作成者:** Claude Code + 美馬 一博

---

## 1. システム概要

### 1.1 目的
大学学科長業務のOutlookメールから未完了タスクを自動抽出し、優先度付きのインタラクティブなHTMLタスクリストとしてGitHub Pagesに公開する。

### 1.2 要件
- メールからアクション必要なタスクを自動抽出
- メルマガ・広告・自動通知は除外
- 自分宛メモ（美馬→美馬）はタスクとして自動追加
- 送信トレイと突合してスレッド完結判定
- 差分スキャン（新着のみ）と全件再スキャンの両方に対応
- スマートフォンからも閲覧可能
- チェックボックスの状態をデバイス間で同期

---

## 2. コンポーネント設計

### 2.1 メール取得レイヤー

#### macOS版（現在）: Claude in Chrome MCP
```
Outlook Web → Claude in Chrome → read_page / get_page_text → テキスト
```
- 制約: Claude Codeセッション内でのみ動作（MCP制約）
- 自動化: 不可（手動指示が必要）

#### Windows版（移植先）: Outlook COM
```
Outlook デスクトップ → win32com.client → Python オブジェクト
```
- 制約: Outlookデスクトップアプリが起動中であること
- 自動化: Windowsタスクスケジューラで完全自動化可能

#### 将来版: Microsoft Graph API
```
Graph API → REST → JSON
```
- 制約: Azure ADアプリ登録（IT管理者承認）が必要
- 自動化: OS問わず完全自動化可能

### 2.2 タスク抽出レイヤー

メールのテキストデータからタスクを構造化JSONに変換する。

**入力:** メールテキスト（件名、差出人、本文、受信日時）
**出力:** タスクJSON（id, priority, title, from, deadline, note, summary, thread_status, ...）

抽出方式:
1. **ルールベース**（`extract_tasks_win.py`実装）: キーワードマッチで優先度・期限を推定
2. **Claude API**（オプション）: メール本文を送信してJSON出力
3. **手動**（現在のmacOS方式）: Claudeセッション内で対話的に抽出

### 2.3 キャッシュレイヤー

`tasks_cache.json` にタスク一覧を永続化。

- 差分スキャン: `last_scan`以降の新着メールのみ処理、既存タスクに追加
- 再スキャン: キャッシュをクリアして全件再構築
- `conversation_id`で同一スレッドの重複を排除

### 2.4 HTML生成レイヤー

`generate_html.py` + `templates/index.html.j2` (Jinja2)

- キャッシュからアクティブタスクと完了タスクを分離
- JSON形式でテンプレートに注入
- `repo/index.html`に出力

### 2.5 デプロイレイヤー

`git add / commit / push` で `kukuri3/daily-tasks` リポジトリにpush。
GitHub Pagesが自動的にHTMLを公開。

---

## 3. データ設計

### 3.1 tasks_cache.json

```json
{
  "last_scan": "ISO 8601 datetime",
  "last_scan_sent": "ISO 8601 datetime",
  "tasks": [Task, ...]
}
```

### 3.2 Task オブジェクト

| フィールド | 型 | 必須 | 説明 |
|---|---|---|---|
| id | string | ○ | タスクID（"t1", "t2", ...） |
| priority | string | ○ | "high" / "mid" / "low" |
| title | string | ○ | タスク名（60文字以内） |
| from | string | ○ | 差出人名 |
| fromEmail | string | | 差出人メールアドレス |
| mailDate | string | | 受信日（YYYY-MM-DD） |
| mailSubject | string | | メール件名 |
| deadline | string | ○ | 期限（自由テキスト） |
| urgent | boolean | | 緊急フラグ |
| note | string | ○ | 備考（200文字以内） |
| summary | string | | メール要約（500文字以内） |
| thread_status | string | | "open" / "waiting" / "resolved" |
| thread_summary | string | | スレッド経緯 |
| related_threads | string[] | | 関連タスクID |
| conversation_id | string | | Outlookスレッド識別子 |
| completed | boolean | | 完了フラグ |

### 3.3 スレッドステータス判定ロジック

```
受信メール → 送信トレイに同一ConversationIDの返信があるか？
  ├─ なし → "open"（未対応）
  └─ あり → 相手の返信がさらにあるか？
       ├─ なし → "waiting"（返答待ち）
       └─ あり → 本文に「承知」「了解」等を含むか？
            ├─ はい → "resolved"（完結）
            └─ いいえ → "open"（追加の要求）
```

---

## 4. HTML UI設計

### 4.1 カテゴリ表示

| 表示順 | カテゴリ | CSSクラス | 対象 |
|---|---|---|---|
| 1 | 📨 要対応 | priority-high | thread_status == "open" |
| 2 | ⏳ 返信待ち | priority-mid | thread_status == "waiting" |
| 3 | ✅ スレッド完結 | priority-low | thread_status == "resolved" かつ未チェック |

各カテゴリ内はpriority順（high→mid→low）でソート。

### 4.2 チェックボックス機能

- チェック → 5秒間「取り消し」ボタン表示 → 5秒後にフェードアウト
- リロード時にチェック済みタスクは非表示
- GitHub API（done.json）でデバイス間同期

### 4.3 完了済みセクション（ページ下部）

- ▶トグルで開閉
- ✕ボタンで完全消去（purge）

### 4.4 レスポンシブ対応

`@media(max-width:640px)`でスマートフォンレイアウトに切替:
- ヘッダー: 縦並び
- タスクメタ: 縦並び
- フォントサイズ縮小

---

## 5. 開発経緯

| 日付 | フェーズ | 内容 |
|---|---|---|
| 3/17 | Phase 0 | Coworkで初期実装。ブラウザ操作でOutlook読取 |
| 3/22 | Phase 1 | Claude Codeに移行。ローカルgit、キャッシュ、テンプレート化 |
| 3/22 | - | メニューバーアプリ(rumps)、launchd定期実行 |
| 3/22 | - | ローカルLLM(Ollama)評価→Sonnet運用に決定 |
| 3/23 | Phase 1.5 | 送信トレイ突合によるスレッド完結判定追加 |
| 3/24 | - | カテゴリ表示（要対応/返信待ち/完結）実装 |
| 3/24 | - | undoトースト（誤チェック救済）実装 |
| 3/24 | - | 要約品質改善（Opusで書き直し） |
| 3/25 | Phase 2 | launchd廃止、リファクタリング（コード45%削減） |
| 3/31 | Phase 3 | Windows移植準備。COM版テンプレート、ドキュメント整備 |

---

## 6. 今後の課題

1. **Windows COM版の実装・テスト** — `extract_tasks_win.py`の実環境テスト
2. **Claude API統合** — ルールベースからAPI抽出への移行（精度向上）
3. **Microsoft Graph API** — IT管理者承認後に実装（OS非依存化）
4. **タスクの自動アーカイブ** — 1ヶ月以上経過したresolvedタスクの自動削除
5. **Windows タスクトレイアプリ** — pystrayでmacOSメニューバーの代替
