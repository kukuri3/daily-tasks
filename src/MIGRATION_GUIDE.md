# Windows移植ガイド

## 概要

macOS + Outlook Web (Claude in Chrome) で動作しているシステムを、
Windows + Outlook デスクトップ (COM) に移植する。

## 移植による変更点

| 項目 | macOS（現在） | Windows（移植先） |
|---|---|---|
| メール取得 | Claude in Chrome MCP | **COM (`win32com`)** |
| タスク抽出 | Claudeセッション内で手動 | **`extract_tasks_win.py` で自動** |
| HTML生成 | `generate_html.py` | **そのまま使用** |
| テンプレート | `templates/index.html.j2` | **そのまま使用** |
| キャッシュ | `tasks_cache.json` | **そのまま使用** |
| 定期実行 | なし（手動） | **Windowsタスクスケジューラ** |
| GUI | rumps メニューバー | **pystray タスクトレイ** |
| デプロイ | `git push` | **そのまま使用** |

## 前提条件

### ソフトウェア

- Windows 10/11
- Python 3.10+ （https://python.org からインストール）
- git （https://git-scm.com からインストール）
- Microsoft Outlook デスクトップ版（Microsoft 365契約に含まれる）

### Pythonパッケージ

```bash
pip install -r requirements.txt
```

`requirements.txt`:
```
jinja2
pywin32        # Windows COM (macOSでは不要)
anthropic      # Claude API (タスク抽出用、オプション)
pystray        # タスクトレイアプリ (macOSのrumps代替)
Pillow         # pystrayの依存
```

## 移植手順

### Step 1: リポジトリのクローン

```bash
git clone https://github.com/kukuri3/daily-task-system.git
cd daily-task-system
pip install -r requirements.txt
```

### Step 2: GitHub Pages リポジトリの準備

```bash
mkdir repo
cd repo
git clone https://github.com/kukuri3/daily-tasks.git .
cd ..
```

gitの認証設定:
```bash
# GitHub PAT を使う場合
cd repo
git remote set-url origin https://YOUR_PAT@github.com/kukuri3/daily-tasks.git
```

### Step 3: Outlook COM の動作確認

Outlookデスクトップアプリを起動した状態で:

```python
python -c "
import win32com.client
outlook = win32com.client.Dispatch('Outlook.Application')
ns = outlook.GetNamespace('MAPI')
inbox = ns.GetDefaultFolder(6)
print(f'受信トレイ: {inbox.Items.Count}件')
for i in range(min(5, inbox.Items.Count)):
    msg = inbox.Items[i+1]
    print(f'  {msg.ReceivedTime} | {msg.SenderName} | {msg.Subject}')
"
```

### Step 4: extract_tasks_win.py の設定

`extract_tasks_win.py` を開き、以下を確認:
- `SCAN_WEEKS`: スキャン対象期間（デフォルト3週間）
- `EXCLUDE_SENDERS`: 除外する差出人リスト（メルマガ等）
- `SELF_EMAIL`: 自分のメールアドレス（自分宛メモ判定用）

### Step 5: 初回実行

```bash
# Outlookからメール取得→タスク抽出→HTML生成→push
python extract_tasks_win.py

# HTMLのみ再生成（メール取得なし）
python generate_html.py
```

### Step 6: Windowsタスクスケジューラで定期実行

1. 「タスクスケジューラ」を開く
2. 「基本タスクの作成」
3. 名前: `Daily Task Extractor`
4. トリガー: 毎週平日、8:00と13:00の2つ
5. 操作: プログラムの開始
   - プログラム: `python`
   - 引数: `C:\path\to\daily-task-system\extract_tasks_win.py`
   - 開始: `C:\path\to\daily-task-system\`
6. 条件: 「コンピューターをAC電源で使用している場合のみ」をオフに

### Step 7: タスクトレイアプリ（オプション）

macOSのメニューバーアプリ（`menubar_app.py`）のWindows版として、
`tray_app_win.py` を使用。`pystray`ライブラリでタスクトレイに常駐。

## COM メール取得の詳細

### フォルダ番号

| 番号 | フォルダ |
|---|---|
| 6 | 受信トレイ (Inbox) |
| 5 | 送信済みアイテム (Sent Items) |
| 3 | 削除済みアイテム |
| 16 | 下書き |

### メッセージオブジェクトの主要プロパティ

| プロパティ | 型 | 内容 |
|---|---|---|
| `Subject` | str | 件名 |
| `SenderName` | str | 差出人名 |
| `SenderEmailAddress` | str | 差出人メールアドレス |
| `ReceivedTime` | datetime | 受信日時 |
| `Body` | str | 本文（プレーンテキスト） |
| `HTMLBody` | str | 本文（HTML） |
| `ConversationID` | str | スレッドID |
| `Attachments` | collection | 添付ファイル |
| `UnRead` | bool | 未読フラグ |

### スレッド完結判定のロジック

```python
# 受信メールのConversationIDで送信トレイを検索
sent_folder = ns.GetDefaultFolder(5)  # Sent Items
for sent_msg in sent_folder.Items:
    if sent_msg.ConversationID == received_msg.ConversationID:
        # 自分が返信している → waiting or resolved
        break
```

## タスク抽出のアプローチ

### 方式A: ルールベース（API不要）

メール件名・本文のキーワードからタスクを抽出:
- 「依頼」「お願い」「確認」「〆切」「期限」→ タスク候補
- 「お知らせ」「ご案内」のみ → 低優先度
- 自分宛メモ → 必ずタスク追加

### 方式B: Claude API（高精度）

メール本文をClaude APIに送信してJSON形式でタスクを抽出。
`extract_tasks_win.py` のデフォルト実装はこの方式。

```python
from anthropic import Anthropic
client = Anthropic()  # ANTHROPIC_API_KEY 環境変数
response = client.messages.create(
    model="claude-sonnet-4-20250514",
    messages=[{"role": "user", "content": prompt}],
)
```

### 方式C: ローカルLLM（コスト0）

評価の結果、8GB RAMでは精度不足と判断（eval/comparison_report.md参照）。
48GB以上のRAMがあれば再検討の余地あり。

## トラブルシューティング

### COM接続エラー

```
pywintypes.com_error: (-2147221005, 'Invalid class string', ...)
```
→ Outlookデスクトップアプリが起動していない。起動してから再実行。

### git pushエラー

```
remote: Permission denied
```
→ `repo/`内のgit remoteにPATが設定されていない。Step 2を確認。

### Outlookのセキュリティ警告

COMアクセス時に「プログラムがメールにアクセスしようとしています」と表示される場合:
- Outlook → ファイル → オプション → セキュリティセンター → プログラムによるアクセス
- 「不審な動作について警告しない」を選択（組織ポリシーで制限されている場合あり）
