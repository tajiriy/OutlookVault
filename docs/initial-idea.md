# OutlookArchiver 初期アイデア整理

## 背景・動機

- 会社の Outlook メールが **90日間で自動削除** されるポリシーに変更される
- 過去のメールを手軽に残し、あとから参照できる仕組みが必要
- できれば通常のメールクライアントよりも **見やすく・使いやすく** したい

## コンセプト

Outlook のメールを SQLite にアーカイブし、高速な検索・快適な閲覧を提供するデスクトップアプリケーション。

## 技術スタック

| 項目 | 選定 |
|------|------|
| 言語 | VB.NET |
| フレームワーク | .NET Framework 4.6.2 / Windows Forms |
| メール取得 | Outlook COM API (デスクトップ版 Outlook) |
| データベース | SQLite (System.Data.SQLite) |
| 全文検索 | SQLite FTS5 |

## 対象フォルダ

- 受信トレイ
- 送信済みアイテム
- その他任意のフォルダ（ユーザーが選択可能）

## 機能一覧

### 1. メール取り込み

| 機能 | 説明 |
|------|------|
| 手動取り込み | ボタン押下で選択フォルダのメールを取り込む |
| 自動取り込み | アプリ起動中に定期ポーリングで新着メールを自動取り込み |
| 重複排除 | Outlook の EntryID / MessageID を利用して同一メールの二重取り込みを防止 |
| 進捗表示 | 取り込み中の件数・進捗をステータスバーに表示 |

### 2. データ保存

#### メール本体 (SQLite)

- 件名、本文（プレーンテキスト＋HTML）、差出人、宛先、CC、BCC
- 受信日時、送信日時
- Outlook EntryID / MessageID（重複排除用）
- フォルダ名（取り込み元）
- 正規化済み件名（`Re:` `FW:` 等を除去。会話グルーピング用）

#### 添付ファイル (ファイルシステム + SQLite メタデータ)

- ファイルの実体は **ファイルシステムに保存**（DB には格納しない）
- 保存先ディレクトリ構造: `attachments/<YYYYMMDD_HHmmss>_<MessageID短縮>/filename.ext`
  - 受信日時ベースのフォルダで同名ファイルの衝突を回避
- SQLite には添付ファイルのメタデータ（ファイル名、パス、サイズ、MIME タイプ）を保持
- メールとの紐付けは外部キーで管理

### 3. メール閲覧

#### 一覧表示（上半分）

- 列: 受信日時、件名、差出人、添付アイコン
- デフォルトソート: 受信日時の降順
- フォルダ別フィルタリング

#### 本文プレビュー（下半分）

- 一覧で選択したメールの本文を表示
- HTML メールの場合は WebBrowser コントロールでレンダリング
- プレーンテキストの場合はテキスト表示

#### 本文表示の切り替え

- プレーンテキスト表示 ↔ HTML 表示のトグルボタンを用意
- HTML メールはまず標準 WebBrowser コントロール（IE エンジン）で表示し、表示崩れが問題になれば CefSharp（Chromium エンジン）への切り替えを検討する

#### 会話ビュー

- **スレッド判定の優先順位**（信頼性の高い順）:
  1. メールヘッダーの `In-Reply-To` / `References` フィールドを使ったスレッドチェーン
  2. 取得できない場合は正規化済み件名（`Re:` `Fw:` `FW:` 等を除去）でフォールバック
- ヘッダーベースの判定により、件名が同じでも無関係なシステム通知メールが誤ってグルーピングされることを防止
- 同一スレッドのメールを **時系列順（昇順）** に本文を連続表示
- 引用部分を除去した「本文のみ」を時系列で並べることで、会話の流れを読みやすくする

### 4. 返信引用の自動除去

通常のメールでは返信の度に過去のやり取りが累積して読みにくくなる。会話ビューでは以下のルールで引用部分を自動除去する。

#### 除去対象

- `-----Original Message-----` 等の区切り線以降を**全削除**
  - 対象パターン例: `-----Original Message-----` / `________________________________` / `From: ... Sent: ... To: ...` 形式のブロック
- **メール末尾から遡って連続する `> ` 行のブロック**を削除
  - 末尾からの連続判定のため、本文中間部の手入力引用を誤って削除するリスクが低い

#### 除去しないもの

- 本文の途中に現れる単発・少数の `> ` 行（手入力の引用と判断）

#### 補足

- 除去前のオリジナル本文は DB に保持し、除去は表示時に適用する（非破壊）
- 「引用を表示」トグルで元の全文を確認可能

### 5. 検索

- **SQLite FTS5** による全文検索
- 検索対象（オプションで選択可能）:
  - 件名
  - 本文
  - 差出人 / 宛先
  - 添付ファイル名
- 検索結果は一覧表示と同じフォーマットで表示
- ヒット箇所のハイライト表示

### 6. 添付ファイル管理

- メールプレビュー内に添付ファイル一覧を表示
- クリックでデフォルトアプリケーションで開く
- 画像・PDF はアプリ内プレビュー対応
  - 画像: PictureBox コントロール
  - PDF: WebBrowser コントロール or 外部ライブラリ

### 7. 削除

- 不要なメールの個別削除・一括削除
- 削除時は添付ファイル（実ファイル）も合わせて削除
- 削除確認ダイアログ
- **削除済みメールの再取り込み防止**: 削除時に `message_id` を `deleted_message_ids` テーブルに記録し、以降の取り込み時にスキップする（トゥームストーン方式）

## DB スキーマ（案）

```sql
-- メール本体
CREATE TABLE emails (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    message_id      TEXT UNIQUE,        -- RFC 2822 Message-ID (重複排除・スレッド用)
    in_reply_to     TEXT,               -- In-Reply-To ヘッダー (スレッド判定用)
    references      TEXT,               -- References ヘッダー (スレッドチェーン)
    thread_id       TEXT,               -- スレッドID (判定後に付与)
    entry_id        TEXT,               -- Outlook EntryID
    subject         TEXT,               -- 件名
    normalized_subject TEXT,            -- 正規化済み件名 (Re:/FW: 除去。フォールバック用)
    sender_name     TEXT,               -- 差出人名
    sender_email    TEXT,               -- 差出人メールアドレス
    to_recipients   TEXT,               -- 宛先 (JSON 配列)
    cc_recipients   TEXT,               -- CC (JSON 配列)
    bcc_recipients  TEXT,               -- BCC (JSON 配列)
    body_text       TEXT,               -- プレーンテキスト本文
    body_html       TEXT,               -- HTML 本文
    received_at     TEXT NOT NULL,      -- 受信日時 (ISO 8601)
    sent_at         TEXT,               -- 送信日時 (ISO 8601)
    folder_name     TEXT,               -- 取り込み元フォルダ名
    has_attachments INTEGER DEFAULT 0,  -- 添付ファイル有無フラグ
    created_at      TEXT DEFAULT (datetime('now', 'localtime')),
    updated_at      TEXT DEFAULT (datetime('now', 'localtime'))
);

-- 全文検索用仮想テーブル
CREATE VIRTUAL TABLE emails_fts USING fts5(
    subject,
    body_text,
    sender_name,
    sender_email,
    content='emails',
    content_rowid='id'
);

-- 添付ファイル
CREATE TABLE attachments (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    email_id        INTEGER NOT NULL REFERENCES emails(id) ON DELETE CASCADE,
    file_name       TEXT NOT NULL,      -- 元のファイル名
    file_path       TEXT NOT NULL,      -- 保存先パス (相対パス)
    file_size       INTEGER,            -- ファイルサイズ (bytes)
    mime_type       TEXT,               -- MIME タイプ
    created_at      TEXT DEFAULT (datetime('now', 'localtime'))
);

-- 添付ファイル名も全文検索に含める
CREATE VIRTUAL TABLE attachments_fts USING fts5(
    file_name,
    content='attachments',
    content_rowid='id'
);

-- 削除済みメールの再取り込み防止（トゥームストーン）
CREATE TABLE deleted_message_ids (
    message_id  TEXT PRIMARY KEY,       -- 削除済みメールの Message-ID
    deleted_at  TEXT DEFAULT (datetime('now', 'localtime'))
);

-- インデックス
CREATE INDEX idx_emails_received_at ON emails(received_at DESC);
CREATE INDEX idx_emails_normalized_subject ON emails(normalized_subject);
CREATE INDEX idx_emails_message_id ON emails(message_id);
CREATE INDEX idx_emails_thread_id ON emails(thread_id);
CREATE INDEX idx_emails_in_reply_to ON emails(in_reply_to);
CREATE INDEX idx_attachments_email_id ON attachments(email_id);
```

## UI レイアウト（案）

```
+---------------------------------------------------------------+
| メニューバー [ファイル] [取り込み] [検索] [設定] [ヘルプ]          |
+---------------------------------------------------------------+
| ツールバー [今すぐ取り込み] [自動取り込み: ●停止/▶開始] [検索ボックス        ] [検索] [設定] |
+---------------------------------------------------------------+
| フォルダ  |  添付 | 受信日時     | 件名          | 差出人       |
| ツリー    |  ---  | ----------- | ------------- | ----------- |
|           |  📎  | 2025/03/18  | 会議の件      | 田中太郎     |
| > 受信    |      | 2025/03/17  | 報告書        | 鈴木花子     |
| > 送信済  |  📎  | 2025/03/16  | Re: 見積もり   | 佐藤次郎     |
| > ...     |      | ...         | ...           | ...         |
+-----------+------+-------------+---------------+-------------+
| 本文プレビュー / 会話ビュー                                      |
|                                                               |
| [通常表示] [会話ビュー] タブ                                     |
|                                                               |
| From: 田中太郎                                                  |
| Date: 2025/03/18 10:30                                        |
| Subject: 会議の件                                               |
|                                                                |
| お疲れ様です。                                                   |
| 明日の会議の件ですが...                                           |
|                                                                |
| 添付: [report.pdf] [data.xlsx]                                 |
+---------------------------------------------------------------+
| ステータスバー: メール総数: 1,234 | 最終取り込み: 2025/03/18 10:00 |
+---------------------------------------------------------------+
```

## 設定項目

| 設定項目 | 説明 | デフォルト値 |
|---------|------|------------|
| DB ファイルパス | SQLite データベースの保存先 | `./data/archive.db` |
| 添付ファイル保存先 | 添付ファイルの保存ディレクトリ | `./data/attachments/` |
| 自動取り込み有効/無効 | 起動時に自動取り込みを開始するか | 有効 |
| 自動取り込み間隔 | ポーリング間隔（分）。起動中にも変更可能 | 10分 |
| 1回の最大取り込み件数 | ポーリングまたは手動取り込み1回あたりの上限件数 | 100件 |
| 対象フォルダ | アーカイブ対象の Outlook フォルダ一覧 | 受信トレイ |
| 会話ビューの並び順 | 同一スレッド内のメールを時系列昇順/降順どちらで表示するか | 昇順（古い順） |
| 本文表示モード | デフォルトをテキスト/HTML どちらにするか | HTML |

## 保存先について

- 基本はローカル PC
- ネットワークドライブも対応（パス指定で切り替え可能）
- SQLite はネットワークドライブでの同時アクセスに制限があるため、**単一ユーザーアクセスを前提** とする

## 将来の展開

- 個人利用で検証後、社内展開を検討
- 社内展開時は設定ファイルの共通化やインストーラーの整備が必要になる可能性あり

## 技術的な留意点

- **Outlook COM API**: アプリ実行時に Outlook が起動している必要がある（またはバックグラウンドで起動）
- **SQLite FTS5**: .NET Framework から利用する場合、System.Data.SQLite の FTS5 対応を確認する必要あり
- **WebBrowser コントロール**: HTML メール表示用。IE ベースのため、表示崩れの可能性あり。まず標準コントロールで実装し、問題があれば CefSharp（Chromium ベースの .NET 向けブラウザコンポーネント。バイナリが +100MB 程度増加するトレードオフあり）への切り替えを検討
- **In-Reply-To / References の取得**: Outlook COM API では `MailItem.PropertyAccessor` を使って MAPI プロパティから取得する。取得できない場合は正規化済み件名でフォールバック
- **大量メール**: 数万件規模のメールを扱う場合、仮想リスト（Virtual ListView）で一覧のパフォーマンスを確保する
