# OutlookArchiver 実装プラン

## 全体方針

6フェーズで段階的に実装する。各フェーズ完了後にビルド確認を行い、動作可能な状態を維持しながら開発を進める。

---

## Phase 1: 基盤整備 ✅

**目標**: NuGetパッケージ追加・データモデル・DB層・設定クラスの実装

### 追加ファイル

| ファイル | 役割 |
|---------|------|
| `Models/Email.vb` | メールデータのPOCOクラス |
| `Models/Attachment.vb` | 添付ファイルデータのPOCOクラス |
| `Config/AppSettings.vb` | アプリ設定の読み書き（App.config経由） |
| `Data/DatabaseManager.vb` | SQLite接続・スキーマ初期化・マイグレーション |
| `Data/EmailRepository.vb` | emails/attachments テーブルのCRUD・FTS5検索 |

### NuGetパッケージ

| パッケージ | バージョン | 用途 |
|----------|-----------|------|
| `System.Data.SQLite` | 1.0.118.0 | SQLite + FTS5 |
| `Microsoft.Office.Interop.Outlook` | 15.0.4797.1004 | Outlook COM API |

### DB スキーマ

```sql
CREATE TABLE emails (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    message_id      TEXT UNIQUE,
    in_reply_to     TEXT,
    references      TEXT,
    thread_id       TEXT,
    entry_id        TEXT,
    subject         TEXT,
    normalized_subject TEXT,
    sender_name     TEXT,
    sender_email    TEXT,
    to_recipients   TEXT,     -- JSON 配列
    cc_recipients   TEXT,     -- JSON 配列
    bcc_recipients  TEXT,     -- JSON 配列
    body_text       TEXT,
    body_html       TEXT,
    received_at     TEXT NOT NULL,
    sent_at         TEXT,
    folder_name     TEXT,
    has_attachments INTEGER DEFAULT 0,
    created_at      TEXT DEFAULT (datetime('now', 'localtime')),
    updated_at      TEXT DEFAULT (datetime('now', 'localtime'))
);

CREATE VIRTUAL TABLE emails_fts USING fts5(subject, body_text, sender_name, sender_email, content='emails', content_rowid='id');

CREATE TABLE attachments (
    id       INTEGER PRIMARY KEY AUTOINCREMENT,
    email_id INTEGER NOT NULL REFERENCES emails(id) ON DELETE CASCADE,
    file_name TEXT NOT NULL,
    file_path TEXT NOT NULL,
    file_size INTEGER,
    mime_type TEXT,
    created_at TEXT DEFAULT (datetime('now', 'localtime'))
);

CREATE VIRTUAL TABLE attachments_fts USING fts5(file_name, content='attachments', content_rowid='id');

CREATE TABLE deleted_message_ids (
    message_id TEXT PRIMARY KEY,
    deleted_at TEXT DEFAULT (datetime('now', 'localtime'))
);
```

---

## Phase 2: Outlook 取り込み ✅

**目標**: Outlook COM APIでメールを取得しDBに保存

### 追加ファイル

| ファイル | 役割 |
|---------|------|
| `Services/OutlookService.vb` | COM APIでフォルダ一覧・メール取得・MAPIプロパティ取得 |
| `Services/ImportService.vb` | 重複排除・添付ファイルのFS保存・トゥームストーン確認 |
| `Services/ThreadingService.vb` | スレッドID付与ロジック（In-Reply-To/References → 正規化件名フォールバック） |

### 処理フロー

```
ImportService.ImportFolder(folderName)
  → OutlookService.GetFolder(folderName)
  → MailItemごとに:
       1. MessageID取得・重複/削除済みチェック
       2. Email オブジェクト生成
       3. ThreadingService.AssignThreadId(email)
       4. EmailRepository.InsertEmail(email)
       5. 添付ファイルをFS保存 + EmailRepository.InsertAttachment()
```

---

## Phase 3: メイン UI ✅

**目標**: 3ペインレイアウト・フォルダツリー・メール一覧

### UI 構成

```
+----------------------------------------------------------+
| メニューバー [ファイル][取り込み][検索][設定][ヘルプ]       |
+----------------------------------------------------------+
| ツールバー: [今すぐ取り込み] [自動: ●/▶] [検索ボックス] [検索] |
+----------------------------------------------------------+
| フォルダツリー (TreeView) | メール一覧 (ListView / VirtualMode) |
|                           |-------------------------------------|
|                           | 本文プレビュー (TabControl)          |
|                           |  [通常表示] [会話ビュー]             |
+----------------------------------------------------------+
| ステータスバー: 総数 n件 | 最終取り込み: yyyy/MM/dd HH:mm       |
+----------------------------------------------------------+
```

### 変更ファイル

- `Form1.vb`: SplitContainer・TreeView・ListView（VirtualMode）・TabControl・StatusBar の配置

---

## Phase 4: メール閲覧 ✅

**目標**: 選択メールのプレビュー表示・添付ファイル操作

### 機能

- 「通常表示」タブ: HTML（WebBrowser）/ プレーンテキスト 切り替えトグル
- ヘッダー情報表示（From / Date / Subject / To）
- 添付ファイルパネル: クリックで既定アプリ起動、画像は PictureBox プレビュー

### 追加ファイル

| ファイル | 役割 |
|---------|------|
| `Controls/EmailPreviewControl.vb` | メールプレビュー UserControl |

---

## Phase 5: 会話ビュー ✅

**目標**: スレッド表示・返信引用の自動除去

### スレッド判定ロジック

1. `In-Reply-To` / `References` ヘッダーによるチェーン追跡（優先）
2. 正規化済み件名（`Re:` `Fw:` `FW:` 等を除去）によるフォールバック

### 引用除去ルール

- `-----Original Message-----` 等の区切り線以降を全削除
- メール末尾から遡って連続する `> ` 行ブロックを削除
- 除去は表示時のみ適用（DB のオリジナル本文は保持）

### 追加ファイル

| ファイル | 役割 |
|---------|------|
| `Services/QuoteStripperService.vb` | 引用除去ロジック |

---

## Phase 6: 検索・削除・設定

**目標**: 全文検索、削除、設定フォーム

### 機能

- FTS5 全文検索（件名・本文・差出人・添付ファイル名）
- 検索結果のヒット箇所ハイライト
- 個別・一括削除（添付ファイル実体も削除・トゥームストーン記録）
- 設定フォーム（DB/添付先パス、自動取り込み、対象フォルダ選択）

### 追加ファイル

| ファイル | 役割 |
|---------|------|
| `Forms/SettingsForm.vb` | 設定フォーム |

---

## 技術的注意点

| 項目 | 内容 |
|-----|------|
| Outlook COM API | 実行時に Outlook が起動している必要あり |
| FTS5 | System.Data.SQLite のバンドルバイナリに含まれる。`USING fts5` が使えることを確認 |
| WebBrowser | IE エンジンベース。表示崩れが問題になれば CefSharp（+100MB）への切り替えを検討 |
| In-Reply-To/References | `MailItem.PropertyAccessor` でMAPIプロパティから取得 |
| 大量メール | ListView を VirtualMode で実装し、数万件でもスクロールがスムーズになるようにする |
| SQLite ネットワークドライブ | 単一ユーザーアクセス前提。同時アクセス不可 |
