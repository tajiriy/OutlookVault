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

## Phase 6: 検索・削除・設定 ✅

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

## Phase 6 以降の追加機能

### メール一覧: 添付アイコン・サイズ列

- **添付アイコン**: `ImageList` + GDI+ 描画のペーパークリップアイコンを添付列（index 0）に表示
  - WinForms ListView の `SmallImageList` は常に列 0 に描画される仕様のため、添付列を `Columns.Insert(0, ...)` で先頭に挿入することで件名と分離
- **サイズ列**: `MailItem.Size`（バイト）を DB の `email_size` 列に保存し、一覧に KB/MB 単位で表示
  - 1 MB 以上: `x.x MB`、未満: `x KB`
- **DBマイグレーション**: `DatabaseManager.ApplyMigrations()` で既存 DB に `ALTER TABLE emails ADD COLUMN email_size` を適用

### メール一覧: 列操作

- **列ソート**: 列ヘッダークリックで昇順/降順切り替え。ヘッダーテキストに `↑` / `↓` を付与
  - デフォルト: 受信日時 降順
  - 検索・フォルダ切り替え後も現在のソート設定を維持
- **列並び替え**: `AllowColumnReorder = True` でドラッグ＆ドロップによる列移動
- **設定の永続化**: アプリ終了時（`FormClosing`）に列幅・列表示順・ソート設定を `App.config` へ保存し、次回起動時に復元
  - `AppSettings` に `EmailListColumnWidths` / `EmailListColumnOrder` / `EmailListSortColumn` / `EmailListSortAscending` を追加

### 添付ファイル保存バグ修正

- **原因**: `OutlookService.SaveAttachments` が相対パスを Outlook COM の `SaveAsFile` に渡していた。`Directory.CreateDirectory` は .NET プロセスの作業ディレクトリで成功（フォルダのみ作成）するが、Outlook COM は Outlook 自身の作業ディレクトリを基準にパスを解釈するため保存失敗
- **修正**: `Path.GetFullPath(...)` で絶対パスに変換してから `SaveAsFile` に渡す
- **その他の改善**:
  - ディレクトリ作成を遅延化（実際に保存するファイルがある場合のみ作成、空フォルダを残さない）
  - `HasAttachments` フラグを OLE 埋め込みオブジェクト（メール署名の画像等）を除いたカウントで判定するよう修正
  - 保存エラーを `ImportResult.Errors` に追加して取り込み結果ダイアログに表示
  - `EmailRepository.GetAttachmentsByEmailId` で DB の相対パスを絶対パスに変換して返す

---

## 取り込み順序の設定・重複スキップ高速化

### 取り込み順序

- `AppSettings` に `ImportOldestFirst` プロパティを追加（デフォルト: `True` = 古い順）
- `ImportService.ImportFolder` でループ方向を設定値に応じて切り替え（`1→N` or `N→1`）
- 設定画面の「自動取り込み」グループにドロップダウン追加（「古い順（推奨）」/「新しい順」）

### 重複スキップ高速化

- **変更前**: 1件ごとに `MessageIdExists()` + `IsMessageIdDeleted()` の2回SQLクエリ（N件で2N回のDB問い合わせ）
- **変更後**: `ImportFolders` 開始時に `GetAllMessageIds()` / `GetAllDeletedMessageIds()` で全IDを `HashSet(Of String)` に一括ロード。以降はメモリ内 `Contains()` で判定
- 新規取り込み分は即座にキャッシュに追加し、同一セッション内の重複も防止

#### スキップ時の COM 操作削減

- **変更前**: スキップ対象でも `ExtractEmailData()` で全プロパティ（本文・受信者・添付ファイル等）を取得してから重複判定
- **変更後**: `OutlookService.ExtractMessageId()` で MessageID のみ軽量取得 → 重複判定 → 新規メールのみ `ExtractEmailData()` でフルデータ抽出
- スキップ時の COM 呼び出しが 15回以上 → 1〜2回に削減され、取り込み済みメールの走査が大幅に高速化

#### Exchange アドレス解決のキャッシュ化

- `OutlookService` に `_exchangeSmtpCache`（`Dictionary(Of String, String)`）を追加
- `GetSenderEmail` / `SerializeRecipients` 内の `GetExchangeUser()` 呼び出しを `ResolveExchangeAddress()` 経由に統一
- 同じ EX アドレスに対する2回目以降のネットワーク往復を回避。社内メールの多い環境で効果大
- **永続化**: `exchange_address_cache` テーブル（`ex_address` PRIMARY KEY, `smtp_address`, `cached_at`）を SQLite に追加
- 取り込み開始時に DB → メモリにロード、取り込み終了時にメモリ → DB に `INSERT OR IGNORE` で書き戻し
- アプリ再起動後も過去に解決済みの Exchange アドレスはキャッシュヒットし、ネットワーク往復が不要

#### 添付ファイル OLE チェックの短絡評価

- `mailItem.Attachments.Count = 0` の場合、OLE チェックループをスキップして即 `HasAttachments = False`

### 設定画面の追加機能

- **データ初期化**: 設定画面に「データ管理」グループを追加。「データ初期化...」ボタンで DB ファイルと添付ファイルディレクトリを一括削除（確認ダイアログ付き）
- **結果ダイアログ非表示**: `AppSettings` に `ShowImportResult` プロパティ追加（デフォルト: `True`）。表示設定グループに「取り込み完了時に結果ダイアログを表示する」チェックボックスを追加。OFF でも エラー発生時は常に表示

---

## インライン画像対応（HTMLメール cid: 参照の解決）

### 問題

HTMLメール中のインライン画像は `<img src="cid:image001@xxxxx">` 形式で参照されている。WebBrowserコントロールはこの `cid:` URL を解決できないため「×」表示になっていた。

### 対応方針

- 取り込み時: 添付ファイルの MAPI プロパティ `PR_ATTACH_CONTENT_ID`（0x3712001F）を参照し、値があればインライン画像として識別・保存
- 表示時: HTML中の `cid:xxx` を `file:///` URL に置換してからWebBrowserにセット
- 添付一覧: `IsInline=True` の添付ファイルは表示しない

### DB スキーマ変更（マイグレーション）

```sql
ALTER TABLE attachments ADD COLUMN content_id TEXT;
ALTER TABLE attachments ADD COLUMN is_inline  INTEGER DEFAULT 0;
```

### 変更ファイル

| ファイル | 変更内容 |
|---------|---------|
| `Models/Attachment.vb` | `ContentId`・`IsInline` プロパティ追加 |
| `Data/DatabaseManager.vb` | マイグレーションで `content_id` / `is_inline` カラム追加 |
| `Data/EmailRepository.vb` | `InsertAttachment` / `MapAttachment` を新カラム対応 |
| `Services/OutlookService.vb` | `GetAttachmentContentId()` ヘルパー追加。インライン画像の ContentId を取得・保存。`HasAttachments` カウントからインライン除外 |
| `Controls/EmailPreviewControl.vb` | `ReplaceCidReferences()` で `cid:` → `file:///` 置換。添付一覧からインライン非表示 |

### 注意事項

既存取り込み済みメールはインライン画像が `IsInline=False` のまま残るため、該当メールは削除して再取り込みが必要。

---

## 設定バグ修正: 結果ダイアログが設定に関わらず表示される

### 問題

「取り込み完了時に結果ダイアログを表示する」をオフにしても結果ダイアログが表示されていた。

### 原因

`MainForm.RunImportAsync()` 内で `MessageBox.Show` を `ShowImportResult` 設定で分岐させる処理が欠落していた。

### 修正

`ShowImportResult` が False かつエラーなしの場合はダイアログを表示しないよう条件を追加。エラーがある場合は設定に関わらず常に表示する。

---

## データ初期化の再起動不要化

### 変更前

データ初期化（DB削除・添付フォルダ削除）後に「アプリを再起動してください」と案内していた。

### 変更後

`SettingsForm` に `DataWasReset As Boolean` プロパティを追加。初期化成功時にフラグをセットしてダイアログを閉じる。`MainForm` 側で `DataWasReset = True` を検出すると `ReinitializeApp()` を呼び出し、`InitializeServices()` の再実行と UI 全リセット（フォルダツリー・メール一覧・プレビュー・検索クエリ）を行う。

---

## 取り込みパフォーマンス改善

### 背景

古い PST からの取り込みで約 5分/100件 かかっていた。Outlook COM の `Body`/`HTMLBody` 取得は Outlook 側の制約で改善困難だが、DB 書き込みとスレッド判定に起因する遅延は大幅に改善できる。

### ① DB 書き込みのバルクトランザクション化

- **変更前**: `InsertEmail` / `InsertAttachment` が 1件ごとに接続を開き自動コミット（SQLite は毎回 fsync）
- **変更後**: `EmailRepository` に `BeginBulk()` / `CommitBulk()` / `RollbackBulk()` を追加。`ImportFolder` 全体を共有トランザクションで包み、50件ごとに中間コミット（クラッシュ時のデータロスを 50件以内に限定）

### ② スレッド判定のインメモリキャッシュ化

- **変更前**: 1件ごとに `MessageId → ThreadId` / `NormalizedSubject → ThreadId` を DB クエリ（最大 3クエリ/件）
- **変更後**: 取り込み開始前に `EmailRepository.GetThreadIdCaches()` で全件を Dictionary にロード。`ThreadingService.LoadCaches()` でセット後、セッション中はメモリ参照のみ。新規割り当てはリアルタイムでキャッシュに追記。終了後は `ClearCaches()` で解放

### ③ 添付 ContentID の二重取得を除去

`ExtractEmailData` での HasAttachments チェックで ContentID 取得（PropertyAccessor）を行っていたが、同じ処理が `SaveAttachments` でも行われていた。`ExtractEmailData` 側を OLE 除外のみに簡略化し COM 呼び出しを削減。

### ④ PRAGMA synchronous=OFF（取り込み中のみ）

- **変更前**: WAL モードのデフォルト `synchronous=NORMAL` で、コミットごとに fsync が走る
- **変更後**: `ImportFolder` 開始時に `synchronous=OFF` に設定し fsync を省略。取り込み完了後（Finally）に `synchronous=NORMAL` に復元。OS クラッシュ時のみデータ破損リスクあり（アプリクラッシュは問題なし）

### ⑤ FTS4 トリガーの一時無効化

- **変更前**: INSERT ごとに FTS4 同期トリガー（emails_fts / attachments_fts）が発火し、全文検索インデックスを更新
- **変更後**: `ImportFolder` 開始時にトリガーを DROP、取り込み完了後に FTS4 の `rebuild` コマンドで一括再構築してからトリガーを再作成。バルク INSERT 時のオーバーヘッドを大幅に削減

### ⑥ Prepared Statement の再利用

- **変更前**: `InsertEmail` / `InsertAttachment` が呼び出しごとに `SQLiteCommand` を new し、SQL パース・コンパイルが毎回発生
- **変更後**: `BeginBulk()` 時に Email / Attachment 用の Prepared Statement を1回だけ作成。バルクモード中はパラメータ値の入れ替えのみで再利用。`CommitBulk()` / `RollbackBulk()` 時に Dispose

### ⑦ BulkCommitInterval の増加

- **変更前**: 50件ごとに中間コミット
- **変更後**: 200件ごとに中間コミット（fsync 回数を 1/4 に削減）

---

## テーブルビューア: 高度なフィルタ

### 概要

テーブルビューアのフィルタ機能を拡張。列指定・部分一致/完全一致・AND/OR 論理演算をサポート。

### フィルタ構文

| 入力形式 | 動作 | 例 |
|---------|------|-----|
| `テキスト` | 全文字列列で部分一致 (OR) | `amazon` |
| `列名: 値` | 指定列で部分一致 (LIKE) | `subject: amazon` |
| `列名= 値` | 指定列で完全一致 (=) | `id= 1` |
| `条件 and 条件` | AND 結合 | `subject: amazon and sender_name: john` |
| `条件 or 条件` | OR 結合 | `id= 1 or id= 2` |

- **演算子優先順位**: AND > OR（`a AND b OR c` → `(a AND b) OR c`）
- **列名**: 大文字小文字無視、完全一致
- **値**: 大文字小文字無視
- **括弧**: 非サポート

### 追加ファイル

| ファイル | 役割 |
|---------|------|
| `Filters/FilterParser.vb` | フィルタ文字列を `DataTable.RowFilter` 式に変換するパーサー |

### 変更ファイル

| ファイル | 変更内容 |
|---------|---------|
| `Forms/TableViewerForm.vb` | `ApplyFilter()` を `FilterParser.Parse()` 呼び出しに差し替え |

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
