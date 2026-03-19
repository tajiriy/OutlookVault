# リファクタリングトラッカー

## サマリ

| ステータス | 件数 |
|-----------|------|
| open      | 8    |
| in-progress | 0  |
| done      | 8    |
| wontfix   | 0    |

## カテゴリ

| カテゴリ | 説明 |
|---------|------|
| resource-management | IDisposable / COM 解放 |
| sql-safety | SQLite パラメータ化・インジェクション対策 |
| winforms | UI スレッド安全性・Designer.vb 分離 |
| vbnet-convention | Option Strict / 命名規則 / 型安全性 |
| performance | ListView 最適化・文字列連結・非同期処理 |
| error-handling | 例外処理・COMException ハンドリング |
| test-coverage | テストケースの追加・エッジケース |
| readability | コメント・命名・責務分離 |

## 項目一覧

### R-001: OutlookService の COM オブジェクト解放漏れ（フォルダ・アイテム列挙）

| 項目 | 値 |
|------|-----|
| ステータス | done |
| 優先度 | High |
| カテゴリ | resource-management |
| ソース | review |
| 対象ファイル | OutlookVault/Services/OutlookService.vb |
| 登録日 | 2026-03-19 |
| 修正日 | 2026-03-19 |

**内容:** `GetAvailableFolderNames`、`CollectFolderNames`、`SearchFolder`、`GetFolderMessageIds`、`SaveAttachments` 等で `Outlook.Stores`、`MAPIFolder`、`Folders`、`Items`、`Attachments`、`Attachment` の COM オブジェクトが `Marshal.ReleaseComObject` で解放されていない。大量メール処理時にメモリ増大や COM ランタイムエラーの原因になり得る。

**対策:** 全14メソッドで COM オブジェクトの `Try...Finally Marshal.ReleaseComObject` パターンを適用。対象: GetAvailableFolderNames, GetExcludedFolderEntryIds, CollectFolderNames, FindFolder, SearchFolder, GetMailItemCount, GetFolderMessageIds, ExtractMessageId, ExtractEmailData (pa/recipients/parentFolder/mailAtts/att), SaveAttachments, ResolveExchangeAddress (entry/exUser), SerializeRecipients (r/addrEntry)。

**メモ:** Dispose で `_ns` のみ解放し `_app` は意図的に残している設計は維持する。R-007 も同時完了。

---

### R-002: SetSynchronousMode の mode パラメータがホワイトリスト未検証

| 項目 | 値 |
|------|-----|
| ステータス | done |
| 優先度 | High |
| カテゴリ | sql-safety |
| ソース | review |
| 対象ファイル | OutlookVault/Data/DatabaseManager.vb |
| 登録日 | 2026-03-19 |
| 修正日 | 2026-03-19 |

**内容:** `ExecuteNonQuery(conn, "PRAGMA synchronous=" & mode & ";")` で文字列を直接連結している。現状は固定文字列 `"OFF"` / `"NORMAL"` のみだが、PRAGMA はパラメータ化クエリが使えないためホワイトリスト検証が必要。

**対策:** `Enum SynchronousMode` (Off/Normal/Full) を `Data` 名前空間に追加。`SetSynchronousMode` のパラメータを `String` → `SynchronousMode` に変更し、内部で `mode.ToString().ToUpperInvariant()` で安全に文字列変換。呼び出し元 (ImportService) とテストも Enum 使用に更新。Full モードのテストを追加。

**メモ:** ImportService.vb 行 171, 288 から呼び出されている。

---

### R-003: EmailRepository が IDisposable を実装していない

| 項目 | 値 |
|------|-----|
| ステータス | done |
| 優先度 | High |
| カテゴリ | resource-management |
| ソース | review |
| 対象ファイル | OutlookVault/Data/EmailRepository.vb |
| 登録日 | 2026-03-19 |
| 修正日 | 2026-03-19 |

**内容:** `_bulkConn`・`_bulkTx`・`_bulkEmailCmd`・`_bulkAttachCmd` をフィールド保持しているが `IDisposable` 未実装。`CommitBulk` / `RollbackBulk` を呼ばずに例外で抜けた場合、リソースが解放されない。

**対策:** `IDisposable` を実装。`Dispose(Boolean)` パターンで `RollbackBulk()` を呼び、未コミットのバルクリソースを確実に解放。`_disposed` フラグで二重呼び出しを防止。`MainForm.Designer.vb` の `Dispose` に `_repo.Dispose()` を追加してフォーム終了時にも解放。

**メモ:** なし

---

### R-004: DeleteSelectedEmails で N+1 クエリ（既存の一括削除メソッドが未使用）

| 項目 | 値 |
|------|-----|
| ステータス | done |
| 優先度 | High |
| カテゴリ | performance |
| ソース | review |
| 対象ファイル | OutlookVault/MainForm.vb |
| 登録日 | 2026-03-19 |
| 修正日 | 2026-03-19 |

**内容:** 選択メール削除時に各 `emailId` ごとに `GetAttachmentsByEmailId`（SELECT）と `DeleteEmail`（DELETE）を個別発行。100件選択で 300+ 回の DB ラウンドトリップ。既存の `DeleteEmailsByIds` 一括削除メソッドが使われていない。

**対策:** `DeleteSelectedEmails` のループを `DeleteEmailsByIds(selectedIds)` 1回の呼び出しに置き換え。戻り値の添付ファイルパスリストで物理削除を実行。DB 操作が単一トランザクションに集約され、N+1 問題を解消。

**メモ:** MainForm.vb 行 758〜773

---

### R-005: UpdateFolderCountsAsync のフォルダ別件数取得が N+1 クエリ

| 項目 | 値 |
|------|-----|
| ステータス | done |
| 優先度 | Medium |
| カテゴリ | performance |
| ソース | review |
| 対象ファイル | OutlookVault/MainForm.vb |
| 登録日 | 2026-03-19 |
| 修正日 | 2026-03-19 |

**内容:** `GetFolderNames()` で N 件取得後、フォルダごとに `GetTotalCount(folder)` を呼ぶループで N 回の SELECT を発行。`LoadFolderTree` でも同様のパターン。

**対策:** `GetFolderCounts() As Dictionary(Of String, Integer)` を `EmailRepository` に追加（GROUP BY で1クエリ取得 + NULL件数加算）。`LoadFolderTree` と `UpdateFolderCountsAsync` の `GetFolderNames` + `GetTotalCount` ループを `GetFolderCounts` 1回の呼び出しに置き換え。

**メモ:** MainForm.vb 行 848〜855、行 185〜187

---

### R-006: EmailPreviewControl の動的コントロールでイベントハンドラが累積・Font 未 Dispose

| 項目 | 値 |
|------|-----|
| ステータス | open |
| 優先度 | Medium |
| カテゴリ | winforms |
| ソース | review |
| 対象ファイル | OutlookVault/Controls/EmailPreviewControl.vb |
| 登録日 | 2026-03-19 |
| 修正日 | - |

**内容:** `LoadAttachments` で `flowAttachments.Controls.Clear()` はコントロール削除するが `AddHandler` のアンサブスクライブをしない。`Font` オブジェクト (`smallFont`) を毎回生成し `Dispose` が呼ばれない。

**対策:** `smallFont` をフィールドに昇格してインスタンス共有する。`Controls.Clear()` 前に各コントロールの `Dispose` を呼ぶか、`RemoveHandler` を実施する。

**メモ:** EmailPreviewControl.vb 行 301〜374、行 344

---

### R-007: PropertyAccessor COM オブジェクト未解放

| 項目 | 値 |
|------|-----|
| ステータス | done |
| 優先度 | Medium |
| カテゴリ | resource-management |
| ソース | review |
| 対象ファイル | OutlookVault/Services/OutlookService.vb |
| 登録日 | 2026-03-19 |
| 修正日 | 2026-03-19 |

**内容:** `IsFolderHidden`、`GetContainerClass`、`GetAttachmentContentId` で `folder.PropertyAccessor` / `att.PropertyAccessor` を取得しているが `Marshal.ReleaseComObject` が呼ばれていない。

**対策:** R-001 と同時対応。`IsFolderHidden`、`GetContainerClass`、`GetAttachmentContentId`、`ExtractMessageId`、`ExtractEmailData` の各 PropertyAccessor に `Try...Finally Marshal.ReleaseComObject(pa)` を適用。

**メモ:** R-001 と同時完了。

---

### R-008: OutlookService の空 Catch ブロック多数

| 項目 | 値 |
|------|-----|
| ステータス | done |
| 優先度 | Medium |
| カテゴリ | error-handling |
| ソース | review |
| 対象ファイル | OutlookVault/Services/OutlookService.vb |
| 登録日 | 2026-03-19 |
| 修正日 | 2026-03-19 |

**内容:** `GetExcludedFolderEntryIds`、`IsFolderHidden`、`GetContainerClass`、`GetMAPIString`、`ResolveExchangeAddress`、`GetAttachmentContentId` に空 Catch ブロックまたはコメントのみの Catch が存在。Exchange 問い合わせ失敗等の原因追跡が困難。

**対策:** 全6箇所の空 Catch を修正。`Catch` → `Catch ex As COMException` に例外型を限定し、意図的に無視する理由をコメントで明示。`ResolveExchangeAddress` は `Logger.Warn` でログ出力を追加。

**メモ:** OutlookService.vb 行 117〜122、行 490 等

---

### R-009: MainForm 非同期処理の空 Catch ブロック

| 項目 | 値 |
|------|-----|
| ステータス | done |
| 優先度 | Medium |
| カテゴリ | error-handling |
| ソース | review |
| 対象ファイル | OutlookVault/MainForm.vb |
| 登録日 | 2026-03-19 |
| 修正日 | 2026-03-19 |

**内容:** `UpdateStatusBarAsync` 内で `GetTotalCount()` と `GetLastImportDate()` の Try-Catch が両方とも空。`UpdateFolderCountsAsync` も同様。DB 接続エラーがユーザーに通知されず原因特定が困難。

**対策:** 全3箇所の空 Catch を `Catch ex As Exception` + `Services.Logger.Warn(メッセージ & ex.Message)` に変更。UI 通知は不要だがログファイルにエラー情報を記録。

**メモ:** MainForm.vb 行 911〜918、行 854

---

### R-010: ImportFolder メソッドの責務過多（約 170 行）

| 項目 | 値 |
|------|-----|
| ステータス | open |
| 優先度 | Medium |
| カテゴリ | readability |
| ソース | review |
| 対象ファイル | OutlookVault/Services/ImportService.vb |
| 登録日 | 2026-03-19 |
| 修正日 | - |

**内容:** `ImportFolder` が差分スキャン判定・アイテムイテレーション・中間コミット管理・同期状態更新・ログ出力・パフォーマンスチューニングを 1 メソッドに集約（約 180 行）。ネストが深く変更時に副作用が生じやすい。

**対策:** `BuildItemsCollection`（Restrict フィルタ判定）、`IterateAndImport`（ループ・中間コミット）、`UpdateSyncState`、`LogImportResult` に分割。各 30〜50 行程度に。

**メモ:** ImportService.vb 行 115〜295

---

### R-011: ソート列がマジックナンバー（0〜4）で管理されている

| 項目 | 値 |
|------|-----|
| ステータス | open |
| 優先度 | Medium |
| カテゴリ | readability |
| ソース | review |
| 対象ファイル | OutlookVault/MainForm.vb |
| 登録日 | 2026-03-19 |
| 修正日 | - |

**内容:** `SortEmailCache` で `Case 0`〜`Case 4` でソート列を判定。列インデックスの意味がコメントのみ。`listViewEmails_ColumnClick` の `e.Column <> 3` も受信日時を意味するが読み取りにくい。列追加時に 3 箇所を同期させる必要あり。

**対策:** `Enum EmailListColumn` を定義（`Attachment = 0`、`Subject = 1` 等）し、`_sortColumn`・設定値・`Select Case` をすべて Enum 参照に統一。

**メモ:** MainForm.vb 行 1001〜1040、行 953

---

### R-012: NormalizeSubject でループ内に毎回 ToLower()

| 項目 | 値 |
|------|-----|
| ステータス | open |
| 優先度 | Low |
| カテゴリ | performance |
| ソース | review |
| 対象ファイル | OutlookVault/Services/ThreadingService.vb |
| 登録日 | 2026-03-19 |
| 修正日 | - |

**内容:** `Do While changed` ループ内で `result.ToLower().StartsWith(prefix)` を毎回呼び、ループごと・プレフィックスごとに `ToLower()` が実行される。

**対策:** `StartsWith(prefix, StringComparison.OrdinalIgnoreCase)` に変更。

**メモ:** ThreadingService.vb 行 174〜178

---

### R-013: FormatRecipientsJson で毎回 New Regex を生成

| 項目 | 値 |
|------|-----|
| ステータス | open |
| 優先度 | Low |
| カテゴリ | performance |
| ソース | review |
| 対象ファイル | OutlookVault/Controls/EmailPreviewControl.vb |
| 登録日 | 2026-03-19 |
| 修正日 | - |

**内容:** メール選択のたびに呼ばれる `FormatRecipientsJson` 内で毎回 `New Regex(...)` を生成。Regex コンパイルコストが蓄積する。

**対策:** パターンを `Private Shared ReadOnly` フィールド（`RegexOptions.Compiled` 付き）に昇格。

**メモ:** EmailPreviewControl.vb 行 276

---

### R-014: JsonStr の JSON エスケープが不完全

| 項目 | 値 |
|------|-----|
| ステータス | open |
| 優先度 | Low |
| カテゴリ | readability |
| ソース | review |
| 対象ファイル | OutlookVault/Services/OutlookService.vb |
| 登録日 | 2026-03-19 |
| 修正日 | - |

**内容:** `JsonStr` メソッドは `\` と `"` のみエスケープ。制御文字（`\t`、`\n`、`\r`）のエスケープがない。受信者名に改行が含まれると不正な JSON になる。

**対策:** `vbCrLf`・`vbCr`・`vbLf`・タブを JSON エスケープシーケンスに置換する処理を追加。

**メモ:** OutlookService.vb 行 520〜523。現状は DB→UI の内部利用のみで実害は出にくい。

---

### R-015: ProcessMailItem のテストがない（COM 依存で単体テスト困難）

| 項目 | 値 |
|------|-----|
| ステータス | open |
| 優先度 | Medium |
| カテゴリ | test-coverage |
| ソース | review |
| 対象ファイル | OutlookVault/Services/ImportService.vb |
| 登録日 | 2026-03-19 |
| 修正日 | - |

**内容:** `ProcessMailItem` はアプリのコアロジック（重複チェック・スレッド付与・DB 挿入・添付保存の統合）だが、`Outlook.MailItem` への直接依存でテストから除外されている。`errorIds` 追加・`skipReason` 分岐等の結合動作が未検証。

**対策:** COM 非依存のロジック部分を `Models.Email` を受け取るメソッドに分離し、NUnit でテスト可能にする。将来的には `IOutlookService` インターフェース抽出も検討。

**メモ:** R-010（ImportFolder 分割）と関連。

---

### R-016: AppSettings.StartWithWindows の Set で空 Catch（レジストリ書き込み失敗が無視）

| 項目 | 値 |
|------|-----|
| ステータス | open |
| 優先度 | Low |
| カテゴリ | error-handling |
| ソース | review |
| 対象ファイル | OutlookVault/Config/AppSettings.vb |
| 登録日 | 2026-03-19 |
| 修正日 | - |

**内容:** `StartWithWindows` の Set プロパティで空 Catch。レジストリアクセス失敗（権限不足等）が無言で失敗し、ユーザーは設定が有効にならない原因が分からない。

**対策:** `Catch ex As UnauthorizedAccessException` 等で具体的例外を捕捉し、`Logger.Warn` でログ出力。呼び出し元の `SettingsForm` でユーザー通知を検討。

**メモ:** AppSettings.vb 行 211

## 変更履歴

| 日付 | 項目 | 変更内容 |
|------|------|---------|
| 2026-03-19 | - | トラッカー新規作成 |
| 2026-03-19 | R-001〜R-016 | code-reviewer による初回全体レビューから 16 件を一括登録 |
| 2026-03-19 | R-002 | done: String→Enum SynchronousMode に変更、テスト追加 |
| 2026-03-19 | R-003 | done: IDisposable 実装、MainForm.Dispose で解放 |
| 2026-03-19 | R-004 | done: DeleteSelectedEmails を DeleteEmailsByIds に置き換え |
| 2026-03-19 | R-001, R-007 | done: OutlookService 全14メソッドの COM オブジェクト解放を一括対応 |
| 2026-03-19 | R-005 | done: GetFolderCounts で N+1 を 1 クエリに集約 |
| 2026-03-19 | R-008, R-009 | done: 空 Catch を COMException 限定+コメント明示+Logger.Warn に修正 |
