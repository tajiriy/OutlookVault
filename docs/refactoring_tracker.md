# リファクタリングトラッカー

## サマリ

| ステータス | 件数 |
|-----------|------|
| open      | 10   |
| in-progress | 0  |
| done      | 21   |
| wontfix   | 1    |

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
| ステータス | done |
| 優先度 | Medium |
| カテゴリ | winforms |
| ソース | review |
| 対象ファイル | OutlookVault/Controls/EmailPreviewControl.vb |
| 登録日 | 2026-03-19 |
| 修正日 | 2026-03-19 |

**内容:** `LoadAttachments` で `flowAttachments.Controls.Clear()` はコントロール削除するが `AddHandler` のアンサブスクライブをしない。`Font` オブジェクト (`smallFont`) を毎回生成し `Dispose` が呼ばれない。

**対策:** `CleanupAttachmentControls` メソッドを追加し、`Controls.Clear()` 前に全 AddHandler の RemoveHandler + pnl.Dispose() を実施。`smallFont` を `_attachSmallFont` フィールドに昇格してインスタンス共有（遅延初期化）。`ClearPreview` でもクリーンアップを呼ぶ。

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
| ステータス | done |
| 優先度 | Medium |
| カテゴリ | readability |
| ソース | review |
| 対象ファイル | OutlookVault/Services/ImportService.vb |
| 登録日 | 2026-03-19 |
| 修正日 | 2026-03-19 |

**内容:** `ImportFolder` が差分スキャン判定・アイテムイテレーション・中間コミット管理・同期状態更新・ログ出力・パフォーマンスチューニングを 1 メソッドに集約（約 180 行）。ネストが深く変更時に副作用が生じやすい。

**対策:** 以下の4メソッドを抽出しImportFolderから呼び出し: `BuildItemsCollection`（差分/フルスキャン判定+Items取得）、`HandleMailItemError`（エラー処理・エラーID登録・ログ出力）、`UpdateSyncState`（同期状態DB更新）、`LogImportResult`（完了ログ出力）。ImoprtFolderは約130行→約80行に縮小。

**メモ:** ImportService.vb 行 115〜295

---

### R-011: ソート列がマジックナンバー（0〜4）で管理されている

| 項目 | 値 |
|------|-----|
| ステータス | done |
| 優先度 | Medium |
| カテゴリ | readability |
| ソース | review |
| 対象ファイル | OutlookVault/MainForm.vb |
| 登録日 | 2026-03-19 |
| 修正日 | 2026-03-19 |

**内容:** `SortEmailCache` で `Case 0`〜`Case 4` でソート列を判定。列インデックスの意味がコメントのみ。`listViewEmails_ColumnClick` の `e.Column <> 3` も受信日時を意味するが読み取りにくい。列追加時に 3 箇所を同期させる必要あり。

**対策:** `Enum EmailListColumn` (Attachment=0, Subject=1, Sender=2, ReceivedAt=3, Size=4) を定義。`_sortColumn` の型を `Integer` → `EmailListColumn` に変更。`SortEmailCache` の `Select Case`、`ColumnClick` の比較、`AppSettings` のデフォルト値をすべて Enum 参照に統一。設定ファイルとの互換性は `CInt`/`CType` 変換で維持。

**メモ:** MainForm.vb 行 1001〜1040、行 953

---

### R-012: NormalizeSubject でループ内に毎回 ToLower()

| 項目 | 値 |
|------|-----|
| ステータス | done |
| 優先度 | Low |
| カテゴリ | performance |
| ソース | review |
| 対象ファイル | OutlookVault/Services/ThreadingService.vb |
| 登録日 | 2026-03-19 |
| 修正日 | 2026-03-19 |

**内容:** `Do While changed` ループ内で `result.ToLower().StartsWith(prefix)` を毎回呼び、ループごと・プレフィックスごとに `ToLower()` が実行される。

**対策:** `result.ToLower().StartsWith(prefix)` → `result.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)` に変更。不要な文字列アロケーションを排除。

**メモ:** ThreadingService.vb 行 174〜178

---

### R-013: FormatRecipientsJson で毎回 New Regex を生成

| 項目 | 値 |
|------|-----|
| ステータス | done |
| 優先度 | Low |
| カテゴリ | performance |
| ソース | review |
| 対象ファイル | OutlookVault/Controls/EmailPreviewControl.vb |
| 登録日 | 2026-03-19 |
| 修正日 | 2026-03-19 |

**内容:** メール選択のたびに呼ばれる `FormatRecipientsJson` 内で毎回 `New Regex(...)` を生成。Regex コンパイルコストが蓄積する。

**対策:** `New Regex(...)` を `Private Shared ReadOnly RecipientJsonPattern` フィールド（`RegexOptions.Compiled` 付き）に昇格。メソッド内ではフィールド参照のみ。

**メモ:** EmailPreviewControl.vb 行 276

---

### R-014: JsonStr の JSON エスケープが不完全

| 項目 | 値 |
|------|-----|
| ステータス | done |
| 優先度 | Low |
| カテゴリ | readability |
| ソース | review |
| 対象ファイル | OutlookVault/Services/OutlookService.vb |
| 登録日 | 2026-03-19 |
| 修正日 | 2026-03-19 |

**内容:** `JsonStr` メソッドは `\` と `"` のみエスケープ。制御文字（`\t`、`\n`、`\r`）のエスケープがない。受信者名に改行が含まれると不正な JSON になる。

**対策:** `Replace(vbCr, "\r").Replace(vbLf, "\n").Replace(vbTab, "\t")` を追加。

**メモ:** OutlookService.vb 行 520〜523。現状は DB→UI の内部利用のみで実害は出にくい。

---

### R-015: ProcessMailItem のテストがない（COM 依存で単体テスト困難）

| 項目 | 値 |
|------|-----|
| ステータス | wontfix |
| 優先度 | Medium |
| カテゴリ | test-coverage |
| ソース | review |
| 対象ファイル | OutlookVault/Services/ImportService.vb |
| 登録日 | 2026-03-19 |
| 修正日 | - |

**内容:** `ProcessMailItem` はアプリのコアロジック（重複チェック・スレッド付与・DB 挿入・添付保存の統合）だが、`Outlook.MailItem` への直接依存でテストから除外されている。`errorIds` 追加・`skipReason` 分岐等の結合動作が未検証。

**対策:** wontfix — COM 依存と非依存のロジックが密結合しており、無理に分離すると可読性が低下する。IOutlookService インターフェース抽出は影響範囲が大きく現時点ではコスト対効果が見合わない。ThreadingService/EmailRepository 等の COM 非依存コンポーネントは既に 215 件のテストでカバーされている。将来アーキテクチャ全体を見直す際に再検討する。

**メモ:** R-010（ImportFolder 分割）と関連。

---

### R-016: AppSettings.StartWithWindows の Set で空 Catch（レジストリ書き込み失敗が無視）

| 項目 | 値 |
|------|-----|
| ステータス | done |
| 優先度 | Low |
| カテゴリ | error-handling |
| ソース | review |
| 対象ファイル | OutlookVault/Config/AppSettings.vb |
| 登録日 | 2026-03-19 |
| 修正日 | 2026-03-19 |

**内容:** `StartWithWindows` の Set プロパティで空 Catch。レジストリアクセス失敗（権限不足等）が無言で失敗し、ユーザーは設定が有効にならない原因が分からない。

**対策:** `Catch ex As SecurityException` と `Catch ex As UnauthorizedAccessException` で例外型を限定し、`Services.Logger.Warn` でログ出力。Get 側も `Catch ex As Exception` にコメント追加。

**メモ:** AppSettings.vb 行 211

### R-017: FindFolder での COM オブジェクト二重解放リスク

| 項目 | 値 |
|------|-----|
| ステータス | done |
| 優先度 | High |
| カテゴリ | resource-management |
| ソース | review |
| 対象ファイル | OutlookVault/Services/OutlookService.vb |
| 登録日 | 2026-03-19 |
| 修正日 | - |

**内容:** `FindFolder` で早期 Return パスと Finally ブロックの両方で `stores`/`root` の `Marshal.ReleaseComObject` が呼ばれ、二重解放が発生する。

**対策:** 早期 Return を Exit For + result 変数に変更し、stores/root/store の解放を Finally に一本化。root = Nothing で二重解放を防止。

**メモ:** R-001 対応時に導入されたエッジケース。修正日: 2026-03-19

---

### R-018: BuildItemsCollection 内の folder.Items COM オブジェクト解放漏れ

| 項目 | 値 |
|------|-----|
| ステータス | done |
| 優先度 | High |
| カテゴリ | resource-management |
| ソース | review |
| 対象ファイル | OutlookVault/Services/ImportService.vb |
| 登録日 | 2026-03-19 |
| 修正日 | - |

**内容:** 差分スキャン時に `folder.Items` で取得した中間 Items COM オブジェクトが `Restrict` 後に解放されない。呼び出し元 `ImportFolder` でも返された `items` に対して `Marshal.ReleaseComObject` を呼んでいない。

**対策:** 差分スキャン時に `Dim allItems = folder.Items` → `items = allItems.Restrict(filter)` → `ReleaseComObject(allItems)` で中間 Items を解放。`ImportFolder` の Finally に `If items IsNot Nothing Then Marshal.ReleaseComObject(items)` を追加。

**メモ:** R-010 でメソッド抽出した際に漏れた箇所。修正日: 2026-03-19

---

### R-019: ImportFolder でのバルクモード不整合（中間 BeginBulk 例外時）

| 項目 | 値 |
|------|-----|
| ステータス | done |
| 優先度 | High |
| カテゴリ | resource-management |
| ソース | review |
| 対象ファイル | OutlookVault/Services/ImportService.vb |
| 登録日 | 2026-03-19 |
| 修正日 | - |

**内容:** 中間コミット後の再 `BeginBulk()` が例外を投げた場合、`RollbackBulk` が呼ばれないケースがある。

**対策:** 中間コミット後の `BeginBulk` を Try/Catch で囲み、失敗時に Logger.Error で記録してから Throw。外側の Catch で RollbackBulk が呼ばれる。

**メモ:** 修正日: 2026-03-19

---

### R-020: SyncDeletions 内の folder.Items.Count で Items 未解放

| 項目 | 値 |
|------|-----|
| ステータス | done |
| 優先度 | Medium |
| カテゴリ | resource-management |
| ソース | review |
| 対象ファイル | OutlookVault/Services/ImportService.vb |
| 登録日 | 2026-03-19 |
| 修正日 | - |

**内容:** `folder.Items.Count` のメソッドチェーンで一時的に生成された `Outlook.Items` COM オブジェクトが解放されない。

**対策:** `Dim tmpItems = folder.Items` → `tmpItems.Count` → `Marshal.ReleaseComObject(tmpItems)` に分離して明示解放。

**メモ:** 修正日: 2026-03-19

---

### R-021: HtmlSanitizerService の正規表現が都度コンパイルされている

| 項目 | 値 |
|------|-----|
| ステータス | done |
| 優先度 | Medium |
| カテゴリ | performance |
| ソース | review |
| 対象ファイル | OutlookVault/Services/HtmlSanitizerService.vb |
| 登録日 | 2026-03-19 |
| 修正日 | - |

**内容:** `Sanitize` メソッド内の全 `Regex.Replace` 呼び出し（7箇所）が静的メソッドで呼び出しのたびにコンパイルされる。

**対策:** ブロックタグ用ペア/自己閉じパターンを `Dictionary(Of String, Regex)` で Compiled キャッシュ。他5パターンも `Private Shared ReadOnly` フィールドに昇格。Sanitize メソッドはフィールド参照のみに簡素化。

**メモ:** R-013 と同パターンの改善。修正日: 2026-03-19

---

### R-022: BuildCopyText のセルキー計算にマジックナンバー 100000

| 項目 | 値 |
|------|-----|
| ステータス | open |
| 優先度 | Medium |
| カテゴリ | readability |
| ソース | review |
| 対象ファイル | OutlookVault/Forms/TableViewerForm.vb |
| 登録日 | 2026-03-19 |
| 修正日 | - |

**内容:** `CLng(cell.RowIndex) * 100000L + CLng(cell.ColumnIndex)` のマジックナンバーに根拠がなく、将来の保守リスク。

**対策:** ビットシフト（`<< 16`）またはタプルキーの Dictionary に変更。

**メモ:** 初回レビューでも指摘あり。

---

### R-023: Logger.Error がスタックトレースを記録しない

| 項目 | 値 |
|------|-----|
| ステータス | done |
| 優先度 | Medium |
| カテゴリ | error-handling |
| ソース | review |
| 対象ファイル | OutlookVault/Services/Logger.vb |
| 登録日 | 2026-03-19 |
| 修正日 | - |

**内容:** `Error(message, ex)` が `ex.Message` のみ記録し、スタックトレースや InnerException が記録されない。

**対策:** `ex.GetType().Name & ": " & ex.Message` → `ex.ToString()` に変更。スタックトレース・InnerException が自動的に含まれる。テストにスタックトレース検証を追加。

**メモ:** 修正日: 2026-03-19

---

### R-024: AppSettings.TargetFolders が毎回 new List を生成

| 項目 | 値 |
|------|-----|
| ステータス | open |
| 優先度 | Medium |
| カテゴリ | performance |
| ソース | review |
| 対象ファイル | OutlookVault/Config/AppSettings.vb |
| 登録日 | 2026-03-19 |
| 修正日 | - |

**内容:** `TargetFolders` ゲッターが呼び出しのたびに ConfigurationManager を読み込み new List を生成。

**対策:** キャッシュするか、呼び出し側でローカル変数に保持するパターンの徹底をコメントで明示。

**メモ:** 現状は主要呼び出し箇所で1回取得しているが、他箇所の注意喚起が必要。

---

### R-025: AttachmentStatsForm.LoadData がアプリ側で全件集計

| 項目 | 値 |
|------|-----|
| ステータス | open |
| 優先度 | Medium |
| カテゴリ | performance |
| ソース | review |
| 対象ファイル | OutlookVault/Forms/AttachmentStatsForm.vb |
| 登録日 | 2026-03-19 |
| 修正日 | - |

**内容:** 全添付ファイルを SELECT して VB.NET 側で拡張子ごとに集計。大量件数で不要なデータ転送。

**対策:** SQL 側で GROUP BY + COUNT/SUM で集計し結果のみ返す。

**メモ:** なし

---

### R-026: ReplaceCidReferences の att.FilePath が相対パスのまま

| 項目 | 値 |
|------|-----|
| ステータス | open |
| 優先度 | Medium |
| カテゴリ | error-handling |
| ソース | review |
| 対象ファイル | OutlookVault/Controls/EmailPreviewControl.vb |
| 登録日 | 2026-03-19 |
| 修正日 | - |

**内容:** `att.FilePath` が DB の相対パスのまま `File.Exists` に渡されており、正しく動作しない可能性。

**対策:** `AttachmentDirectory` と結合して絶対パスに変換してから使用。

**メモ:** AttachmentSaveAs_Click でも同様。

---

### R-027: JsonStr の制御文字エスケープが不完全（U+0000〜U+001F）

| 項目 | 値 |
|------|-----|
| ステータス | open |
| 優先度 | Medium |
| カテゴリ | error-handling |
| ソース | review |
| 対象ファイル | OutlookVault/Services/OutlookService.vb |
| 登録日 | 2026-03-19 |
| 修正日 | - |

**内容:** `\r`/`\n`/`\t` 以外の制御文字（`\b`/`\f` 等）が未エスケープ。

**対策:** `System.Text.Json` か制御文字を `\uXXXX` に変換するループを追加。

**メモ:** R-014 で `\r`/`\n`/`\t` は対応済みだが、残りの制御文字が未対応。

---

### R-028: AppSettingsTests がデフォルト値を検証していない

| 項目 | 値 |
|------|-----|
| ステータス | open |
| 優先度 | Medium |
| カテゴリ | test-coverage |
| ソース | review |
| 対象ファイル | OutlookVault.Tests/AppSettingsTests.vb |
| 登録日 | 2026-03-19 |
| 修正日 | - |

**内容:** テスト名が「デフォルト値が False」と主張しているが、型チェックのみで実際の値を検証していない。

**対策:** デフォルト値を返すテスト用コンストラクタか IAppSettings インターフェースを導入。

**メモ:** なし

---

### R-029: ImportService のテストが存在しない

| 項目 | 値 |
|------|-----|
| ステータス | open |
| 優先度 | Medium |
| カテゴリ | test-coverage |
| ソース | review |
| 対象ファイル | OutlookVault/Services/ImportService.vb |
| 登録日 | 2026-03-19 |
| 修正日 | - |

**内容:** `UpdateSyncState`/`LogImportResult` 等の純ロジック部分がテスト可能だがテストファイルがない。

**対策:** `UpdateSyncState`/`LogImportResult` を `Friend` 公開してテスト追加。

**メモ:** R-015 (wontfix) はインターフェース抽出が必要だったが、抽出済みのサブメソッドは直接テスト可能。

---

### R-030: MainForm_FormClosing で Application.DoEvents を使用

| 項目 | 値 |
|------|-----|
| ステータス | open |
| 優先度 | Low |
| カテゴリ | winforms |
| ソース | review |
| 対象ファイル | OutlookVault/MainForm.vb |
| 登録日 | 2026-03-19 |
| 修正日 | - |

**内容:** スピンウェイトループで `Application.DoEvents()` を使っており、リエントラント問題のリスクがある。

**対策:** `async/await` パターンまたは `Task.Delay` ポーリングに変更。

**メモ:** なし

---

### R-031: GetTableData で全件を DataTable にロード

| 項目 | 値 |
|------|-----|
| ステータス | open |
| 優先度 | Low |
| カテゴリ | performance |
| ソース | review |
| 対象ファイル | OutlookVault/Data/DatabaseManager.vb |
| 登録日 | 2026-03-19 |
| 修正日 | - |

**内容:** `attachments` 等のテーブルで `SELECT *` 全件取得。大量件数でメモリ消費。

**対策:** ページングまたは件数上限を導入。

**メモ:** なし

---

### R-032: FormatEmailSize / FormatFileSize の重複実装

| 項目 | 値 |
|------|-----|
| ステータス | open |
| 優先度 | Low |
| カテゴリ | readability |
| ソース | review |
| 対象ファイル | OutlookVault/MainForm.vb, OutlookVault/Controls/EmailPreviewControl.vb, OutlookVault/Forms/AttachmentStatsForm.vb |
| 登録日 | 2026-03-19 |
| 修正日 | - |

**内容:** バイト数→人間可読文字列の変換が3箇所に重複実装。

**対策:** 共通ヘルパークラスに統合。

**メモ:** なし

---

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
| 2026-03-19 | R-011 | done: Enum EmailListColumn でソート列のマジックナンバーを排除 |
| 2026-03-19 | R-012, R-013 | done: StringComparison.OrdinalIgnoreCase 使用、Regex を Shared Compiled に昇格 |
| 2026-03-19 | R-014, R-016 | done: JSON エスケープに制御文字追加、レジストリ Catch を例外型限定+Logger.Warn |
| 2026-03-19 | R-006 | done: 動的コントロールの RemoveHandler+Dispose、Font フィールド昇格 |
| 2026-03-19 | R-010 | done: ImportFolder を 4 サブメソッドに分割（180行→80行） |
| 2026-03-19 | R-015 | wontfix: COM 密結合で分離コスト高、既存テスト 215 件で十分カバー |
| 2026-03-19 | R-017〜R-032 | 2回目の code-reviewer レビューから 16 件を一括登録 |
| 2026-03-19 | R-017, R-018, R-019 | done: COM 二重解放修正、Items 解放漏れ修正、BeginBulk 例外処理追加 |
| 2026-03-19 | R-020, R-021 | done: SyncDeletions Items 解放、HtmlSanitizer Regex Compiled 化 |
| 2026-03-19 | R-023 | done: Logger.Error に ex.ToString() でスタックトレースを記録 |
