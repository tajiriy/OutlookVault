# 20260318-001: SQLiteException 'no such module: fts5'

## 発生日
2026-03-18

## エラーメッセージ
```
System.Data.SQLite.SQLiteException: 'SQL logic error
no such module: fts5'
```

## 再現方法
1. アプリを起動する
2. `DatabaseManager.Initialize()` → `CreateFtsTables()` 内で FTS5 仮想テーブルを作成しようとするとクラッシュ

## 原因
`System.Data.SQLite` NuGet パッケージ (1.0.118.0) に付属する `SQLite.Interop.dll` が FTS5 拡張機能を**無効**にしてビルドされていた。
同パッケージの公式ドキュメントでは FTS5 は含まれると記載されているが、
環境によっては `stub.system.data.sqlite.core.netframework` 経由で提供される
インタロップ DLL が FTS5 非対応のバージョンになる場合がある。

## 対応方法
FTS5 の代わりに **FTS4** を使用するよう変更した。

### `Data/DatabaseManager.vb`
- `CreateFtsTables()`: `USING fts5(...)` → `USING fts4(content="...")` に変更
  - `content_rowid=` オプションは FTS4 不要のため削除（rowid を自動使用）
- `CreateFtsTriggers()`:
  - 削除・更新トリガーの FTS5 専用構文 `INSERT INTO fts(fts, rowid, ...) VALUES ('delete', ...)` を廃止
  - FTS4 対応の `DELETE FROM emails_fts WHERE rowid = old.id` に変更
  - 複数 SQL を1つの文字列で渡すと `ExecuteNonQuery` が途中で止まることがあるため、トリガーごとに個別呼び出しに変更

### `Data/EmailRepository.vb`
- `SearchEmails()`: FTS5 専用の `ORDER BY rank` → `ORDER BY e.received_at DESC` に変更

## 所見
- FTS4 は SQLite 3.7.4 以降で利用可能で、事実上すべての SQLite ビルドに含まれる
- FTS5 と FTS4 の基本的な `MATCH` クエリ構文は互換性があり、検索機能への影響なし
- 将来的に確実に FTS5 が含まれる SQLite ライブラリ（例: `Microsoft.Data.Sqlite`）への移行を検討してもよい
