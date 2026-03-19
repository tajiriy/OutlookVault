# フォルダツリー幅復元時の NullReferenceException

- 発生日: 2026-03-20
- 重要度: 高（起動時クラッシュ）

## 再現方法

アプリケーションを起動する。

## 原因

`MainForm_Load` でフォルダツリー幅の復元処理（`_settings.FolderTreeWidth`）を
`InitializeServices()` より前に配置したため、`_settings` が `Nothing` の状態で
アクセスして `NullReferenceException` が発生した。

`_settings` は `InitializeServices()` 内で `Config.AppSettings.Instance` に初期化される。

## 対応

フォルダツリー幅の復元処理を `InitializeServices()` の後に移動した。

## 所見

`MainForm_Load` 内でフィールドを参照する処理を追加する際は、
そのフィールドの初期化タイミング（`InitializeServices()` 前後）を確認すること。
`_settings`, `_repo` 等の主要フィールドは `InitializeServices()` で初期化される。
