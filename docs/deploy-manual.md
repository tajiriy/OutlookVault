# OutlookArchiver デプロイマニュアル

## 目次

1. [概要](#概要)
2. [デプロイスクリプトによるデプロイ](#デプロイスクリプトによるデプロイ)
3. [手動デプロイ](#手動デプロイ)
4. [データの移行](#データの移行)
5. [デプロイ後の確認](#デプロイ後の確認)
6. [アップデート](#アップデート)
7. [トラブルシューティング](#トラブルシューティング)

---

## 概要

OutlookArchiver を開発環境から任意のフォルダ（例: `C:\Tools\OutlookArchiver`）にデプロイする手順を説明します。

### 前提条件

- Windows 10 以降
- .NET Framework 4.6.2 がインストール済み
- Microsoft Outlook がインストール済み（メール取り込み時に必要）

### デプロイ先のフォルダ構成

```
C:\Tools\OutlookArchiver\
├── OutlookArchiver.exe          # 実行ファイル
├── OutlookArchiver.exe.Config   # アプリケーション設定
├── EntityFramework.dll          # 依存ライブラリ
├── EntityFramework.SqlServer.dll
├── Microsoft.Office.Interop.Outlook.dll
├── System.Data.SQLite.dll
├── System.Data.SQLite.EF6.dll
├── System.Data.SQLite.Linq.dll
├── x64/
│   └── SQLite.Interop.dll       # 64bit ネイティブDLL
├── x86/
│   └── SQLite.Interop.dll       # 32bit ネイティブDLL
└── data/                        # データフォルダ（初回起動時に自動作成）
    ├── archive.db               # SQLite データベース
    └── attachments/             # 添付ファイル保存先
```

---

## デプロイスクリプトによるデプロイ

プロジェクトルートの `deploy.ps1` を使用するのが最も簡単な方法です。

### デフォルトのデプロイ先（`C:\Tools\OutlookArchiver`）

```powershell
.\deploy.ps1
```

### デプロイ先を指定する場合

```powershell
.\deploy.ps1 -DeployDir "D:\MyApp\OutlookArchiver"
```

スクリプトは以下を自動で行います:

1. Release ビルドの実行
2. デプロイ先フォルダの作成
3. 実行ファイルと依存DLLのコピー
4. SQLite ネイティブDLL（x64/x86）のコピー

**注意:** データフォルダ（`data/`）はスクリプトではコピーしません。既存データの移行は手動で行ってください。

---

## 手動デプロイ

### 1. Release ビルド

```powershell
.\build.ps1 -Config Release
```

または MSBuild を直接実行:

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" OutlookArchiver\OutlookArchiver.vbproj /t:Restore /p:Configuration=Release
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" OutlookArchiver\OutlookArchiver.vbproj /p:Configuration=Release
```

ビルド成果物は `OutlookArchiver/bin/Release/` に出力されます。

### 2. ファイルのコピー

`OutlookArchiver/bin/Release/` から以下のファイルをデプロイ先にコピーします:

| ファイル | 必須 | 説明 |
|---------|:----:|------|
| `OutlookArchiver.exe` | Yes | 実行ファイル |
| `OutlookArchiver.exe.Config` | Yes | アプリケーション設定 |
| `EntityFramework.dll` | Yes | Entity Framework |
| `EntityFramework.SqlServer.dll` | Yes | Entity Framework SQL Server |
| `Microsoft.Office.Interop.Outlook.dll` | Yes | Outlook COM 相互運用 |
| `System.Data.SQLite.dll` | Yes | SQLite ADO.NET プロバイダ |
| `System.Data.SQLite.EF6.dll` | Yes | SQLite EF6 サポート |
| `System.Data.SQLite.Linq.dll` | Yes | SQLite LINQ サポート |
| `x64/SQLite.Interop.dll` | Yes | SQLite ネイティブ (64bit) |
| `x86/SQLite.Interop.dll` | Yes | SQLite ネイティブ (32bit) |
| `OutlookArchiver.pdb` | No | デバッグシンボル |
| `OutlookArchiver.xml` | No | XML ドキュメント |

---

## データの移行

開発環境や別の環境からデータを移行する場合、`data` フォルダをそのままコピーします。

### 移行手順

1. **アプリケーションを終了する**（移行元・移行先の両方）

2. **data フォルダをコピー**
   ```
   コピー元: <旧環境>\data\
   コピー先: C:\Tools\OutlookArchiver\data\
   ```

3. **コピー対象の確認**
   ```
   data\
   ├── archive.db           # メールデータベース（必須）
   ├── archive.db-wal       # WAL ファイル（存在する場合はコピー）
   ├── archive.db-shm       # 共有メモリファイル（存在する場合はコピー）
   └── attachments\          # 添付ファイル（必須）
       ├── 20260101_120000_abc12345\
       │   └── document.pdf
       └── ...
   ```

### 移行時の注意点

- `archive.db-wal` と `archive.db-shm` はアプリ稼働中に存在するファイルです。アプリを正常終了してからコピーしてください。
- 添付ファイルのパスはデータベース内に相対パスで保存されているため、`data` フォルダの構造を変えなければそのまま動作します。
- `data` フォルダの名前や配置を変更したい場合は、`OutlookArchiver.exe.Config` の `DbFilePath` と `AttachmentDirectory` を編集してください。

---

## デプロイ後の確認

1. `OutlookArchiver.exe` をダブルクリックで起動
2. データを移行した場合、メール一覧にデータが表示されることを確認
3. 添付ファイル付きメールを開き、添付ファイルが正しく表示されることを確認
4. 新規にメールを取り込み、正常に動作することを確認

---

## アップデート

アプリケーションをアップデートする場合は、`deploy.ps1` を再実行するだけで完了します。

```powershell
.\deploy.ps1
```

- 実行ファイルと依存DLLのみが上書きされます
- `data` フォルダ（データベース・添付ファイル）は影響を受けません

---

## トラブルシューティング

### 起動時にエラーが表示される

- .NET Framework 4.6.2 がインストールされているか確認してください
- すべての依存DLLがコピーされているか確認してください（特に `x64/SQLite.Interop.dll`）

### データが表示されない

- `data` フォルダが正しい場所にあるか確認してください
- `OutlookArchiver.exe.Config` のパス設定が正しいか確認してください

### 添付ファイルが開けない

- `data\attachments\` フォルダが正しくコピーされているか確認してください
- `attachments` フォルダ内のサブフォルダ構成が維持されているか確認してください
