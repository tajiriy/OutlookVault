# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## プロジェクト概要

OutlookVault は .NET Framework 4.6.2 / VB.NET の Windows Forms デスクトップアプリケーション。Outlook のメール保管操作を提供する。

## ビルド・実行コマンド

```bash
# ビルド（Debug）— 本体 + テストプロジェクト
./build.sh

# リリースビルド
./build.sh Release

# テスト実行
./test.sh
```

> `build.sh` を必ず使用すること。直接 `msbuild` コマンドは使わない。

## 技術スタック

- **言語**: VB.NET
- **フレームワーク**: .NET Framework 4.6.2
- **UI**: Windows Forms
- **テスト**: NUnit 3.14 / NUnit3TestAdapter 4.5
- **プロジェクト形式**: 旧形式 (.vbproj / MSBuild ToolsVersion 15.0)

## アーキテクチャ

- ソリューション (`OutlookVault.sln`) に `OutlookVault/`（本体）と `OutlookVault.Tests/`（テスト）を含む構成
- エントリポイント: `My.MyApplication` → `MainForm`
- `*.Designer.vb` は Windows Forms デザイナーとの共有ファイル（下記「Windows Forms Designer の扱い」参照）

## 開発の進め方

- コーディングを始める前に、必ず実装プランを提示してユーザーの承認を得る
- プランを提示したら、承認されるまでコーディングに着手しない
- プランには変更対象のファイル、追加するクラスやメソッド、処理の流れなどを含める

## ブランチ運用ルール

- 基本的に `main` ブランチで直接開発する（1人開発のため）
- 大規模な実験的変更など、必要に応じて `feature-<機能名>` ブランチを使う

## コミット前の確認事項

- コミット前に必ず `msbuild OutlookVault/OutlookVault.vbproj` でビルドが通ることを確認する
- ビルドエラーがある状態ではコミットしない

## Windows Forms Designer の扱い

Windows Forms の標準パターンに従い、Designer.vb とコードビハインドの役割を分離する。

### Designer.vb (`InitializeComponent()`) に書くもの
- コントロールのインスタンス生成（`New`）
- プロパティ設定（`Name`, `Text`, `Size`, `Location`, `Anchor`, `Dock` など）
- レイアウト・配置（`Controls.Add`, `SuspendLayout` / `ResumeLayout`）
- イベントハンドラの紐付け（`AddHandler` / `Handles` 句に対応する宣言）

### コードビハインド (.vb) に書くもの
- イベントハンドラの実装
- ビジネスロジック・データ処理
- 動的に生成・変更するコントロールの操作

### Claude Code が Designer.vb を編集する際のルール
- 既存の `InitializeComponent()` のコードパターン（インデント、記述順、コメント形式）に合わせる
- `SuspendLayout` / `ResumeLayout` の構造を壊さない
- フィールド宣言（`Friend WithEvents`）を Designer.vb のクラス末尾に追加する
- 編集後はビルドして壊れていないことを確認する

### 編集対象外
- `My Project/` 配下のデザイナー生成ファイル (`*.Designer.vb`) は自動生成のため手動編集しないこと

## バージョン管理

- セマンティックバージョニング (`MAJOR.MINOR.PATCH`) を採用
- バージョンの正は `My Project/AssemblyInfo.vb` の `AssemblyVersion` / `AssemblyFileVersion`
- **PATCH**: `build.sh` 実行時に自動で +1（ビルドごとにインクリメント）
- **MINOR / MAJOR**: ユーザーの指示で上げる。または機能追加・破壊的変更を検知した場合は Claude が提案する
- `AboutForm` が `AssemblyVersion` を読み取って表示する

## VB.NET コーディング規約

- すべての `.vb` ソースファイル（自動生成ファイルを除く）の先頭に以下を明示すること:
  ```vb
  Option Explicit On
  Option Strict On
  Option Infer Off
  ```
- これらはプロジェクト設定（`.vbproj`）でも有効だが、ファイル単体で意図が伝わるよう各ファイルにも記述する
- 自動生成ファイル（`My Project/*.Designer.vb` 等）は対象外
