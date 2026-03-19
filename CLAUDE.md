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
- `MainForm.Designer.vb` はデザイナー自動生成ファイル。手動編集しないこと

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

## 注意事項

- `*.Designer.vb` ファイルは Windows Forms デザイナーが管理するため、直接編集せず `InitializeComponent()` 内のコードはデザイナーに任せる
- UI コントロールの追加・変更は Designer.vb ではなくコードビハインド側（`MainForm.vb`）で対応するか、Designer.vb の `InitializeComponent()` 内に正しい形式で追記する
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
