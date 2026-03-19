# OutlookVault

Outlook のメールをローカルの SQLite データベースに取り込んで保管するデスクトップアプリケーションです。取り込んだメールは Outlook を起動しなくても閲覧・検索できます。

## 背景

Outlook メールが 90 日間で自動削除されるポリシーに対応するため、過去のメールを手軽に残し、あとから参照できる仕組みとして開発しました。

## 主な機能

- **メール取り込み** — Outlook COM API 経由で受信トレイ・送信済みアイテム等のメールを SQLite に保存
- **自動取り込み** — 一定間隔または指定時刻での定期取り込み
- **メール閲覧** — フォルダツリー・メール一覧・プレビューの 3 ペイン構成
- **会話ビュー** — スレッド単位でメールをまとめて表示
- **検索** — SQLite による高速なメール検索
- **タスクトレイ常駐** — バックグラウンドで自動取り込みを継続

## 動作要件

- Windows 10 以降
- .NET Framework 4.6.2
- Microsoft Outlook（取り込み時に起動している必要があります）

## 技術スタック

| 項目 | 技術 |
|------|------|
| 言語 | VB.NET |
| フレームワーク | .NET Framework 4.6.2 / Windows Forms |
| メール取得 | Outlook COM API |
| データベース | SQLite (System.Data.SQLite) |
| テスト | NUnit 3.14 |

## ビルド

```bash
# Debug ビルド（本体 + テストプロジェクト）
./build.sh

# Release ビルド
./build.sh Release
```

## テスト

```bash
./test.sh
```

## プロジェクト構成

```
OutlookVault/
├── OutlookVault/           # 本体
│   ├── Config/             # アプリケーション設定
│   ├── Controls/           # カスタムコントロール
│   ├── Data/               # データベース管理・リポジトリ
│   ├── Filters/            # メール検索フィルター
│   ├── Forms/              # 各種フォーム
│   ├── Models/             # データモデル
│   ├── Services/           # ビジネスロジック
│   └── MainForm.vb         # メインウィンドウ
├── OutlookVault.Tests/     # テストプロジェクト
└── docs/                   # ドキュメント
    └── user-manual.md      # ユーザーマニュアル
```

## ドキュメント

- [ユーザーマニュアル](docs/user-manual.md)
- [デプロイ手順](docs/deploy-manual.md)

## ライセンス

Private — 社内利用を想定した非公開プロジェクトです。
