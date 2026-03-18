# 設定ダイアログで NullReferenceException

## 日付
2026-03-18

## 再現方法
設定ダイアログを開く（メニュー → 設定）

## 症状
`System.NullReferenceException: 'オブジェクト参照がオブジェクト インスタンスに設定されていません。'`
`SettingsForm.LoadSettings()` の `cboImportOrder.SelectedIndex` でクラッシュ。

## 原因
`SettingsForm.Designer.vb` の `InitializeComponent()` 内で、`lblImportOrder` と `cboImportOrder` のプロパティ設定・Controls.Add・コントロール宣言は追加したが、冒頭の `New` によるインスタンス生成を追加し忘れていた。

```vb
' これが漏れていた
Me.lblImportOrder = New System.Windows.Forms.Label()
Me.cboImportOrder = New System.Windows.Forms.ComboBox()
```

## 対応方法
`InitializeComponent()` 冒頭のインスタンス生成ブロックに上記2行を追加。

## 所見
Designer.vb を手動編集する際は、以下の4箇所すべてに追加が必要：
1. `New` によるインスタンス生成（冒頭）
2. プロパティ設定（中盤）
3. 親コントロールへの `Controls.Add`
4. `Friend WithEvents` コントロール宣言（末尾）

1つでも漏れると NullReferenceException になる。
