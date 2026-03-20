# 添付ファイルがインラインと誤判定され表示されない

## 発生日
2026-03-20

## 再現方法
1. Gmail から添付ファイル付きメールを Outlook に送信する
2. OutlookVault で取り込む
3. メールプレビューで添付ファイルパネルが表示されない

## 原因
`OutlookService.SaveAttachments` で添付ファイルのインライン判定に ContentId の有無のみを使用していた。

```vb
model.IsInline = Not String.IsNullOrEmpty(contentId)
```

Gmail などのメールクライアントは、通常の添付ファイル（非インライン）にも ContentId を付与する。
そのため、すべての添付ファイルが `IsInline = True` と誤判定された。

### 影響
- `is_inline = 1` で DB に保存される
- `ImportService.ProcessMailItem` のインライン画像のみ判定で `has_attachments = 0` に更新される
- `EmailPreviewControl.LoadAttachments` で `If att.IsInline Then Continue For` によりスキップされる
- 結果として添付ファイルがプレビューに一切表示されない

## 対応方法
ContentId があるだけでなく、HTML 本文中に `cid:そのContentId` への参照が実際に存在するかを確認するように変更。

```vb
model.IsInline = Not String.IsNullOrEmpty(contentId) AndAlso
                 Not String.IsNullOrEmpty(bodyHtml) AndAlso
                 bodyHtml.IndexOf("cid:" & contentId, StringComparison.OrdinalIgnoreCase) >= 0
```

## 所見
- ContentId の存在だけではインライン判定は不十分
- 正確な判定には HTML 本文内の cid: 参照の有無を確認する必要がある
- 既にインポート済みのメールは再取り込みが必要（削除して再インポート）
