# AboutForm のアイコン画像が消える

## 発生日
2026-03-20

## 再現方法
1. FileHelper.GetAppIcon() で app.ico をキャッシュする共通メソッドを導入
2. AboutForm で従来通り `New Icon(icoPath, 256, 256)` でファイルから再読み込み
3. バージョン情報ダイアログを開くと picIcon にアイコン画像が表示されない

## 原因
`GetAppIcon()` が `New Drawing.Icon(icoPath)` でファイルを開き、Icon オブジェクトを静的フィールドにキャッシュしている。Icon クラスはファイルハンドルを保持し続けるため、同じファイルを別の Icon インスタンスで再オープンしようとすると失敗（または空の画像になる）。

## 対応方法
AboutForm ではファイルパスから読み直すのではなく、キャッシュ済みの Icon オブジェクトからサイズ指定で Bitmap を生成するように変更。

```vb
' 修正前（ファイルロックで失敗）
Dim ico As New Icon(icoPath, 256, 256)
picIcon.Image = ico.ToBitmap()

' 修正後（キャッシュから生成）
Dim largeIcon As New Icon(appIcon, 256, 256)
picIcon.Image = largeIcon.ToBitmap()
```

## 所見
Icon ファイルを共有キャッシュする場合、他の箇所でファイルパスから直接読み込むコードがないか確認が必要。Icon クラスのファイルロック特性に注意。
