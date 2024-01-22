# Qiita用のリポジトリ
https://qiita.com/DEmodoriGatsuO/items/c0030ead94a1bec867f1

## Goal
[VBAのMsgBox 関数](https://learn.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/msgbox-function)の機能を再現

## VBAのMsgBox 関数
- [VBAのMsgBox 関数](https://learn.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/msgbox-function)

https://learn.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/msgbox-function

```
MsgBox (prompt, [ buttons, ] [ title, ] [ helpfile, context ])
```
- [ helpfile, context ] は割愛

### [buttons引数](https://learn.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/msgbox-function)

https://learn.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/msgbox-function#settings

- **ボタンの種類**
- **メッセージボックスのアイコン**
- システム モーダル
- ヘルプ
- 文字の右揃えの設定

|再現方法|データ型|説明|
|---|---|---|
|ボタン|テーブル型|Lookupで配列を表示|
|アイコン|テーブル型|指定したアイコンをそのままコンポーネントに渡す、かつ色も指定|
|デフォルト|整数|列挙型が好きな場合は、カスタマイズしてください|


|定数|説明|
|---|---|
|pwOKOnly|OK のみ
|pwOKCancel|OK、キャンセル
|pwAbortRetryIgnore|中止、再試行、無視 
|pwYesNoCancel|はい、いいえ、キャンセル 
|pwYesNo|はい、いいえ 
|pwRetryCancel|再試行、キャンセル

### アイコンの設定
四つの種類のアイコンが存在します。
![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/2734332/e6231289-4bd7-2786-4b03-6794d9ad041a.png)

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/2734332/3de89d5a-eb91-c1d9-7734-b444138a2b2c.png)

|定数|値|説明|
|---|---|---|
|pwCritical|16|重大なメッセージ アイコン
|pwQuestion|32|警告クエリ アイコン
|pwExclamation|48|警告メッセージ アイコン
|pwInformation|64|情報メッセージ アイコン

```plaintext:MsgBoxButtons　- 定数に対するレコード
LookUp(
        Table(
            { pwIconStyle:"pwCritical", Icon:Icon.CancelBadge, Color:RGBA(255,0,0,1), Visible: true, Description:"重大なメッセージ"},
            { pwIconStyle:"pwQuestion", Icon:Icon.Help, Color:RGBA(0,134,208,1), Visible: true, Description:"警告クエリ"},
            { pwIconStyle:"pwExclamation", Icon:Icon.Warning, Color:RGBA(255,191,0,1), Visible: true, Description:"警告メッセージ"},
            { pwIconStyle:"pwInformation", Icon:Icon.Information, Color:RGBA(0,134,208,1), Visible: true, Description:"情報メッセージ"}
    ),
    // 後述するキー値で取得
    pwIconStyle = Self.pwIconStyle
)
```

### 戻り値
こちらは`レコード型`で、テキストも定数も値も返す

|テキスト|定数|値|
|---|---|---|
|OK|pwOK|1|
|キャンセル|pwCancel|2|
|中止|pwAbort|3|
|再試行|pwRetry|4|
|無視|pwIgnore|5|
|はい|pwYes|6|
|いいえ|pwNo|7|

```plaintext:値のテーブル
Table(
    { Text:"OK", Enum:"pwOK", Value:1},
    { Text:"キャンセル", Enum:"pwCancel", Value:2},
    { Text:"中止", Enum:"pwAbort", Value:3},
    { Text:"再試行", Enum:"pwRetry", Value:4},
    { Text:"無視", Enum:"pwIgnore", Value:5},
    { Text:"はい", Enum:"pwYes", Value:6},
    { Text:"いいえ", Enum:"pwNo", Value:7}
)
```

上記の中から`Lookup`でデータを抜き出して取得します。


# コンポーネント
https://youtu.be/T7cuxfuCMBw

## 入出力
まずはコンポーネントに渡す`カスタム プロパティ`を下記のとおり、設定しました。

### ■ 入力
|Name|データ型|値の説明|
|---|---|---|
|MsgBoxTitle|テキスト|メッセージボックスのタイトル|
|MsgBoxPrompt|テキスト|メッセージボックスの文章|
|pwOKOnly|テキスト|`MsgBoxButtons`のキー値|
|pwIconStyle|テキスト|`MsgBoxIconStyle`のキー値|

### ■ 出力
|Name|データ型|値の説明|
|---|---|---|
|DialogVisible|ブール値|メッセージボックスの表示・非表示を制御|
|MsgBoxResult|レコード|クリックしたボタンから値を出力|

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/2734332/d6fe9e42-4ffa-8afa-dace-82300e32af2d.png)

## Issue
 - Notレスポンシブ対応
     - 横長の画面対応想定
     - Smallサイズには小さすぎる課題があります。
