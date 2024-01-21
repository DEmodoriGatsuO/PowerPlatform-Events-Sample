# はじめに
今回は[こちらの記事](https://qiita.com/DEmodoriGatsuO/items/c0030ead94a1bec867f1)の続きで、コンポーネントでメッセージボックスを作っていきます。

https://qiita.com/DEmodoriGatsuO/items/c0030ead94a1bec867f1

[VBAのMsgBox 関数](https://learn.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/msgbox-function)の機能を再現してみます。

## VBAのMsgBox 関数
まずは、[VBAのMsgBox 関数](https://learn.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/msgbox-function)をおさらいしましょう。
このような時は公式のlearnが一番ですね！

https://learn.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/msgbox-function

```
MsgBox (prompt, [ buttons, ] [ title, ] [ helpfile, context ])
```

`prompt` は前回の記事で紹介した通りですが、`buttons` そして、Msgboxの`戻り値`に焦点を当てて作っていきたいと思います。

:::note warn
[ helpfile, context ] は割愛します。
:::

### [buttons引数](https://learn.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/msgbox-function)

[公式](https://learn.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/msgbox-function#settings)を参照します。

https://learn.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/msgbox-function#settings

- **ボタンの種類**
- **メッセージボックスのアイコン**
- システム モーダル
- ヘルプ
- 文字の右揃えの設定

上記が設定で入り混じっているように感じます。（個人の感想です）

このうち、`ボタンの種類`と`メッセージボックスのアイコン`を下記のようにして再現します。

:::note warn
ほかのところはごめんなさい🙇
:::

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

```plaintext:MsgBoxButtons　- 定数に対するレコード
LookUp(
        Table(
        {
            Constant: "pwOKOnly",
            Buttons: [
                "OK"
            ],
            Description: "OKボタンのみ"
        },
        {
            Constant: "pwOKCancel",
            Buttons: [
                "OK",
                "キャンセル"
            ],
            Description: "OK、キャンセル"
        },
        {
            Constant: "pwAbortRetryIgnore",
            Buttons: [
                "中止",
                "再試行",
                "無視"
            ],
            Description:"中止、再試行、無視"
        },
        {
            Constant: "pwYesNoCancel",
            Buttons: [
                "はい",
                "いいえ",
                "キャンセル"
            ],
            Description: "はい、いいえ、キャンセル"
        },
        {
            Constant: "pwYesNo",
            Buttons: [
                "はい",
                "いいえ"
            ], Description: "はい、いいえ"
        },
        {
            Constant: "pwRetryCancel",
            Buttons: [
                "再試行",
                "キャンセル"
            ],
            Description: "再試行、キャンセル"
        }
    ),
    // 後述するキー値で取得
    Constant = Self.pwMsgBoxStyle
)
```

### アイコンの設定
関数に着眼すると、
- vbCritical
- vbQuestion
- vbExclamation
- vbInformation

四つの種類のアイコンが存在します。
![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/2734332/e6231289-4bd7-2786-4b03-6794d9ad041a.png)

今回は、値の部分を割愛し、文字列からアイコンを設定するようにします。
なお、画像を追加する意味も特にないと思っているので、Power Appsの組み込みのアイコンで対処します。

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
クリックしたボタンに応じて、戻り値を設定します。
こちらは`レコード型`で、テキストも定数も値も返す形にしましょう。

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

:::note
別にいらないな、と思う部分は、ガンガン削ってください。
:::

# コンポーネントをカスタマイズする
文字を出す前に、完成系の[YouTube](https://youtu.be/T7cuxfuCMBw)と、[Power Apps]()を載せておきます。

<iframe width="560" height="315" src="https://www.youtube.com/embed/T7cuxfuCMBw" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture" loading="lazy" allowfullscreen></iframe>

## 入出力
まずはコンポーネントに渡す`カスタム プロパティ`を下記のとおり、設定しました。

### ■ 入力
|Name|データ型|値の説明|
|---|---|---|
|MsgBoxTitle|テキスト|メッセージボックスのタイトル|
|MsgBoxPrompt|テキスト|メッセージボックスの文章|
|pwOKOnly|テキスト|`MsgBoxButtons`のキー値|
|MsgBoxButtons|レコード|メッセージボックスのボタンを設定|
|pwIconStyle|テキスト|`MsgBoxIconStyle`のキー値|
|MsgBoxIconStyle|レコード|アイコンの設定|

:::note
カスタム プロパティは`表示名`、`名前`が存在しますが、
作者は英語で、かつ内容も全て統一しています。
:::

### ■ 出力
|Name|データ型|値の説明|
|---|---|---|
|DialogVisible|ブール値|メッセージボックスの表示・非表示を制御|
|MsgBoxResult|レコード|クリックしたボタンから値を出力|

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/2734332/d6fe9e42-4ffa-8afa-dace-82300e32af2d.png)

## 1. アイコンの追加
[前回の記事](https://qiita.com/DEmodoriGatsuO/items/c0030ead94a1bec867f1)との差分はアイコンの存在です。

メッセージが表示される赤線の枠に、設定がされた場合、アイコンが出るように、
`Icon` コントロールを追加します。

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/2734332/b61bed08-5574-a781-1185-1edad72e3ae7.png)

便利なことに`Icon`は`プロパティ`で表示を制御できます。合わせて`Color`も渡し、表示を似せていきましょう。

:::note 
文字に書ききれない部分もあるので、動画で内容を確認してください。
:::

## 2. ボタン コントロール
コンテナにボタンを3つ設置し、`配列`の要素数でボタンの数、テキストをコントロールします。

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/2734332/3bb17b92-9fd5-7d12-8022-ff7e1cbae863.png)

:::note alert
現在の方法では課題が存在しています。
コンポーネントを表示する際に、配列がコンポーネントで評価されないうちに
メッセージボックスが評価されています。

対処方法を模索中です。
`ギャラリーコントロール`で動的に出せたほうが美しいように感じていますが、
現段階では断念しています。
:::

ボタンの`OnSelect`には

```plaintext:OnSelect
Set(Return,
    LookUp(
        Table(
            { Text:"OK", Enum:"pwOK", Value:1},
            { Text:"キャンセル", Enum:"pwCancel", Value:2},
            { Text:"中止", Enum:"pwAbort", Value:3},
            { Text:"再試行", Enum:"pwRetry", Value:4},
            { Text:"無視", Enum:"pwIgnore", Value:5},
            { Text:"はい", Enum:"pwYes", Value:6},
            { Text:"いいえ", Enum:"pwNo", Value:7}
        ),Text = Self.Text
    )
);
Set(MsgBoxVisible,!MsgBoxVisible);
```

出力値を制御しています。
こちらの`Set`関数で定義した値が、出力値として返されます。

## 既知の課題
 - コンポーネント内のコントロールに、入力値から狙った値が反映されない
     - テーブル、レコードに基づく、コントロールの設定が間に合わないケースが多くあります・・・
 - Notレスポンシブ対応
     - 横長の画面対応想定
     - Smallサイズには小さすぎる課題があります。

丸一日粘りましたが、本日はギブアップです・・・。
**改善策を模索してみたいと思います！**

# おわりに
使いまわしによって開発が加速できると、本当に素晴らしいですね！
Power Apps全然ﾜｶﾗﾝ・・・

**良いPower Lifeを！**
