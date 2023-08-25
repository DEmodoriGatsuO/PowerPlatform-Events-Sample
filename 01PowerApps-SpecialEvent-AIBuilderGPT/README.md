# AI Builder GPT
## Apps
 - AIBuilderGPK_1_0_0_0.zip
  - ライブデモで作成したアンマネージドソリューションです

## Event
このアプリは、2023年8月25日に実施した[Power Apps Special event！Azure OpenAI Chatアプリを作ろう！](https://gatsuo.connpass.com/event/291029/)の<br>
イベント内で作成したサンプルになります

![完成イメージ](https://github.com/DEmodoriGatsuO/PowerPlatform-Events-Sample/blob/main/01PowerApps-SpecialEvent-AIBuilderGPT/asset/app_image.png)

<br>
READMEでは作成手順も含めて解説します。
<br>

### 注意点
イベントでは、`GPTでテキストを作成する`機能を利用していますが、2023.08.25時点では、一般公開されていません。<br>
一般公開に関する情報は下記を参考にしてください。

 - [AI Builder で Azure OpenAI GPT 機能にアクセスする](https://learn.microsoft.com/ja-jp/power-platform-release-plan/2022wave2/ai-builder/preview-access-openai-gpt-capabilities-ai-builder?source=recommendations)
 - [テキスト生成モデルの概要 (プレビュー)](https://learn.microsoft.com/ja-jp/ai-builder/prebuilt-azure-openai?source=recommendations)
 - [Power Apps (プレビュー) でテキスト生成モデルを使用する](https://learn.microsoft.com/ja-jp/ai-builder/azure-openai-model-papp)
 - [Power Automate (プレビュー) でテキスト世代モデルを使用する](https://learn.microsoft.com/ja-jp/ai-builder/azure-openai-model-pauto)

### 上記機能を一般公開前にプレビューで利用するためには
Power Platform 管理センターで、`米国`の`開発者`環境を作成し、AI BuilderのプレビューをONにする必要があります。

## アプリの機能
 - AI Builderの`GPTでテキストを作成する`機能を利用した、Text Chat
 - AI Builderの機能は、Azure OpenAIがもとになっている
 - クリアしたチャット履歴を呼び戻す機能

## アプリに必要なコネクタ
- Office 365 User
<br>

## 利用に必要な条件
- AI Builderのクレジットがないと利用できません
<br>

# 作成手順
## カラー（サンプル）
|コントロール|Hex|RGBA
|---|---|---
|Header|1072BC|RGBA(16, 114, 188, 1)
|Sidebar|000000|RGBA(0, 0, 0, 1)
|Title|FA9B70|RGBA(250, 155, 112, 1)
|Main|343541|RGBA(52, 53, 65, 1)

## 大枠
1. 水平コンテナーというコントロールを使い、画面の横幅を 2 : 8に分割する
    - 2の部分がサイドバー、8の部分がメインセクション
1. サイドバー(2)の部分に、垂直コンテナーを挿入する
    - メインセクション(8)の部分に、垂直コンテナーを挿入する
3.  このメインセクションに、3つ水平コンテナーを挿入する。大きさは高さを8等分する
    - 1つ目（最上部）のコンテナーはヘッダー、1/8の高さ
    - 2つ目（真ん中）のコンテナーはチャット画面、6/8の高さ
    - 3つ目（最下部）のコンテナーはフッター、1/8の高さ
4. メインセクションをアレンジする
5. サイドバーをアレンジする

### 1. 画面いっぱいに広がる水平コンテナーを挿入する
1. X,Yを0に変更
1. HeightをApp.Height
1. WidthをApp.Width
1. 境界半径は0に
1. ドロップシャドウはNoneに
     - コンテナにデフォルトで入ってしまう境界半径とドロップシャドウはNoneに基本していきます

### 2. 1で配置した水平コンテナーに、2つの垂直コンテナーを入れます
#### サイドバー用の設定をしましょう（左側のコンテナ
1. Sidebar用の垂直コンテナー
    1. 幅(伸縮可能)をオフにします
    1. Widthを`App.Width * 0.2` or `Parent.Width * 0.2`に設定
    1. 塗りつぶしで黒にします（RGBA(0,0,0,0)）

### 2. 1で配置した水平コンテナーに、2つの垂直コンテナーを入れます
#### サイドバー用の設定をしましょう（左側のコンテナ
1. Sidebar用の垂直コンテナー
    1. 幅(伸縮可能)をオフにします
    1. Widthを`App.Width * 0.2` or `Parent.Width * 0.2`に設定
    1. 塗りつぶしで黒にします（RGBA(0,0,0,0)）


### 3. 2で挿入した左側のコンテナー(メインセクション)に3つの水平コンテナーを挿入します
1. 最上部は`ヘッダー用``
    1. 高さ(伸縮可能)をオフにします
    1. Heightを`App.Height / 8` or `Parent.Height / 8`に設定
    1. 塗りつぶしを青にします
1. ヘッダーにテキストラベルを挿入
    1. HeightをParent.Heightに
    1. WidthをParent.Widthに
    1. Fontは好きなものに、サイズも好みで（例はLato 40）
    1. Colorは橙色に
1. 真ん中のコンテナ
    1. Heightを`App.Height / 8 * 6` or `Parent.Height / 8 * 6`
    1. 塗りつぶしを灰色に
1. 真ん中のコンテナに高さが伸縮可能なギャラリーを挿入

1. フッターとなるコンテナ
    1. 高さ(伸縮可能)をオフにします
    1. Heightを`App.Height / 8` or `Parent.Height / 8`に設定
    1. 塗りつぶしを灰色にします

1. テキストインプットとアイコンを挿入

1. サイドバーの設定
    1. パディングを40に
    1. ボタンを追加
    1. ゴミ箱アイコンを追加
    1. ギャラリーを追加

## データの追加
- GPT役の画像 - いらすとやの画像を利用
- Office 365 ユーザーコネクタ
- AI Builder

## 関数を挿入
### 送信アイコン
```yaml
If(TextInput1.Text <> "",
    UpdateContext({_prompt:TextInput1.Text});
    Collect(colChat,{_img:_photo,_text:TextInput1.Text});
    // AI Builderが使える場合
    // Power Apps
    Collect(colChat,{_img:ai_computer_sousa_robot,_text:'GPT でテキストを作成する'.Predict(_prompt).Text});
    //// Power Automate
    // Collect(colChat,{_img:ai_computer_sousa_robot,_text:GPTbyAIBuilder.Run(TextInput1.Text).result});
    Reset(TextInput1);
);
``` 

### + New Chatボタン
```yaml
Collect(colHistory,{col:colChat});
Clear(colChat);
``` 

### アプリのOnVisibleに
```yaml
UpdateContext({_photo:Office365ユーザー.UserPhotoV2(User().Email)});
```

## ラーニングパス
 - [AI Builder での GPT によるテキストの作成](https://learn.microsoft.com/ja-jp/training/modules/ai-builder-text-generation/)
