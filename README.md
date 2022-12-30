# **AddReceiptDateToCalendar**

## **1.ツール概要**

Google Apps Scriptを使用して、下記の宅配サービスのお届けメールを確認し、Googleカレンダーに予定を追加するソースコードです。
+ Amazon
+ Nosh

Amazonで説明すると、概要は下記図のとおりです。<br>
※Noshも同じ感じです

![tool-img](https://github.com/sakagami0615/AddReceiptDateToCalendar/raw/image/img/tool.png)
※メールとカレンダーの日付がずれていますが、気にしないでください…

<br>

## **2.前提条件**

本ツールを使用する際は下記の条件をクリアする必要があります。

+ Amazon,Noshのアカウントに登録しているメールがGMailであること。
+ 注文時に届くメールを迷惑メールに設定していたり、受信トレイをスキップしていないこと。
+ Googleカレンダーを使用していること。

<br>

## **3.準備**

+ Googleアカウントにログインしていない場合は、ログインする。
+ Google Apps Scriptのページにアクセスする。（検索すれば出てきます）
+ 「新しいプロジェクト」を作成する。（名前は好きなものに変えてもらってかまいません）
+ 「コード.gs」に記載されているコードをいったん削除する。
+ 削除した後、本リポジトリにある「XXXXX.js」に対応するgsファイルを作成し、jsのコードをコピペする。
+ [main.gs]を開き、デバックボタンの隣にあるプルダウン「Runner」を選択する。
+ 実行をクリックすると、承認の要求が来るはずなので、許可する。<br>
そのあとにGMailとGoogleカレンダーに関する要求も来るはずなので、許可する。
+ トリガー（時計のマーク）から「トリガーを追加」をクリック。
+ 下記のように設定し、保存する。
    + 実行する関数：Runner
    + イベントのソース：時間主導型
    + 時間ベースのトリガータイプ：時間ベースのタイマー
    + 時間の間隔：2時間おき（ここの設定は任意で変えても問題ない）
    + ※その他はデフォルトでOK

時間の間隔で設定した時間おきにGoogle Apps Scriptが作動します。<br>
※つまり、このツールはメールの受信をトリガーにしているわけではないのです…<br>

ちなみに一回の処理で20件のメールを捌けるので、1時間内に20件以上の注文をする場合があれば、もっと短い時間に設定するのが良いと思います。

<br>

## **4.補足**

### **パラメータに関して**

パラメータは、`amazon.js`と`nosh.js`の上部にクラスとして作成しています。<br>


```javascript
class AmazonParam {
    constructor() {
        Object.defineProperty(this, "TARGET_QUERY", {value: "from:(Amazon.co.jp) ご注文の確認"});
        Object.defineProperty(this, "N_SEARCH_MAILS", {value: 20});
        Object.defineProperty(this, "MOVE_ARCHIVE", {value: true});
        Object.defineProperty(this, "EVENT_TITLE", {value: "Amazon荷物"});
        Object.defineProperty(this, "EVENT_COLOR", {value: CalendarApp.EventColor.PALE_GREEN});
        Object.defineProperty(this, "EVENT_DETAIL_FORMAT", {value: "https://www.amazon.co.jp/gp/your-account/order-details/ref=ppx_yo_dt_b_order_details_o01?ie=UTF8&orderID="});
    }
}
```

+ `MOVE_ARCHIVE`<br>
  確認したメールをアーカイブに移動させるかどうかのフラグ(True: アーカイブ化する)
+ `EVENT_TITLE`<br>
  Googleカレンダーに追加する予定のタイトル。
+ `EVENT_COLOR`<br>
  Googleカレンダーに追加する予定の色。
+ `EVENT_DETAIL` or `EVENT_DETAIL_FORMAT`<br>
  Googleカレンダーに追加する予定の詳細。


> [補足] <br>
> NoshParamクラスの`EVENT_DETAIL`はユーザ毎に紐づくURLっぽいので、自身のURLに変えておくと良いかもです。
