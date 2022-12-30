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


class AmazonReceipt extends AmazonParam {
    constructor(beginDate, endDate, orderId) {
        super();
        this.beginDate = beginDate
        this.endDate = endDate
        this.detail = this.EVENT_DETAIL_FORMAT + orderId
        this.title = this.EVENT_TITLE
        this.color = this.EVENT_COLOR
    }
}


class AmazonMailConverter {

    /**
     * お届け日時が記載されている文字列から配送時間を取得する。
     * @param {String} dateString お届け日時が記載されている文字列
     * @return {Object} 日付(day)と時間(time)を格納した連想配列。(時間がない場合,timeはnullとなる)
     */
    DateString2DayAndTimeString(dateString) {
        // 不要な文字の削除 かつ カンマで分割して日付のみ取得し、前後の空白文字を削除
        const fixDateString = dateString.replace("-", "").split(",")[1].trim();

        const splitItems = fixDateString.split(" ");
        if (splitItems.length >= 2) {
            return {day: splitItems[0], time: splitItems[1]};
        }
        else {
            return {day: splitItems[0], time: null};
        }
    }

    /**
     * 注文の内容から、お届け期間ごとのデータ配列を作成する。
     * @param {String} year メールの受信年
     * @param {String} receiptString 注文の内容
     * @param {String} orderId 配送番号。
     * @return {Object} カレンダーに追加する情報をまとめた配列。
     */
    CreateDeliveryPeriodReceiptInfoList(year, receiptString, orderId) {
        // メールの内容を、改行ごとに分割
        let receiptLines = receiptString.split("\r\n").filter(Boolean);

        // 空行の削除および前後空白をトリミング
        receiptLines = receiptLines.filter(Boolean);
        receiptLines = receiptLines.map((line) => {
            return line.trim();
        });

        // 「お届け予定日時」、「お届け予定日」の文字を検索時、インデックスを取得
        let timeIndexs = receiptLines.flatMap((v, i) => (v === "お届け予定日時：" ? i : []));
        let dateIndexs = receiptLines.flatMap((v, i) => (v === "お届け予定日：" ? i : []));

        let periodReceiptInfos = [];

        // 「お届け予定日時」の日付/時間を取得
        periodReceiptInfos = periodReceiptInfos.concat(timeIndexs.map((index) => {
            // 「お届け予定日時」の1,2行下の行から期間の開始と終了を取得
            const beginDateDict = this.DateString2DayAndTimeString(receiptLines[index + 1]);
            const endDateDict = this.DateString2DayAndTimeString(receiptLines[index + 2]);
            const beginDate = string2Date(year + "/" + beginDateDict.day, beginDateDict.time);
            const endDate = string2Date(year + "/" + endDateDict.day, endDateDict.time);
            return new AmazonReceipt(beginDate, endDate, orderId);
        }, this));

        // 「お届け予定日」の日付を取得
        periodReceiptInfos = periodReceiptInfos.concat(dateIndexs.map((index) => {
            // 「お届け予定日」の1行下の行から日時を取得
            const beginDateDict = this.DateString2DayAndTimeString(receiptLines[index + 1]);
            const beginDate = string2Date(year + "/" + beginDateDict.day, beginDateDict.time);
            return new AmazonReceipt(beginDate, null, orderId);
        }, this));

        return periodReceiptInfos;
    }

    /**
     * メールの件名からカレンダーに追加する際に必要な情報を配列に変換する。
     * @param {String} year メールの受信年
     * @param {String} body メールの文章
     * @return {Object} カレンダーに追加する情報をまとめた配列。
     */
    ConvertMail2ReceiptInfo(year, body) {
        // Bodyを分割
        const header = body.match(/[\s\S]*(?=領収書\/購入明細書)/)[0];
        const content = body.match(/(?<=注文履歴：)[\s\S]*(?=Amazon.co.jp でのご注文について詳しくは、下記URLから、注文についてのヘルプページをご確認ください。)/)[0];

        // 注文番号を取得（重複番号は削除する）
        const idResult = header.match(/(?<=注文番号： )[\S]*/g).filter((value, index, self) => {
            return self.indexOf(value) === index;
        });

        // 注文が複数の場合は、注文ごとに分割して代入し直す
        let receiptMessages = [content];
        if (idResult.length >= 2) {
            let tmp = content.split(/={3,}/g).slice(1);
            receiptMessages = tmp.filter((e) => {return e.match(/\S/) != null;});
        }

        // 注文の内容から商品のお届け情報を取得し、配列に格納
        let receiptInfos = [];
        receiptMessages.forEach((receipt) => {
            // 注文番号を取得
            const orderId = receipt.match(/(?<=注文番号： )[\S]*/)[0];
            // お届け期間を取得する
            receiptInfos = receiptInfos.concat(this.CreateDeliveryPeriodReceiptInfoList(year, receipt, orderId));
        }, this);

        return receiptInfos;
    }
}


class AmazonProcess extends AmazonParam {

    /**
     * メール件名についている注文番号からKindleの注文かを判別する。
     * 注文番号の先頭に「D」がついている場合はKindle注文と判断する。
     * - Amazonの注文の場合：123-4567890-1234567
     * - Kindleの注文の場合：D23-4567890-1234567
     * @param {Object} message メール情報。
     * @return {Boolean} true :Kindleでの注文、false:Amazonでの注文。
    */
    IsKindleOrder(message) {
        // ID検索に不要な文字列（「Amazon.co.jp」）を削除
        let subject = message.getSubject().replace("Amazon.co.jp", "");

        // A~Z、数値、ハイフンを含む文字列を削除
        let pattern = /[A-Z\d\-]+/;
        let orderId = subject.match(pattern)[0];

        // 先頭に「D」がついているかを確認
        if (orderId[0] === "D"){
            return true;
        }
        return false;
    }

    /**
     * Googleカレンダーに登録済みのお届け情報かを確認する。
     * @param {String} receiptInfo カレンダーに追加する情報をまとめた連想配列。
     * @return {Boolean}} カレンダーに存在するか(true: 存在する、false: 存在しない)。
     */
    IsEnableCalender(receiptInfo) {
        // イベントの詳細にあるURLから登録済みかを判定する
        const calendar = CalendarApp.getDefaultCalendar();
        const events = calendar.getEventsForDay(receiptInfo.beginDate);

        for (let i=0; i < events.length; ++i) {
            let eventDesc = events[i].getDescription();
            let eventColor = events[i].getColor();

            // 色が異なる場合は、カレンダーの予定を削除する
            // カレンダーにあり色も一致する場合は、カレンダーの予定をそのままにしてスキップする
            if (eventDesc == receiptInfo.detail && eventColor == this.EVENT_COLOR) {
                return true;
            }
            // カレンダーにあるが色が異なる場合は、カレンダーの予定を削除して新たに予定を追加する
            else if (eventDesc == receiptInfo.detail && eventColor != this.EVENT_COLOR) {
                events[i].deleteEvent();
                return false;
            }
        }
        return false;
    }

    /**
     * GMailからAmazonのお届け情報を取得する。
     * @return {List}} お届け情報を格納した連想配列。
     */
    AcquireReceiptInfo() {
        const converter = new AmazonMailConverter();
        let receiptInfos = [];

        GmailApp.search(this.TARGET_QUERY, 0, this.N_SEARCH_MAILS + 1).forEach((thread) => {

            // カレンダー追加に必要な情報を抽出
            thread.getMessages().forEach((message) => {

                // Kindle注文の場合は除外
                if (this.IsKindleOrder(message)){
                    // 既読にする
                    message.markRead();
                    return;
                }

                // お届け情報を取得し、配列に格納
                const year = message.getDate().getFullYear();
                const body = message.getPlainBody();
                receiptInfos = receiptInfos.concat(converter.ConvertMail2ReceiptInfo(year, body));

                // 既読にする
                message.markRead();
            });

            // アーカイブに移動
            if (this.MOVE_ARCHIVE) {
                thread.moveToArchive();
            }
        });
        return receiptInfos;
    }

    /**
     * お届け情報をGoogleカレンダーに登録する。
     * @param {List} receiptInfos お届け情報
     */
    ResistReceiptInfos2Calendar(receiptInfos) {
        receiptInfos.forEach((receipt, index) => {
            // 既に登録されている注文であればスキップ
            if (this.IsEnableCalender(receipt)) return;

            // カレンダーに追加
            addGoogleCalender(receipt);

            // メッセージをconsoleに表示
            const consoleMessage = "◆ Amazon Receipt Number " + index + "\n" +
                                    "- Begin Date : " + receipt.beginDate + "\n" +
                                    "- End Date   : " + receipt.endDate + "\n" +
                                    "- Detail     : " + receipt.detail + "\n";
            console.log(consoleMessage);
        });
    }
}


function ResistAmazonReceiptToCalendar() {
    const amazon = new AmazonProcess();
    receiptInfos = amazon.AcquireReceiptInfo();
    amazon.ResistReceiptInfos2Calendar(receiptInfos);
}
