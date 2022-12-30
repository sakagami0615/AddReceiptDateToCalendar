class NoshParam {
    constructor() {
        Object.defineProperty(this, "TARGET_QUERY", {value: "from:(nosh.jp) 【ご注文ありがとうございます】"});
        Object.defineProperty(this, "N_SEARCH_MAILS", {value: 20});
        Object.defineProperty(this, "MOVE_ARCHIVE", {value: true});
        Object.defineProperty(this, "EVENT_TITLE", {value: "Nosh宅配"});
        Object.defineProperty(this, "EVENT_COLOR", {value: CalendarApp.EventColor.ORANGE});
        Object.defineProperty(this, "EVENT_DETAIL", {value: "https://nosh.jp/mypage/subscription/379139"});
    }
}


class NoshReceipt extends NoshParam {
    constructor(beginDate, endDate) {
        super();
        this.beginDate = beginDate
        this.endDate = endDate
        this.detail = this.EVENT_DETAIL
        this.title = this.EVENT_TITLE
        this.color = this.EVENT_COLOR
    }
}


class NoshMailConverter {

    /**
     * 配送時間のラベル文字を配送時間の開始〜終了時間に変換する。
     * @param {String} label 配送時間のラベル文字列。
     * @return {Array} [開始時間, 終了時間]の文字配列。(対象外の入力の場合はnullを返す)
     */
    DeliveryLabel2Time(label) {
        switch (label) {
            case "午前中(12時まで)":
                return ["8:00", "12:00"];
            case "14時から16時":
                return ["14:00", "16:00"];
            case "16時から18時":
                return ["16:00", "18:00"];
            case "18時から20時":
                return ["18:00", "20:00"];
            case "19時から21時":
                return ["19:00", "21:00"];
        }
        return null;
    }

    /**
     * お届け日時が記載されている文字列から配送時間を取得する。
     * @param {String} dateString お届け日時が記載されている文字列
     * @return {Object} 日付(day)と時間(time)を格納した連想配列。(時間がない場合,timeはnullとなる)
     */
    DateString2DayAndTimeString(dateString) {
        // 不要な文字と日付文字の削除
        const fixDateString = dateString.replace(/年|月/g, "/").replace(/日|指定なし/g, "").trim();

        const splitItems = fixDateString.split(" ");
        if (splitItems.length >= 2) {
            return {day: splitItems[0], time: this.DeliveryLabel2Time(splitItems[1])};
        }
        else {
            return {day: splitItems[0], time: null};
        }
    }

    /**
     * メールの件名からカレンダーに追加する際に必要な情報に変換する。
     * @param {String} body メールの文章
     * @return {Object} カレンダーに追加する情報をまとめたオブジェクト。
     */
    ConvertMail2ReceiptInfo(body) {
        // メールの内容を、改行ごとに分割
        const receiptLines = body.split("\r\n");

        // 「お届け予定日時」、「お届け予定日」の文字を検索し、インデックスを取得
        const dateIndex = receiptLines.flatMap((v, i) => (v === "<お届け日時>" ? i : []))[0];

        // 「お届け予定日時」の日付/時間を取得
        const dateDict = this.DateString2DayAndTimeString(receiptLines[dateIndex + 1]);

        // 配送日時の開始と終了時間を格納数する
        if (dateDict.time != null){
            const beginDate = string2Date(dateDict.day, dateDict.time[0]);
            const endDate = string2Date(dateDict.day, dateDict.time[1]);
            return new NoshReceipt(beginDate, endDate);
        }
        else {
            const beginDate = string2Date(dateDict.day, null);
            return new NoshReceipt(beginDate, null);
        }
    }

}


class NoshProcess extends NoshParam {

    /**
     * メールの件名からカレンダーに追加する際に必要な情報を取得する。
     * @param {String} receiptInfo カレンダーに追加する情報をまとめた連想配列。
     * @return {Boolean}} カレンダーに存在するか(true: 存在する、false: 存在しない)。
     */
    IsEnableCalender(receiptInfo) {
        // イベントの詳細にあるURLから登録済みかを判定する
        const calendar = CalendarApp.getDefaultCalendar();
        const events = calendar.getEventsForDay(receiptInfo.beginDate);

        for (let i=0; i < events.length; ++i) {
            let eventTitle = events[i].getTitle();
            let eventColor = events[i].getColor();

            // 色が異なる場合は、カレンダーの予定を削除する
            // カレンダーにあり色も一致する場合は、カレンダーの予定をそのままにしてスキップする
            if (eventTitle == this.CALENDAR_TITLE && eventColor == this.CALENDAR_COLOR) {
                return true;
            }
            // カレンダーにあるが色が異なる場合は、カレンダーの予定を削除して新たに予定を追加する
            if (eventTitle == this.CALENDAR_TITLE && eventColor != this.CALENDAR_COLOR) {
                events[i].deleteEvent();
                return false;
            }
        }
        return false;
    }

    /**
     * GMailからNoshのお届け情報を取得する。
     * @return {List}} お届け情報を格納した連想配列。
     */
    AcquireReceiptInfo() {
        const converter = new NoshMailConverter();
        let receiptInfos = [];

        GmailApp.search(this.TARGET_QUERY, 0, this.N_SEARCH_MAILS + 1).forEach((thread) => {

            // カレンダー追加に必要な情報を抽出
            thread.getMessages().forEach((message) => {

                let body = message.getPlainBody();
                receiptInfos.push(converter.ConvertMail2ReceiptInfo(body));

                // 既読にする
                message.markRead();
            }, this);

            // アーカイブに移動
            if (this.MOVE_ARCHIVE) {
                thread.moveToArchive();
            }
        }, this);
        return receiptInfos;
    }

    /**
     * お届け情報をGoogleカレンダーに登録する。
     * @param {List} receiptInfos お届け情報
     */
    ResistReceiptInfos2Calendar(receiptInfos) {
        receiptInfos.forEach((receipt, index) => {
            // 既に登録されている注文であればスキップ
            // 既に登録されている注文であればスキップ
            if (this.IsEnableCalender(receipt)) return;

            // 届け日をカレンダーに追加
            addGoogleCalender(receipt);

            // メッセージをconsoleに表示
            const consoleMessage = "◆ Nosh Receipt Number " + (index + 1) + "\n" +
                                    "- Begin Date : " + receipt.beginDate + "\n" +
                                    "- End Date   : " + receipt.endDate + "\n" +
                                    "- Detail     : " + receipt.detail + "\n";
            console.log(consoleMessage);
        });
    }
}


function ResistNoshReceiptToCalendar() {
    const nosh = new NoshProcess();
    receiptInfos = nosh.AcquireReceiptInfo();
    nosh.ResistReceiptInfos2Calendar(receiptInfos);
}
