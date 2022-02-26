//--------------------------------------------------------------------------------
// メール件名についている注文番号からKindleの注文かを判別
// 注文番号の先頭に「D」がついている場合はKindle注文と判断
// - Amazonの注文の場合：123-4567890-1234567
// - Kindleの注文の場合：D23-4567890-1234567
// 戻り値
// - true :Kindleでの注文
// - false:Amazonでの注文
//--------------------------------------------------------------------------------
function amazonJudgeKindleOrder(message) {
    // 事前処理
    // ID検索に不要な文字列（「Amazon.co.jp」）を削除
    let subject = message.getSubject();
    subject = subject.replace("Amazon.co.jp", "");

    // A~Z、数値、ハイフンを含む文字列を削除
    let pattern = /[A-Z\d\-]+/;
    let order_id = subject.match(pattern)[0];

    // 先頭に「D」がついているかを確認
    if (order_id[0] === "D") return true;

    return false;
}


//--------------------------------------------------------------------------------
// メール本文から取得した日時文字列から、日付/時間を抽出する
//--------------------------------------------------------------------------------
function amazonDate2DayAndTime(date_str) {

    // 不要な文字の削除
    let fix_date_str = date_str.replace('-', '');

    // カンマで分割して日付のみ取得し、前後の空白文字を削除
    fix_date_str = fix_date_str.split(",")[1];
    fix_date_str = fix_date_str.trim();

    let extract_day_str = fix_date_str;
    let extract_time_str = null;

    // 空白分割し、日付と時間を分ける
    // 分割できない場合は、日付のみ返し、時間はnullとする
    let split_item = fix_date_str.split(" ");
    if (split_item.length >= 2) {
        extract_day_str = split_item[0];
        extract_time_str = split_item[1];
    }

    return [extract_day_str, extract_time_str];
}


//--------------------------------------------------------------------------------
// メールから抽出した文章から、お届け期間を取得する
//--------------------------------------------------------------------------------
function amazonExtractDeliveryPeriodList(year, receipt_string) {
    // メールの内容を、改行ごとに分割
    let receipt_lines = receipt_string.split('\r\n');

    // 空行の削除および前後空白をトリミング
    receipt_lines = receipt_lines.filter(Boolean);
    receipt_lines = receipt_lines.map(function( line ) {
        return line.trim();
    });

    // 「お届け予定日時」、「お届け予定日」の文字を検索時、インデックスを取得
    let time_indexs = receipt_lines.flatMap((v, i) => (v === 'お届け予定日時：' ? i : []));
    let date_indexs = receipt_lines.flatMap((v, i) => (v === 'お届け予定日：' ? i : []));

    let period_list = [];

    // 「お届け予定日時」の日付/時間を取得
    time_indexs.forEach(function (index) {
        // 「お届け予定日時」の1,2行下の行から期間の開始と終了を取得
        let begin_day_time = amazonDate2DayAndTime(receipt_lines[index + 1]);
        let end_day_time = amazonDate2DayAndTime(receipt_lines[index + 2]);
        const info = {
            begin_date: string2Date(year + '/' + begin_day_time[0], begin_day_time[1]),
            end_date  : string2Date(year + '/' + end_day_time[0], end_day_time[1])
        };
        period_list.push(info);
    });

    // 「お届け予定日」の日付を取得
    date_indexs.forEach(function (index) {
        // 「お届け予定日」の1行下の行から日時を取得
        let begin_day_time = amazonDate2DayAndTime(receipt_lines[index + 1]);
        let info = {
            begin_date: string2Date(year + '/' + begin_day_time[0], begin_day_time[1]),
            end_date  : null
        };
        period_list.push(info);
    });

    return period_list;
}


//--------------------------------------------------------------------------------
// メールの件名からカレンダーに追加する際に必要な情報を取得する
//--------------------------------------------------------------------------------
function amazonExtractReceiptDataList(year, body) {
    // Bodyを分割
    let header = body.match(/[\s\S]*(?=領収書\/購入明細書)/)[0];
    let content = body.match(/(?<=注文履歴：)[\s\S]*(?=Amazon.co.jp でのご注文について詳しくは、下記URLから、注文についてのヘルプページをご確認ください。)/)[0];

    // 注文番号を取得（重複番号は削除する）
    let id_result = header.match(/(?<=注文番号： )[\S]*/g);
    id_result = id_result.filter(function(value, index, self){ return self.indexOf(value) === index;});

    // 注文が複数の場合は、注文ごとに分割して代入し直す
    let receipt_result = [content];
    if (id_result.length >= 2) {
        // 注文ごとに要素を分割
        // 分割時、最初の要素と空白文字列の要素は削除する
        receipt_result = content.split(/={3,}/g).slice(1);
        receipt_result = receipt_result.filter(function(e){return e.match(/\S/) != null;});
    }

    let receipt_datas = [];

    // 注文の内容から商品のお届け情報を取得し、配列に格納
    receipt_result.forEach(function (receipt) {
        // 注文番号を取得
        let order_id = receipt.match(/(?<=注文番号： )[\S]*/)[0];
        // お届け期間を取得する
        let date_info_list = amazonExtractDeliveryPeriodList(year, receipt);

        // 取得した届け日情報を配列に格納
        date_info_list.forEach(function (date_info) {
            let info = {
                begin_date: date_info.begin_date,
                end_date: date_info.end_date,
                title: "Amazon荷物",
                detail: "https://www.amazon.co.jp/gp/your-account/order-details/ref=ppx_yo_dt_b_order_details_o01?ie=UTF8&orderID=" + order_id
            };
            receipt_datas.push(info);
        });
    });

    return receipt_datas;
}


//--------------------------------------------------------------------------------
// カレンダーに追加済みかを判定
//--------------------------------------------------------------------------------
function amazonJudgeEnableCalender(receipt_info) {
    // イベントの詳細にあるURLから登録済みかを判定する
    let result = false;
    const calendar = CalendarApp.getDefaultCalendar();
    const events = calendar.getEventsForDay(receipt_info.begin_date);
    for (let i=0; i < events.length; ++i) {
        let event_desc = events[i].getDescription();
        if (event_desc === receipt_info.detail) {
            result = true;
            break;
        }
    }
    return result;
}


//--------------------------------------------------------------------------------
// エントリー関数
//--------------------------------------------------------------------------------
function amazonAddReceiptDateToCalendar() {

    // 既読のメールも対象にした方が漏れが防げるため、既読メールも対象にする
    const TARGET_QUERY = "from:(Amazon.co.jp) ご注文の確認";                // 探索するメールのクエリ
    //const TARGET_QUERY = "is:unread from:(Amazon.co.jp) ご注文の確認";    // 探索するメールのクエリ(未読のみ)
    const SEARCH_NUM = 20;                                                // 検索回数
    const MOVE_ARCHIVE = true                                             // 処理実施後のメールをアーカイブへ移動させる機能の切り替えフラグ

    console.log("/// AddAmazonReceiptDateToCalendar BEGIN ///\n");

    let search_count = 0;
    GmailApp.search(TARGET_QUERY, 0, SEARCH_NUM + 1).forEach(function(thread) {

        thread.getMessages().forEach(function (message) {

            // Kindle注文の場合は除外
            if (amazonJudgeKindleOrder(message)){
                // 既読にする
                message.markRead();
                return;
            }

            // カレンダー追加に必要な情報を抽出
            let year = message.getDate().getFullYear();
            let body = message.getPlainBody();
            let receipt_datas = amazonExtractReceiptDataList(year, body);

            // 届け日をカレンダーに追加
            for (let i=0; i < receipt_datas.length; ++i) {

                // 既に登録されている注文であればスキップ
                if (amazonJudgeEnableCalender(receipt_datas[i])) continue;

                // カレンダーに追加
                addGoogleCalender(receipt_datas[i]);
                // 既読にする
                message.markRead();
                // メッセージをconsoleに表示
                let console_message = "◆Amazon Receipt Number " + (search_count + 1) + "\n";
                console_message += "  Begin Date : " + receipt_datas[i].begin_date + "\n";
                console_message += "  End Date   : " + receipt_datas[i].end_date + "\n";
                console_message += "  Detail     : " + receipt_datas[i].detail + "\n";
                console.log(console_message);
                search_count += 1;
            }
        });

        // アーカイブに移動
        if (MOVE_ARCHIVE) {
            thread.moveToArchive();
        }
    });

    console.log("/// AddAmazonReceiptDateToCalendar END ///\n");
}
