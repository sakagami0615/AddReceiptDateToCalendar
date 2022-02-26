//--------------------------------------------------------------------------------
// メール本文から取得した日時文字列から、日付/時間を抽出する
//--------------------------------------------------------------------------------
function noshDate2DayAndTimeLabel(date_str) {

    // 不要な文字の削除
    let fix_date_str = date_str.replace('指定なし', '').trim();

    // 日付文字変換
    fix_date_str = fix_date_str.replace('年', '/');
    fix_date_str = fix_date_str.replace('月', '/');
    fix_date_str = fix_date_str.replace('日', '');

    let extract_day_str = fix_date_str;
    let extract_time_label = null;

    // 空白分割し、日付と時間を分ける
    // 分割できない場合は、日付のみ返し、時間はnullとする
    let split_item = fix_date_str.split(" ");
    if (split_item.length >= 2) {
        extract_day_str = split_item[0];
        extract_time_label = split_item[1];
    }

    return [extract_day_str, extract_time_label];
}


//--------------------------------------------------------------------------------
// 時間のラベル情報を数値情報に変換する
//--------------------------------------------------------------------------------
function noshLabel2Time(label) {

    if (label === '午前中(12時まで)'){
        return ['8:00', '12:00'];
    }
    else if (label === '14時から16時'){
        return ['14:00', '16:00'];
    }
    else if (label === '16時から18時'){
        return ['16:00', '18:00'];
    }
    else if (label === '18時から20時'){
        return ['18:00', '20:00'];
    }
    else if (label === '19時から21時'){
        return ['19:00', '21:00'];
    }

    return null;
}

//--------------------------------------------------------------------------------
// メールの件名からカレンダーに追加する際に必要な情報を取得する
//--------------------------------------------------------------------------------
function noshExtractReceiptData(body) {
    // メールの内容を、改行ごとに分割
    let receipt_lines = body.split('\r\n');

    // 「お届け予定日時」、「お届け予定日」の文字を検索時、インデックスを取得
    const date_index = receipt_lines.flatMap((v, i) => (v === '<お届け日時>' ? i : []))[0];
    // 「お届け予定日時」の日付/時間を取得
    let day_time = noshDate2DayAndTimeLabel(receipt_lines[date_index + 1]);
    time_values = noshLabel2Time(day_time[1]);

    let date_info = {};
    if (time_values != null){
        date_info['begin_date'] = string2Date(day_time[0], time_values[0]);
        date_info['end_date'] = string2Date(day_time[0], time_values[1]);
    }
    else {
        date_info['begin_date'] = string2Date(day_time[0], null);
        date_info['end_date'] = null;
    }

    let receipt_data = {
        begin_date: date_info.begin_date,
        end_date: date_info.end_date,
        title: "Nosh宅配",
        detail: "https://nosh.jp/mypage/subscription/379139"
    };
    return receipt_data;
}


//--------------------------------------------------------------------------------
// カレンダーに追加済みかを判定
//--------------------------------------------------------------------------------
function noshJudgeEnableCalender(receipt_info) {
    // イベントの詳細にあるURLから登録済みかを判定する
    let result = false;
    const calendar = CalendarApp.getDefaultCalendar();
    const events = calendar.getEventsForDay(receipt_info.begin_date);
    for (let i=0; i < events.length; ++i) {
        let title = events[i].getTitle();
        if (title === 'Nosh宅配') {
            result = true;
            break;
        }
    }
    return result;
}


//--------------------------------------------------------------------------------
// エントリー関数
//--------------------------------------------------------------------------------
function noshAddReceiptDateToCalendar() {

    const TARGET_QUERY = "from:(nosh.jp) 【ご注文ありがとうございます】";      // 探索するメールのクエリ
    const SEARCH_NUM = 20;                                                // 検索回数
    const MOVE_ARCHIVE = true                                             // 処理実施後のメールをアーカイブへ移動させる機能の切り替えフラグ

    console.log("/// AddNoshReceiptDateToCalendar BEGIN ///\n");

    let search_count = 0;
    GmailApp.search(TARGET_QUERY, 0, SEARCH_NUM + 1).forEach(function(thread) {

        thread.getMessages().forEach(function (message) {

            // カレンダー追加に必要な情報を抽出
            let body = message.getPlainBody();
            let receipt_data = noshExtractReceiptData(body);

            // 既に登録されている注文であればスキップ
            if (noshJudgeEnableCalender(receipt_data)) return;

            // 届け日をカレンダーに追加
            addGoogleCalender(receipt_data);

            // 既読にする
            message.markRead();

            // メッセージをconsoleに表示
            let console_message = "◆Nosh Receipt Number " + (search_count + 1) + "\n";
            console_message += "  Begin Date : " + receipt_data.begin_date + "\n";
            console_message += "  End Date   : " + receipt_data.end_date + "\n";
            console_message += "  Detail     : " + receipt_data.detail + "\n";
            console.log(console_message);
            search_count += 1;
        });

        // アーカイブに移動
        if (MOVE_ARCHIVE) {
            thread.moveToArchive();
        }
    });

    console.log("/// AddNoshReceiptDateToCalendar END ///\n");
}
