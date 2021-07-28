//--------------------------------------------------------------------------------
// 処理制御用フラグ
//--------------------------------------------------------------------------------
const IS_DEBUG = false;          // テストと本番の切り替えフラグ
const ADD_CALENDER = true;       // カレンダー追加機能の切り替えフラグ
const MOVE_ARCHIVE = true;       // 処理実施後のメールをアーカイブへ移動させる機能の切り替えフラグ


//--------------------------------------------------------------------------------
// エントリー関数
//--------------------------------------------------------------------------------
function AddReceiptDateToCalendar() {
  // 既読のメールも対象にした方が漏れが防げるため、既読メールも対象にする
  const query = "from:(Amazon.co.jp) ご注文の確認";
  const max_num = 20;
  
  //const query = "is:unread from:(Amazon.co.jp) ご注文の確認";
  //const max_num = 10;

  AddReceiptDateToCalendarRunner(query, max_num);
}


//--------------------------------------------------------------------------------
// メール件名についている注文番号からKindleの注文かを判別
// 注文番号の先頭に「D」がついている場合はKindle注文と判断
// - Amazonの注文の場合：123-4567890-1234567
// - Kindleの注文の場合：D23-4567890-1234567
// 戻り値
// - true :Kindleでの注文
// - false:Amazonでの注文
//--------------------------------------------------------------------------------
function JudgeAmazonKindleOrder(message) {
  // 事前処理
  // ID検索に不要な文字列（「Amazon.co.jp」）を削除
  let subject = message.getSubject();
  subject = subject.replace("Amazon.co.jp", "");
  
  // A~Z、数値、ハイフンを含む文字列を削除
  let pattern = /[A-Z\d\-]+/;
  let order_id = subject.match(pattern)[0];

  // 先頭に「D」がついているかを確認  
  if (order_id[0] === "D") {
    return true;
  }
  return false;
}


//--------------------------------------------------------------------------------
// メール本文から取得した日時文字列から、日付/時間を抽出する
//--------------------------------------------------------------------------------
function ExtractDateDayTime(date_str) {
  // 不要な文字の削除
  date_str = date_str.replace('-', '');

  // カンマで分割して日付のみ取得し、前後の空白文字を削除
  let extract_date = date_str.split(",")[1];
  extract_date = extract_date.trim();

  // 空白分割し、日付と時間を分ける
  // 分割できない場合は、日付のみ返し、時間はnullとする
  let extract_day = extract_date;
  let extract_time = null;
  let split_item = extract_date.split(" ");
  if (split_item.length >= 2) {
    extract_day = split_item[0];
    extract_time = split_item[1];
  }

  return [extract_day, extract_time];
}


//--------------------------------------------------------------------------------
// 文字列の日付情報をDate型に変換
// timeがnullの場合は、時間指定無しのDate変数を作成する
//--------------------------------------------------------------------------------
function ConvertTypeString2Date(day_str, time_str, year) {

  let day_split = day_str.split('/');
  let month = Number(day_split[0]) - 1;
  let day = Number(day_split[1]);

  let date = null;

  if (time_str != null) {
    let time_split = time_str.split(':');
    let hour = Number(time_split[0]);
    let min = Number(time_split[1]);
    date = new Date(year, month, day, hour, min);
  } else {
    date = new Date(year, month, day);
  }

  return date;
}


//--------------------------------------------------------------------------------
// メールから抽出した文章から、お届け期間を取得する
//--------------------------------------------------------------------------------
function ExtractDeliveryPeriodList(year, receipt_string) {
  // メールの内容を、改行ごとに分割
  let receipt_lines = receipt_string.split('\r\n');

  // 空行の削除および前後空白をトリミング
  receipt_lines = receipt_lines.filter(Boolean);
  receipt_lines = receipt_lines.map(function( line ) {
    return line.trim();
  });

  // 「お届け予定日時」、「お届け予定日」の文字を検索時、インデックスを取得
  const time_indexs = receipt_lines.flatMap((v, i) => (v === 'お届け予定日時：' ? i : []));
  const date_indexs = receipt_lines.flatMap((v, i) => (v === 'お届け予定日：' ? i : []));

  let period_list = [];

  // 「お届け予定日時」の日付/時間を取得
  time_indexs.forEach(function (index) {
    // 「お届け予定日時」の1,2行下の行から期間の開始と終了を取得
    let begin_day_time = ExtractDateDayTime(receipt_lines[index + 1]);
    let end_day_time = ExtractDateDayTime(receipt_lines[index + 2]);
    const info = {
      begin_date: ConvertTypeString2Date(begin_day_time[0], begin_day_time[1], year),
      end_date  : ConvertTypeString2Date(end_day_time[0], end_day_time[1], year)
    };
    period_list.push(info);
  });

  // 「お届け予定日」の日付を取得
  date_indexs.forEach(function (index) {
    // 「お届け予定日」の1行下の行から日時を取得
    let begin_day_time = ExtractDateDayTime(receipt_lines[index + 1]);
    const info = {
      begin_date: ConvertTypeString2Date(begin_day_time[0], begin_day_time[1], year),
      end_date  : null
    };
    period_list.push(info);
  });
  
  return period_list;
}


//--------------------------------------------------------------------------------
// メールの件名からカレンダーに追加する際に必要な情報を取得する
//--------------------------------------------------------------------------------
function ExtractReciptDataList(year, body) {
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

  let recipt_datas = [];

  // 注文の内容から商品のお届け情報を取得し、配列に格納
  receipt_result.forEach(function (receipt) {
    // 注文番号を取得
    let order_id = receipt.match(/(?<=注文番号： )[\S]*/)[0];
    // お届け期間を取得する
    let date_info_list = ExtractDeliveryPeriodList(year, receipt);
    
    // 取得した届け日情報を配列に格納
    date_info_list.forEach(function (date_info) {
      let info = {
        begin_date: date_info.begin_date,
        end_date: date_info.end_date,
        title: "Amazon荷物",
        detail: "https://www.amazon.co.jp/gp/your-account/order-details/ref=ppx_yo_dt_b_order_details_o01?ie=UTF8&orderID=" + order_id
      };
      recipt_datas.push(info);
    });
  });

  return recipt_datas;
}


//--------------------------------------------------------------------------------
// カレンダーに追加済みかを判定
//--------------------------------------------------------------------------------
function JudgeEnableCalender(recipt_info) {
  // イベントの詳細にあるURLから登録済みかを判定する
  let result = false;
  const calendar = CalendarApp.getDefaultCalendar();
  const events = calendar.getEventsForDay(recipt_info.begin_date);
  for (let i=0; i < events.length; ++i) {
    let event_desc = events[i].getDescription();
    if (event_desc === recipt_info.detail) {
      result = true;
      break;
    }
  }
  return result;
}

//--------------------------------------------------------------------------------
// カレンダーに予定を追加
//--------------------------------------------------------------------------------
function AddCalender(recipt_info) {

  const calendar = CalendarApp.getDefaultCalendar();
  const option = {
    description: recipt_info.detail
  }

  // 終日イベントの判定
  if (recipt_info.end_date == null) {
    calendar.createAllDayEvent(recipt_info.title, recipt_info.begin_date, option)
  }
  else {
    calendar.createEvent(recipt_info.title, recipt_info.begin_date, recipt_info.end_date, option);
  }
}


//--------------------------------------------------------------------------------
// メイン処理関数
//--------------------------------------------------------------------------------
function AddReceiptDateToCalendarRunner(query, max_num) {
  
  GmailApp.search(query, 0, max_num + 1).forEach(function(thread) {
    
    thread.getMessages().forEach(function (message) {
      // ----------------------------------------
      // 通常モード
      // ----------------------------------------
      if (!IS_DEBUG) {
        // Kindle注文の場合は除外
        if (JudgeAmazonKindleOrder(message)){
          // 既読にする
          message.markRead();
          return;
        }
        
        // カレンダー追加に必要な情報を抽出
        let year = message.getDate().getFullYear();
        let body = message.getPlainBody();

        let recipt_datas = ExtractReciptDataList(year, body);
        for (let i=0; i < recipt_datas.length; ++i) {
          // カレンダーに登録されていない注文を登録対象とする
          if (ADD_CALENDER && !JudgeEnableCalender(recipt_datas[i])) {
            // カレンダーに追加
            AddCalender(recipt_datas[i]);
            // 既読にする
            message.markRead();
            // メッセージをconsoleに表示
            let console_message = "============================================================\n";
            console_message += "Add Amazon Recipt to Google Calender\n";
            console_message += ("Begin Date : " + recipt_datas[i].begin_date) + "\n";
            console_message += ("End Date   : " + recipt_datas[i].end_date) + "\n";
            console_message += ("Detail     : " + recipt_datas[i].detail) + "\n";
            console_message += "============================================================"
            console.log(console_message);
          }
        }
      }
      // ----------------------------------------
      // デバッグモード
      // ----------------------------------------
      else {
        // Kindle注文の場合は除外
        if (JudgeAmazonKindleOrder(message)){
          return;
        }
        let year = message.getDate().getFullYear();
        let recipt_datas = ExtractReciptDataList(year, message.getPlainBody());
        for (let i=0; i < recipt_datas.length; ++i) {
          let console_message = "============================================================\n";
          console_message += ("Calender Flag : " + JudgeEnableCalender(recipt_datas[i])) + "\n";
          console_message += ("Begin Date    : " + recipt_datas[i].begin_date) + "\n";
          console_message += ("End Date      : " + recipt_datas[i].end_date) + "\n";
          console_message += ("Detail        : " + recipt_datas[i].detail) + "\n";
          console_message += "============================================================";
          console.log(console_message);
        }
      }
    });

    // アーカイブに移動
    if (!IS_DEBUG && MOVE_ARCHIVE) {
      thread.moveToArchive();
    }
  });
}
