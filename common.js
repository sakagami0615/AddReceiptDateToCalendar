//--------------------------------------------------------------------------------
// 文字列の日付情報をDate型に変換
// timeがnullの場合は、時間指定無しのDate変数を作成する
//--------------------------------------------------------------------------------
function string2Date(day_str, time_str) {

    let day_split = day_str.split('/');
    let year = Number(day_split[0])
    let month = Number(day_split[1]) - 1;   // 月は0を1月として扱うため、-1する必要がある
    let day = Number(day_split[2]);

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
// カレンダーに予定を追加
//--------------------------------------------------------------------------------
function addGoogleCalender(add_info) {

    const calendar = CalendarApp.getDefaultCalendar();
    const option = {
        description: add_info.detail
    }

    // 終日イベントの判定
    if (add_info.end_date == null) {
        calendar.createAllDayEvent(add_info.title, add_info.begin_date, option)
    }
    else {
        calendar.createEvent(add_info.title, add_info.begin_date, add_info.end_date, option);
    }
}
