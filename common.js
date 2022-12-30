//--------------------------------------------------------------------------------
// 文字列の日付情報をDate型に変換
// timeがnullの場合は、時間指定無しのDate変数を作成する
//--------------------------------------------------------------------------------
function string2Date(dayString, timeString) {

    let daySplit = dayString.split('/');
    let year = Number(daySplit[0])
    let month = Number(daySplit[1]) - 1;   // 月は0を1月として扱うため、-1する必要がある
    let day = Number(daySplit[2]);

    let date = null;

    if (timeString != null) {
        let timeSplit = timeString.split(':');
        let hour = Number(timeSplit[0]);
        let min = Number(timeSplit[1]);
        date = new Date(year, month, day, hour, min);
    } else {
        date = new Date(year, month, day);
    }

    return date;
}


//--------------------------------------------------------------------------------
// カレンダーに予定を追加
//--------------------------------------------------------------------------------
function addGoogleCalender(addInfo) {

    const calendar = CalendarApp.getDefaultCalendar();
    const option = {
        description: addInfo.detail
    }

    // 終日イベントの判定
    if (addInfo.end_date == null) {
        const event = calendar.createAllDayEvent(addInfo.title, addInfo.beginDate, option);
        event.setColor(addInfo.color);
    }
    else {
        const event = calendar.createEvent(addInfo.title, addInfo.beginDate, addInfo.endDate, option);
        event.setColor(addInfo.color);
    }
}
