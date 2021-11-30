const ORIGINAL_MENU_NAME = "勤怠管理メニュー";
const GET_CALENDAR_INFO_BUTTON_NAME = "勤怠情報の取得";
const TITLE_NAME = "勤怠管理";
const COMPLETE_PROCESS_TITLE = "勤務表作成完了";
const COMPLETE_PROCESS_MESSAGE = "お仕事お疲れ様でした。";
const ERROR_TITLE = "エラー発生";
const ERROR_MESSAGE_INVALID_SHEET_NAME = "無効なシート名です。";
const ERROR_MESSAGE_NO_DATA = "指定月の稼働は、現在ありません。";
const CALENDAR_ID = "ENTER_HERE@group.calendar.google.com";

/*
【コーディング・ルール】
- 関数：キャメルケース（例：updateCalendarInfo()）
- グローバル定数：コンスタントケース（例：ORIGINAL_MENU_NAME）
- その他の定数および変数：スネークケース（例：sheet_name）
*/

function onOpen() {
    // オリジナルメニューの追加
    // TIPS: スプレッドシートが開いたタイミングで実行されます。
    let ui = SpreadsheetApp.getUi();
    ui.createMenu(ORIGINAL_MENU_NAME)
        .addItem(GET_CALENDAR_INFO_BUTTON_NAME, "updateCalendarInfo")
        .addToUi();
}

function updateCalendarInfo() {
    try {
        // タイトルの記入（B2セルへ記入）
        const sheet_name = SpreadsheetApp.getActiveSpreadsheet()
            .getActiveSheet()
            .getName(); // 選択中（アクティブ）のシート名を取得
        const sheet_title = TITLE_NAME + "（" + sheet_name + "）"; // 月の情報をタイトルに付加
        SpreadsheetApp.getActiveSpreadsheet()
            .getActiveSheet()
            .getRange(2, 2)
            .setValue(sheet_title); // B2セルへタイトルを書き込み

        // カレンダーの取得
        const my_calendar = CalendarApp.getCalendarById(CALENDAR_ID);

        // 選択月の初日と末日を取得
        const calc_month_string =
            SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName() +
            "/01";
        const start_day = new Date(calc_month_string + " 00:00:00");
        const end_day = new Date(
            start_day.getFullYear(),
            start_day.getMonth() + 1,
            1
        );
        if (start_day == "Invalid Date") {
            // エラー処理：シートの名前が正しくない時
            throw ERROR_MESSAGE_INVALID_SHEET_NAME;
        }

        // 選択月のカレンダー内の情報（イベント）を取得
        const events = my_calendar.getEvents(start_day, end_day);
        if (events.length == 0) {
            // エラー処理：イベントが何もない時
            throw ERROR_MESSAGE_NO_DATA;
        }

        // カレンダー情報の変数初期化
        let record_list = []; // 予定情報の配列（List）
        let start_time; // 開始時間（Date）
        let end_time; // 終了時間（Date）
        let schedule_name = ""; // 予定の名前（String）
        let work_time = 0.0; // 勤務時間（Float）
        let mtg_time = 0.0; // MTG時間（Float）
        let work_time_total = 0.0; // 勤務合計時間（Float）
        let mtg_time_total = 0.0; // MTG合計時間（Float）

        for (const event of events) {
            start_time = event.getStartTime();
            end_time = event.getEndTime();
            schedule_name = event.getTitle();
            work_time = (end_time - start_time) / (60 * 60 * 1000);
            if (schedule_name.match(/MTG/)) {
                mtg_time = work_time;
            } else {
                mtg_time = 0;
            }

            const record = [
                start_time,
                end_time,
                schedule_name,
                work_time,
                mtg_time,
            ];
            record_list.push(record);

            work_time_total = work_time_total + work_time;
            mtg_time_total = mtg_time_total + mtg_time;
        }

        // レコードを書き込み
        // TIPS: 一行ずつ書き込むと時間がかかるため、このように配列に格納して、最後にまとめて記入します。
        SpreadsheetApp.getActiveSheet()
            .getRange(4, 2, record_list.length, record_list[0].length)
            .setValues(record_list);

        // 合計時間（全勤務およびMTG）の書き込み
        SpreadsheetApp.getActiveSheet()
            .getRange(2, 5)
            .setValue(work_time_total); // B5セルへ勤務合計時間を書き込み
        SpreadsheetApp.getActiveSheet().getRange(2, 6).setValue(mtg_time_total); // B6セルへMTG合計時間を書き込み

        Browser.msgBox(
            COMPLETE_PROCESS_TITLE,
            COMPLETE_PROCESS_MESSAGE,
            Browser.Buttons.OK
        );
        return;
    } catch (e) {
        Browser.msgBox(ERROR_TITLE, e, Browser.Buttons.OK);
        return;
    }
}
