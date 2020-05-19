import * as ss from './calendar';
import * as variables from './variables';

function _writeCalendar(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    recordCalendar: ss.Calendar
): void {
    // 二次元配列転置用lambda式
    // シートにあるデータから
    // - calendarと同期するevent dataの開始cellの位置
    // を取得
    const rowTemp = sheet.getRange(1, 2, 2, 1).getValues() as number[][];
    const schemes: { row: number; column: number } = {
        row: rowTemp[1][0],
        column: rowTemp[0][0],
    };
    console.log(`データ開始行: ${schemes.row}`);
    console.log(`データ開始列: ${schemes.column}`);

    if (isNaN(schemes.row)) {
        console.log(
            '記録用のデータがありません。データの列の位置がずれている可能性があります'
        );
        return undefined;
    }

    interface Record {
        event: ss.Event;
        id: string;
    }
    // sheetから記録を入手
    //  1. task name
    //  2. remark
    //  3. 開始時刻のyyyy
    //  4. 開始時刻のMM
    //  5. 開始時刻のdd
    //  6. 開始時刻のhh
    //  7. 開始時刻のmm
    //  8. 終了時刻のyyyy
    //  9. 終了時刻のMM
    // 10. 終了時刻のdd
    // 11. 終了時刻のhh
    // 12. 終了時刻のmm
    // 13. event ID
    const records: Record[] = sheet
        .getRange(schemes.row, schemes.column, sheet.getLastRow() - 1, 13)
        .getValues()
        .map((record) => {
            return {
                event: new ss.Event(
                    record[0],
                    new ss.TimeSpan(
                        ss.getDateFixed(
                            record[2],
                            record[3],
                            record[4],
                            record[5],
                            record[6]
                        ),
                        ss.getDateFixed(
                            record[7],
                            record[8],
                            record[9],
                            record[10],
                            record[11]
                        )
                    ),
                    record[1]
                ),
                id: record[12],
            };
        });

    // 書き込み
    for (let i = 0; i < records.length; i++) {
        const record = records[i];

        // task nameが空白のときは
        // 読み飛ばす
        if (record.event.title == '') continue;
        console.log(`setting the event '${record.event.title}'...`);
        console.log(`start: ${record.event.start}`);
        console.log(`end: ${record.event.end}`);
        console.log(`description: ${record.event.description}`);

        // 既に登録済みの記録であれば、更新する
        if (record.id != '') {
            recordCalendar.ModifyEvent(record.id, record.event);
            console.log(`done.`);
            continue;
        }

        // event IDを新規登録する
        const eventId: string = recordCalendar.Add(record.event);
        console.log(`event ID: ${eventId}`);
        sheet
            .getRange(schemes.row + i, schemes.column + 13 - 1)
            .setValue(eventId);
        console.log(`done.`);
    }
}

function writeCalendar(): void {
    const sheet = SpreadsheetApp.getActiveSheet();
    if (sheet.getRange(3, 2, 1, 1).getValue() == 0) {
        // calendar用のsheetでなければ何もしない
        return undefined;
    }
    // SpreadSheetからdateの予定を
    // 取得
    if (sheet == null) {
        console.log("the target sheet doesn't exist.");
        return undefined;
    }
    _writeCalendar(sheet, new ss.Calendar(variables.eventCarendarId));
}

function writeSchedule(): void {
    const sheet = SpreadsheetApp.getActiveSheet();
    // SpreadSheetからdateの予定を
    // 取得
    if (sheet == null) {
        console.log("the target sheet doesn't exist.");
        return undefined;
    }
    _writeCalendar(
        sheet,
        new ss.Calendar('kua4bd6695fov7jrl9cmfu3o7o@group.calendar.google.com')
    );
}
