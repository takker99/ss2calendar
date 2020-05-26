import { Calendar, Event, getDateFixed, TimeSpan } from './calendar';

function _writeCalendar(sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
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
        event: Event;
        id: string;
        calendarId: string;
    }
    // sheetから記録を入手
    // 一番最後の要素の値をenumの要素数計算に使用する
    const enum RecordDataIndex {
        Title = 0,
        Expectation,
        ActualAction,
        Reason,
        Measure,
        FirstStep,
        EmotionTag,
        Remarks,
        StartYear,
        StartMonth,
        StartDate,
        StartHour,
        StartMinute,
        EndYear,
        EndMonth,
        EndDate,
        EndHour,
        EndMinute,
        EventId,
        CalendarId,
    }

    const records: Record[] = sheet
        .getRange(
            schemes.row,
            schemes.column,
            sheet.getLastRow() - 1,
            RecordDataIndex.CalendarId + 1 // 読み込むrecordの総数
        )
        .getValues()
        .map((record) => {
            return {
                event: new Event(
                    record[RecordDataIndex.Title],
                    new TimeSpan(
                        getDateFixed(
                            record[RecordDataIndex.StartYear],
                            record[RecordDataIndex.StartMonth],
                            record[RecordDataIndex.StartDate],
                            record[RecordDataIndex.StartHour],
                            record[RecordDataIndex.StartMinute]
                        ),
                        getDateFixed(
                            record[RecordDataIndex.EndYear],
                            record[RecordDataIndex.EndMonth],
                            record[RecordDataIndex.EndDate],
                            record[RecordDataIndex.EndHour],
                            record[RecordDataIndex.EndMinute]
                        )
                    ),
                    (record[RecordDataIndex.Expectation] != ''
                        ? '# 作業予定内容\n\n' +
                          record[RecordDataIndex.Expectation] +
                          '\n\n'
                        : '') +
                        (record[RecordDataIndex.ActualAction] != ''
                            ? '# 実際の作業結果\n\n' +
                              record[RecordDataIndex.ActualAction] +
                              '\n\n'
                            : '') +
                        (record[RecordDataIndex.EmotionTag] != ''
                            ? '## 作業時の心情\n\n' +
                              record[RecordDataIndex.EmotionTag] +
                              '\n\n'
                            : '') +
                        (record[RecordDataIndex.Reason] != ''
                            ? '# 何故そうなったか\n\n' +
                              record[RecordDataIndex.Reason] +
                              '\n\n'
                            : '') +
                        (record[RecordDataIndex.Measure] != ''
                            ? '# 次回はどうする\n\n' +
                              record[RecordDataIndex.Measure] +
                              '\n\n'
                            : '') +
                        (record[RecordDataIndex.FirstStep] != ''
                            ? '# まず何をするか\n\n' +
                              record[RecordDataIndex.FirstStep] +
                              '\n\n'
                            : '') +
                        (record[RecordDataIndex.Remarks] != ''
                            ? '# 備考\n\n' + record[RecordDataIndex.Remarks]
                            : '')
                ),
                id: record[RecordDataIndex.EventId],
                calendarId: record[RecordDataIndex.CalendarId],
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

        const recordCalendar: Calendar = new Calendar(record.calendarId);
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
            .getRange(schemes.row + i, schemes.column + RecordDataIndex.EventId)
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
    _writeCalendar(sheet);
}

function writeSchedule(): void {
    const sheet = SpreadsheetApp.getActiveSheet();
    // SpreadSheetからdateの予定を
    // 取得
    if (sheet == null) {
        console.log("the target sheet doesn't exist.");
        return undefined;
    }
    _writeCalendar(sheet);
}
