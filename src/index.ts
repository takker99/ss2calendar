import { Calendar, Event, getDateFixed, TimeSpan } from './calendar';
import { OnEditEventObject, GoogleCalendarEventObject } from './EventObject';
import {
    SettingInfo,
    writingAreaLength,
    toTag,
    recordLength,
} from './settingInfo';
import { SettingManager } from './settingManager';
import * as moment from 'moment';
const Moment = { moment: moment }; // GAS対策 cf. https://qiita.com/awa2/items/d24df6abd5fd5e4ca3d9

interface Record {
    event: Event;
    row?: number;
    eventId: string;
    calendarId: string;
}

// 二次元配列からrecordsを取得する
function getRecords(
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    rawRecords: any[][],
    firstRow: number,
    setting: SettingInfo
): Record[] {
    return rawRecords.map(
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (rawRecord: any[], index: number): Record => {
            return {
                row: firstRow + index,
                event: new Event(
                    rawRecord[setting.record.read.title - 1],
                    new TimeSpan(
                        getDateFixed(
                            rawRecord[setting.record.read.start.year - 1],
                            rawRecord[setting.record.read.start.month - 1],
                            rawRecord[setting.record.read.start.date - 1],
                            rawRecord[setting.record.read.start.hour - 1],
                            rawRecord[setting.record.read.start.minute - 1]
                        ),
                        getDateFixed(
                            rawRecord[setting.record.read.end.year - 1],
                            rawRecord[setting.record.read.end.month - 1],
                            rawRecord[setting.record.read.end.date - 1],
                            rawRecord[setting.record.read.end.hour - 1],
                            rawRecord[setting.record.read.end.minute - 1]
                        )
                    ),
                    rawRecord[setting.record.read.description - 1]
                ),
                eventId: rawRecord[setting.record.read.eventId - 1],
                calendarId: rawRecord[setting.record.read.calendarId - 1],
            };
        }
    );
}

// eventを更新する
function updateEvent(
    record: Required<Record>,
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    setting: SettingInfo
): void {
    // 1. task nameが空白
    // 2. calendar Idが空白
    // のとき読み飛ばす
    if (record.event.title == '' || record.calendarId == '') {
        console.log('skip updating');
        return;
    }
    console.log(`setting the event '${record.event.title}'...`);
    console.log(`start: ${record.event.start}`);
    console.log(`end: ${record.event.end}`);
    console.log(`description: ${record.event.description}`);

    const recordCalendar = new Calendar(record.calendarId);
    // 既に登録済みの記録であれば、更新する
    if (record.eventId != '') {
        recordCalendar.Modify(record.eventId, record.event);
        console.log(`done.`);
        return;
    }

    // event IDを新規登録する
    const eventId = recordCalendar.Add(record.event);
    console.log(`event ID: ${eventId}`);
    sheet.getRange(record.row, setting.record.write.eventId).setValue(eventId);
    console.log(`done.`);
}

// Google Calendarから、直近一ヶ月以内で最後に更新したeventのIDを取得する
// eslint-disable-next-line @typescript-eslint/no-unused-vars
function getLastUpdatedEvent(
    calendarId: string
): GoogleAppsScript.Calendar.CalendarEvent {
    const now = Moment.moment().zone('+09:00');
    return CalendarApp.getCalendarById(calendarId)
        .getEvents(now.subtract(1, 'months').toDate(), now.toDate())
        .reduce((a, b) => (a.getLastUpdated() < b.getLastUpdated() ? b : a));
}

// Google Apps Script形式のcalendar eventからRecordを生成する
function toRecord(
    event: GoogleAppsScript.Calendar.CalendarEvent,
    calendarId: string
): Record {
    const startTime = Moment.moment({
        years: event.getStartTime().getFullYear(),
        months: event.getStartTime().getMonth(),
        day: event.getStartTime().getDate(),
        hours: event.getStartTime().getHours(),
        minutes: event.getStartTime().getMinutes(),
    });
    const endTime = Moment.moment({
        years: event.getEndTime().getFullYear(),
        months: event.getEndTime().getMonth(),
        day: event.getEndTime().getDate(),
        hours: event.getEndTime().getHours(),
        minutes: event.getEndTime().getMinutes(),
    });
    const timeSpan = new TimeSpan(startTime, endTime);
    return {
        event: new Event(event.getTitle(), timeSpan, event.getDescription()),
        eventId: event.getId(),
        calendarId: calendarId,
    };
}
function writeSpreadSheet(e: GoogleCalendarEventObject): void {
    const changedEvent = getLastUpdatedEvent(e.calendarId);
    // 繰り返しeventは対象外
    if (changedEvent.isRecurringEvent()) return;
    const changedRecord = toRecord(changedEvent, e.calendarId);

    // 現在のsheetを読み込む
    const sheet = SpreadsheetApp.getActiveSheet();
    if (sheet == null) {
        console.error("the target sheet doesn't exist.");
        return;
    }

    // sheetに書き込む
    writeEvent(changedRecord, sheet);

    //TODO: 対応するeventIdが存在しなかったら、
    // - 新規作成する
    //   if calendarがrecord,tus,tus-cv以外
    // - 作成をskipしたとlogに書き込む
    //   otherwise
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function updateConditionalFormat(): void {
    const sheet = SpreadsheetApp.getActiveSheet();
    if (sheet == null) {
        console.error("the target sheet doesn't exist.");
        return;
    }

    const settings = SettingManager.load();
    if (settings == undefined) return;
    console.log('Got the setting information');
    console.log(`settings: ${JSON.stringify(settings)}`);

    const titles = sheet.getRange(
        settings.record.firstLine,
        settings.record.write.title,
        sheet.getLastRow(),
        1
    );

    // 条件付き書式を再作成する
    //   各task終了時間に応じてtaskの色分けをする
    sheet.clearConditionalFormatRules();
    const rules = [
        SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(
                '=hour(timevalue($I4))*60+minute(timevalue($I4))<4*60'
            )
            .setBackground('#57bb8a')
            .setRanges([titles])
            .build(),
        SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(
                '=hour(timevalue($I4))*60+minute(timevalue($I4))<8*60'
            )
            .setBackground('#b7e1cd')
            .setRanges([titles])
            .build(),
        SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(
                '=hour(timevalue($I4))*60+minute(timevalue($I4))<12*60'
            )
            .setBackground('#ffd666')
            .setRanges([titles])
            .build(),
        SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(
                '=hour(timevalue($I4))*60+minute(timevalue($I4))<16*60'
            )
            .setBackground('#f7981d')
            .setRanges([titles])
            .build(),
        SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(
                '=hour(timevalue($I4))*60+minute(timevalue($I4))<20*60'
            )
            .setBackground('#e67c13')
            .setRanges([titles])
            .build(),
        SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(
                '=hour(timevalue($I4))*60+minute(timevalue($I4))<24*60'
            )
            .setBackground('#351c75')
            .setRanges([titles])
            .build(),
    ];
    sheet.setConditionalFormatRules(rules);
}

function getHHMM(time: moment.Moment): string {
    return `${('00' + time.hours().toString()).slice(-2)}:${(
        '00' + time.minutes().toString()
    ).slice(-2)}`;
}
// event をsheetに書き込む
function writeEvent(
    record: Record,
    sheet: GoogleAppsScript.Spreadsheet.Sheet
): void {
    const setting = SettingManager.load();
    if (setting == undefined) return;

    if (record.row) {
        const range = sheet.getRange(
            record.row,
            setting.record.columnFlont,
            1,
            writingAreaLength(setting)
        );
        const temp = new Array(writingAreaLength(setting));
        temp[setting.record.write.tag - 1] = toTag(setting, record.calendarId);
        temp[setting.record.write.start.year - 1] = record.event.start.year();
        temp[setting.record.write.start.month - 1] =
            record.event.start.month() + 1;
        temp[setting.record.write.start.date - 1] = record.event.start.date();
        temp[setting.record.write.start.time - 1] = getHHMM(record.event.start);
        temp[setting.record.write.end.year - 1] = record.event.end.year();
        temp[setting.record.write.end.month - 1] = record.event.end.month() + 1;
        temp[setting.record.write.end.date - 1] = record.event.end.date();
        temp[setting.record.write.end.time - 1] = getHHMM(record.event.end);
        temp[setting.record.write.title - 1] = record.event.title;
        temp[setting.record.write.eventId - 1] = record.eventId;
        range.setValues([temp]);

        // 入力規則を追加
        const rules: GoogleAppsScript.Spreadsheet.DataValidation[] = new Array(
            writingAreaLength(setting)
        );
        rules[
            setting.record.write.tag - 1
        ] = SpreadsheetApp.newDataValidation()
            .requireValueInRange(setting.tagList)
            .build();
        rules[
            setting.record.write.emotionTag - 1
        ] = SpreadsheetApp.newDataValidation()
            .requireValueInRange(setting.emotionList)
            .build();

        range.clearDataValidations();
        range.setDataValidations([rules]);
    } else {
        if (record.eventId != '') {
            // event IDから変更を適用する行を検索する
            const searchedEventRow =
                (sheet
                    .getRange(
                        setting.record.firstLine,
                        setting.record.write.eventId,
                        sheet.getLastRow() - setting.record.firstLine + 1,
                        1
                    )
                    .getValues()[0] as string[]).findIndex(
                    (value) => value == record.eventId
                ) + setting.record.firstLine;
            // 行が存在したらそこに書き込む
            if (searchedEventRow != undefined) {
                record.row = searchedEventRow;
                return writeEvent(record, sheet);
            }
        }
        // 新しい行を追加
        sheet.insertRowAfter(sheet.getLastRow());
        record.row = sheet.getLastRow() + 1;
        writeEvent(record, sheet);
    }
}

function _writeCalendar(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    changedRange: GoogleAppsScript.Spreadsheet.Range,
    setting: SettingInfo
): void {
    // 変更範囲にどのrecordも含まれていなければ何もしない
    if (
        setting.record.columnFlont > changedRange.getLastColumn() ||
        setting.record.columnEnd < changedRange.getColumn() ||
        setting.record.firstLine > changedRange.getLastRow() ||
        sheet.getLastRow() < changedRange.getRow()
    ) {
        console.log('No record is changed.');
        return;
    }

    // 変更されたrecordを含む範囲を取得する
    const fixedRowIndex = Math.max(
        setting.record.firstLine,
        changedRange.getRow()
    );
    const rawRecords = sheet
        .getRange(
            fixedRowIndex,
            setting.record.columnFlont,
            Math.min(sheet.getLastRow(), changedRange.getLastRow()) -
                fixedRowIndex +
                1,
            recordLength(setting)
        )
        .getValues();

    // sheetから記録を入手
    const records = getRecords(rawRecords, fixedRowIndex, setting);
    console.log(`${records.length} records is going to be updated`);

    // 書き込み
    for (const record of records) {
        updateEvent(record, sheet, setting);
    }
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function writeCalendar(e: OnEditEventObject): void {
    // 変更されたのが設定情報の場合
    /* if (e.source.getActiveSheet().getName() == 'settings') { */
    /*     return; */
    /* } */

    const sheet = SpreadsheetApp.getActiveSheet();
    if (sheet == null) {
        console.error("the target sheet doesn't exist.");
        return;
    }
    console.log(`the current sheet: ${e.source.getActiveSheet().getName()}`);

    const settings = SettingManager.load();
    if (settings == undefined) return;
    /* console.log('Got the setting information'); */
    /* console.log(`settings: ${JSON.stringify(settings)}`); */

    if (!settings.isSync) {
        // calendar用のsheetでなければ何もしない
        console.log('synchronization is not available.');
        return;
    }

    // SpreadSheetからdateの予定を
    // 取得
    _writeCalendar(sheet, e.range, settings);
}

// 新しい記録dataを追加する
// 終了時刻は開始時刻と同じにする
// eslint-disable-next-line @typescript-eslint/no-unused-vars
function addRecord(): void {
    const sheet = SpreadsheetApp.getActiveSheet();
    if (sheet == null) {
        console.error("the target sheet doesn't exist.");
        return;
    }

    // 記録dataを作成する
    const now = Moment.moment().zone('+09:00');
    console.log(`現在時刻: ${now}`);
    const event = new Event('', new TimeSpan(now, now), '');
    const newRecord: Record = {
        row: 0,
        event: event,
        eventId: '',
        calendarId: '',
    };

    // 記録dataを書き込む
    writeEvent(newRecord, sheet);
}
