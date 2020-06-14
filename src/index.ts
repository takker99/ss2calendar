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
): Required<Record>[] {
    // 1. titleが空
    // 2. calendarIdが空
    // 3. 開始時刻が空
    // 4. 終了時刻が空
    // なrecordは除外する
    return rawRecords
        .filter(
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            (rawRecord: any[]): boolean =>
                rawRecord[setting.record.read.title - 1] != '' &&
                rawRecord[setting.record.read.calendarId - 1] != '' &&
                rawRecord[setting.record.write.start.time - 1] != '' &&
                rawRecord[setting.record.write.end.time - 1] != ''
        )
        .map(
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            (rawRecord: any[], index: number): Required<Record> => {
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
    console.log(`title: ${record.event.title}`);
    console.log(`start: ${record.event.start}`);
    console.log(`end: ${record.event.end}`);
    console.log(`description: ${record.event.description}`);
    console.log(`row index: ${record.row}`);

    // 1. task nameが空白
    // 2. calendar Idが空白
    // 3. 開始時刻or終了時刻が不正
    // のとき読み飛ばす
    if (record.event.title == '') {
        console.log('The title is empty. Skip updating.');
        return;
    }
    if (record.calendarId == '') {
        console.log('The Calendar ID is empty. Skip updating.');
        return;
    }
    if (!record.event.start.isValid()) {
        console.log('The start time is invalied. Skip updating.');
        console.log(`start: ${record.event.start}`);
        return;
    }
    if (!record.event.end.isValid()) {
        console.log('The end time is invalid. Skip updating.');
        console.log(`end: ${record.event.end}`);
        return;
    }
    const metaData: { [key: string]: string } = {
        spreadSheetId: sheet.getParent().getId(),
        sheetName: sheet.getSheetName(),
        row: record.row.toString(),
    };

    const recordCalendar = new Calendar(record.calendarId);
    // 既に登録済みの記録であれば、更新する
    if (record.eventId != '') {
        console.log('Updating the event...');
        const eventId = recordCalendar.Modify(
            record.eventId,
            record.event,
            metaData
        );
        if (eventId != undefined) {
            sheet
                .getRange(record.row, setting.record.write.eventId)
                .setValue(eventId);
        }

        console.log(`done.`);
        return;
    }

    // event IDを新規登録する
    console.log('Registering a new event...');
    const eventId = recordCalendar.Add(record.event, metaData);
    sheet.getRange(record.row, setting.record.write.eventId).setValue(eventId);
    console.log(`done.`);
}

// Google Calendarから、前後一ヶ月以内で最後に更新したeventを取得する
// eslint-disable-next-line @typescript-eslint/no-unused-vars
function getLastUpdatedEvent(
    calendarId: string
): GoogleAppsScript.Calendar.CalendarEvent | undefined {
    const now = Moment.moment().zone('+09:00');
    const aMonthAgo = now.subtract(1, 'month').toDate();
    const aMonthAfter = now.add(2, 'month').toDate();
    console.log(
        `Getting events from ${CalendarApp.getCalendarById(
            calendarId
        ).getName()} between ${aMonthAgo} and ${aMonthAfter}...`
    );
    const events = CalendarApp.getCalendarById(calendarId).getEvents(
        aMonthAgo,
        aMonthAfter
    );
    // このようにDateの生成を介さないで直接引数に代入すると失敗する。
    // const events = CalendarApp.getCalendarById(calendarId).getEvents(now_.subtract(1, 'months').toDate(), now_.toDate());
    console.log(`Got ${events.length} events.`);
    return events.length > 1
        ? events.reduce((a, b) =>
              a.getLastUpdated() < b.getLastUpdated() ? b : a
          )
        : undefined;
}

// Google Apps Script形式のcalendar eventからRecordを生成する
function toRecord(
    event: GoogleAppsScript.Calendar.CalendarEvent,
    calendarId: string
): Required<Record> {
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
        row: Number(event.getTag('row')),
    };
}

// event をsheetに書き込む
function writeEvent(
    record: Record,
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    setting?: SettingInfo
): void {
    if (setting == undefined) {
        setting = SettingManager.load();
    }
    if (setting == undefined) {
        console.error('Failed to load settings');
        return;
    }

    if (record.row) {
        console.log('Start updating the sheet');
        console.log(`Target spread sheet: ${sheet.getParent().getName()}`);
        console.log(`Target sheet: ${sheet.getSheetName()}`);
        console.log(`Row index: ${record.row}`);
        const range = sheet.getRange(
            record.row,
            setting.record.columnFlont,
            1,
            writingAreaLength(setting)
        );

        const getHHMM = (time: moment.Moment): string => {
            const temp = time.toDate();
            return Utilities.formatDate(temp, 'JST', 'HH:mm');
            /* return `${('00' + time.hours().toString()).slice(-2)}:${( */
            /*     '00' + time.minutes().toString() */
            /* ).slice(-2)}`; */
        };

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

        // description は扱いが特殊。
        // 1. 既存セルに値が存在する場合：
        //    既存の値を使う
        // 2. 値がない場合
        //    remarks欄にeventのdescriptionを代入する
        const oldValue = range.getValues()[0];
        const oldDescription = [
            oldValue[setting.record.write.expectation],
            oldValue[setting.record.write.actualAction],
            oldValue[setting.record.write.emotionTag],
            oldValue[setting.record.write.remarks],
        ];
        if (oldDescription.length > 0) {
            temp[setting.record.write.expectation] = oldDescription[0];
            temp[setting.record.write.actualAction] = oldDescription[1];
            temp[setting.record.write.emotionTag] = oldDescription[2];
            temp[setting.record.write.remarks] = oldDescription[3];
        } else {
            temp[setting.record.write.remarks] = record.event.description;
        }

        console.log(`These data will be writen:\n${temp}`);
        range.setValues([temp]);
        console.log('done.');

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
        console.log(
            'Row index is unknown. Search event IDs for the row index.'
        );
        if (record.eventId != '') {
            // event IDから変更を適用する行を検索する
            const result = (sheet
                .getRange(
                    setting.record.firstLine,
                    setting.record.write.eventId,
                    sheet.getLastRow() - setting.record.firstLine + 1,
                    1
                )
                .getValues()[0] as string[]).findIndex(
                (value) => value == record.eventId
            );
            // 行が存在したらそこに書き込む
            if (result > 0) {
                console.log('Found the row index');
                record.row = result + setting.record.firstLine;
                return writeEvent(record, sheet);
            }
        }
        // 新しい行を追加
        console.log('Create new line and write this record');
        sheet.insertRowAfter(sheet.getLastRow());
        record.row = sheet.getLastRow() + 1;
        writeEvent(record, sheet);
    }
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function writeSpreadSheet(e: GoogleCalendarEventObject): void {
    const changedEvent = getLastUpdatedEvent(e.calendarId);

    // 繰り返しeventは対象外
    if (changedEvent == undefined || changedEvent.isRecurringEvent()) {
        console.log('No event in a last month need to be changed.');
        return;
    }

    // 更新したいeventが記録されているsheetを開く
    if (changedEvent.getTag('spreadSheetId') == undefined) {
        // sheetが指定されていないeventは、Calendarから作成したeventとみなす
        // どう対処したらいいか思いつかないので、Sheetには書き込まないでおく
        // - どのsheetに書き込めばいいかわからない
        console.log('Skip synchronizing.');
        return;
    }
    const sheet = SpreadsheetApp.openById(
        changedEvent.getTag('spreadSheetId')
    ).getSheetByName(changedEvent.getTag('sheetName'));

    if (sheet == null) {
        console.error("the target sheet doesn't exist.");
        return;
    }

    const changedRecord = toRecord(changedEvent, e.calendarId);
    // for debug
    console.log(
        `This event has been changed and is going to be writen at ${sheet.getSheetName()}:`
    );
    console.log(`Title: ${changedRecord.event.title}`);
    console.log(`Description: ${changedRecord.event.description}`);
    console.log(`Start with ${changedRecord.event.start.toISOString()}`);
    console.log(`End with ${changedRecord.event.end.toISOString()}`);
    console.log(`Event Id:  ${changedRecord.eventId}`);

    // sheetに書き込む
    writeEvent(changedRecord, sheet);

    //TODO: 対応するeventIdが存在しなかったら、
    // - 新規作成する
    //   if calendarがrecord,tus,tus-cv以外
    // - 作成をskipしたとlogに書き込む
    //   otherwise
}

// actionを実行した時間帯に応じて、action nameの色を変える
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
            .setFontColor('#FFFFFF')
            .setRanges([titles])
            .build(),
    ];
    sheet.setConditionalFormatRules(rules);
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
