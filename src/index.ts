import { Calendar, Event, TimeSpan, CalendarEventId } from './calendar';
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

export interface Record {
    event: Event;
    row?: number;
    id: CalendarEventId;
    isRecord: boolean; // 記録ならtrue
}

// 二次元配列からrecordsを取得する
function getRecords(
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    rawRecords: any[][],
    firstRow: number,
    setting: SettingInfo
): Required<Record>[] | undefined {
    // 記録用calendar idを取得する
    const recordCalendarId = PropertiesService.getScriptProperties().getProperty(
        'recordCalendarId'
    );
    if (recordCalendarId == null) {
        console.error(
            'The calendar ID for records could not be found in ScriptProperties.'
        );
        return undefined;
    }

    // 記録用eventと予定用eventを取得する
    const result: Required<Record>[] = [];

    // 値が正常なeventのみをresultに入れる函数
    // @return 成功: true, 値が不正: false
    const pushEvent = (
        title: string,
        start: number,
        end: number,
        description: string,
        id: CalendarEventId,
        isRecord: boolean,
        index: number
    ): boolean => {
        const getDateFixed = (unixSecond: number): moment.Moment =>
            Moment.moment.unix(unixSecond).subtract(9, 'hour');
        const timeSpan = new TimeSpan(getDateFixed(start), getDateFixed(end));
        // 1. task nameが空白
        // 2. 開始時刻or終了時刻が不正
        // のとき読み飛ばす
        if (title == '') {
            console.log('The title is empty. Skip updating.');
            return false;
        }
        if (!timeSpan.start.isValid()) {
            console.log('The start time is invalid. Skip updating.');
            console.log(`end: ${timeSpan.start}`);
            return false;
        }
        if (!timeSpan.end.isValid()) {
            console.log('The end time is invalid. Skip updating.');
            console.log(`end: ${timeSpan.end}`);
            return false;
        }

        const event = new Event(title, timeSpan, description);
        result.push({
            row: index + firstRow,
            event: event,
            id: id,
            isRecord: isRecord,
        });
        return true;
    };
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    rawRecords.forEach((rawRecord: any[], index: number) => {
        if (rawRecord[setting.record.read.schedule.canUpdate - 1]) {
            pushEvent(
                rawRecord[setting.record.read.schedule.title - 1] as string,
                rawRecord[setting.record.read.schedule.start - 1] as number,
                rawRecord[setting.record.read.schedule.end - 1] as number,
                rawRecord[
                    setting.record.read.schedule.description - 1
                ] as string,
                new CalendarEventId(
                    rawRecord[setting.record.read.schedule.calendarId - 1],
                    rawRecord[setting.record.read.schedule.eventId - 1]
                ),
                false,
                index
            );
        }
        if (rawRecord[setting.record.read.record.canUpdate - 1]) {
            pushEvent(
                rawRecord[setting.record.read.record.title - 1] as string,
                rawRecord[setting.record.read.record.start - 1] as number,
                rawRecord[setting.record.read.record.end - 1] as number,
                rawRecord[setting.record.read.record.description - 1] as string,
                new CalendarEventId(
                    recordCalendarId,
                    rawRecord[setting.record.read.record.eventId - 1]
                ),
                true,
                index
            );
        }
    });
    return result;
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

    const metaData: { [key: string]: string } = {
        spreadSheetId: sheet.getParent().getId(),
        sheetName: sheet.getSheetName(),
        isRecord: record.isRecord.toString(),
        row: record.row.toString(),
    };

    // Calendar に eventを登録する
    console.log('Updating the event...');
    const id = Calendar.Push(record.id, record.event, metaData);
    if (!id) return;
    // event Idを上書きする
    sheet
        .getRange(
            record.row,
            record.isRecord
                ? setting.record.write.record.eventId
                : setting.record.write.schedule.eventId
        )
        .setValue(id.eventId);

    console.log(`done.`);
    return;
}

// event をsheetに書き込む
export function writeEvent(
    record: Record,
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    setting?: SettingInfo
): void {
    if (setting == undefined) {
        setting = SettingManager.load();
    }
    if (setting == undefined) return;

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
        const duration = record.event.period.duration;
        const zero = (n: number): string => String(n).padStart(2, '0');

        const temp = new Array(writingAreaLength(setting));
        // 予定か記録かで書き込むセルが変わる
        if (record.isRecord) {
            temp[setting.record.write.record.title - 1] = record.event.title;
            temp[
                setting.record.write.record.start.year - 1
            ] = record.event.start.year();
            temp[setting.record.write.record.start.month - 1] =
                record.event.start.month() + 1;
            temp[
                setting.record.write.record.start.date - 1
            ] = record.event.start.date();
            temp[setting.record.write.record.start.time - 1] = getHHMM(
                record.event.start
            );
            temp[setting.record.write.record.duration - 1] = `${zero(
                duration.hours()
            )}:${zero(duration.minutes())}`;
            temp[setting.record.write.record.eventId - 1] = record.id.eventId;

            // description の更新はしない
        } else {
            temp[setting.record.write.schedule.tag - 1] = toTag(
                setting,
                record.id.calendarId
            );
            temp[
                setting.record.write.schedule.start.year - 1
            ] = record.event.start.year();
            temp[setting.record.write.schedule.start.month - 1] =
                record.event.start.month() + 1;
            temp[
                setting.record.write.schedule.start.date - 1
            ] = record.event.start.date();
            temp[setting.record.write.schedule.start.time - 1] = getHHMM(
                record.event.start
            );
            temp[setting.record.write.schedule.duration - 1] = `${zero(
                duration.hours()
            )}:${zero(duration.minutes())}`;
            temp[setting.record.write.schedule.title - 1] = record.event.title;
            temp[setting.record.write.schedule.description - 1] =
                record.event.description;
            temp[setting.record.write.schedule.eventId - 1] = record.id.eventId;
        }

        console.log(`These data will be writen:\n${temp}`);
        range.setValues([temp]);
        console.log('done.');

        // 入力規則を追加
        const rules: GoogleAppsScript.Spreadsheet.DataValidation[] = new Array(
            writingAreaLength(setting)
        );
        rules[
            setting.record.write.schedule.tag - 1
        ] = SpreadsheetApp.newDataValidation()
            .requireValueInRange(setting.tagList)
            .build();
        rules[
            setting.record.write.record.emotion - 1
        ] = SpreadsheetApp.newDataValidation()
            .requireValueInRange(setting.emotionList)
            .build();

        range.clearDataValidations();
        range.setDataValidations([rules]);
    } else {
        console.log('Row index is unknown.');
        if (record.id.eventId != '') {
            console.log('Search event IDs for the row index.');
            // event IDから変更を適用する行を検索する
            const result = (sheet
                .getRange(
                    setting.record.firstLine,
                    record.isRecord
                        ? setting.record.read.record.eventId
                        : setting.record.read.schedule.eventId,
                    sheet.getLastRow() - setting.record.firstLine + 1,
                    1
                )
                .getValues() as string[][]).findIndex(
                (value) => value[0] == record.id.eventId
            );
            // 行が存在したらそこに書き込む
            if (result > 0) {
                console.log('Found the row index');
                record.row = result + setting.record.firstLine;
                return writeEvent(record, sheet);
            }
            console.log('Not found.');
        }
        // 新しい行を追加
        console.log('Create new line and write this record');
        sheet.insertRowAfter(sheet.getLastRow());
        record.row = sheet.getLastRow() + 1;
        writeEvent(record, sheet);
    }
}

export function updateEventFromSheet(
    row: number,
    length: number,
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    setting: SettingInfo
): void {
    const rawRecords = sheet
        .getRange(
            row,
            setting.record.columnFlont,
            length,
            recordLength(setting)
        )
        .getValues();

    // sheetから記録を入手
    const records = getRecords(rawRecords, row, setting);
    if (!records) return;
    console.log(`${records.length} records is going to be updated`);

    // 書き込み
    for (const record of records) {
        updateEvent(record, sheet, setting);
    }
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function writeSpreadSheet(e: GoogleCalendarEventObject): void {
    const {
        // eslint-disable-next-line @typescript-eslint/no-unused-vars
        calendarExists,
        // eslint-disable-next-line @typescript-eslint/no-unused-vars
        eventExists,
        event,
        id,
        metaData,
    } = Calendar.GetLastUpdatedEvent(new CalendarEventId(e.calendarId), [
        'spreadSheetId',
        'sheetName',
    ]);

    if (!event || !id) {
        console.log('No event in a last month is changed.');
        return;
    }

    // 更新したいeventが記録されているsheetを開く
    if (!metaData) {
        // sheetが指定されていないeventは、Calendarから作成したeventとみなす
        // - どう対処したらいいか思いつかないので、Sheetには書き込まないでおく
        //   - どのsheetに書き込めばいいかわからない
        console.log('Skip synchronizing.');
        return;
    }
    const sheet = SpreadsheetApp.openById(
        metaData['spreadSheetId']
    ).getSheetByName(metaData['sheetName']);

    if (!sheet) {
        console.error("the target sheet doesn't exist.");
        return;
    }

    // sheetに書き込む
    writeEvent({ event: event, id: id, isRecord: false }, sheet);
}

function _writeCalendar(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    changedRange: GoogleAppsScript.Spreadsheet.Range,
    setting: SettingInfo
): void {
    // 同期対象でないsheetの場合は何もしない
    if (
        !(sheet
            .getRange(
                setting.record.isSync.row,
                setting.record.isSync.column,
                1,
                1
            )
            .getValue() as boolean)
    ) {
        console.log(
            `synchronization of sheet '${sheet.getName()}' is not available.`
        );
        return;
    }

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

    // 更新する
    updateEventFromSheet(
        fixedRowIndex,
        Math.min(sheet.getLastRow(), changedRange.getLastRow()) -
            fixedRowIndex +
            1,
        sheet,
        setting
    );
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function writeCalendar(e: OnEditEventObject): void {
    // 変更されたのが設定情報の場合
    /* if (e.source.getActiveSheet().getName() == 'settings') { */
    /*     return; */
    /* } */

    const sheet = e.source.getActiveSheet();
    if (!sheet) {
        console.error("the target sheet doesn't exist.");
        return;
    }
    console.log(`the current sheet: ${sheet.getName()}`);

    const settings = SettingManager.load();
    if (!settings) return;

    if (!settings.isSync) {
        // calendar用のsheetでなければ何もしない
        console.log('synchronization is not available.');
        return;
    }

    // SpreadSheetからdateの予定を
    // 取得
    _writeCalendar(sheet, e.range, settings);
}
