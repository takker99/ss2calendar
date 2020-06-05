import * as moment from 'moment';
import { OnEditEventObject } from './EventObject';
import { SettingManager } from './settingManager';
const Moment = { moment: moment };

// userが編集したときに発火する函数
// scriptの編集には反応しないので注意
// eslint-disable-next-line @typescript-eslint/no-unused-vars
function onEdit(e: OnEditEventObject): void {
    console.log(`AuthMode: ${e.authMode}`);
    console.log(`Changed range: ${e.range}`);
    console.log(`User: ${e.user}`);
    console.log(`Changed spread sheet: ${e.source}`);
    if (e.oldValue) {
        console.log(`s/${e.oldValue}/${e.value}`);
    } else {
        console.log('More than one cells are changed');
    }
    console.log(
        `Trigger ID: ${e.triggerUid ? e.triggerUid : 'simple trigger'}`
    );
}

// 機能テスト用のscript
// eslint-disable-next-line @typescript-eslint/no-unused-vars
function myFunction(): void {
    // sheet&rangeの取得
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
        'record'
    );
    if (sheet == null) {
        return undefined;
    }
    const range = sheet.getRange(20, 2, 1, 1);

    // 書式を指定して値を取得する
    // 書式をあとで戻すために、現在の書式情報を保存しておく
    const temp = range.getNumberFormats();
    range.setNumberFormats([['@']]);
    const rawData = range.getValues();
    console.log(`時刻形式のdataの中身は${rawData[0][0]}でした`);
    range.setNumberFormats(temp);
}

// 新しい記録dataを追加する
// 終了時刻は空白にする
// eslint-disable-next-line @typescript-eslint/no-unused-vars
function addRecord(): void {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
        'record'
    );
    if (sheet == null) {
        return undefined;
    }
    const settingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
        'settings'
    );
    if (settingSheet == null) return;
    const settings = new SettingManager(settingSheet);

    const now = Moment.moment().zone('+09:00');
    console.log(`現在時刻: ${now}`);
    const nowDate = [now.year(), now.month() + 1, now.date()];
    const nowTime = `${now.hours()}:${now.minutes()}`;

    sheet.appendRow(['test', ...nowDate, nowTime, ...nowDate, '']);

    // 数式を記入
    sheet
        .getRange(sheet.getLastRow(), settings.record.index.timeSpan)
        .setFormulaR1C1('=RC[-1]-RC[-5]');

    // 入力規則を追加
    console.log(`emotion: ${settings.emotionList.getValues() as string[][]}`);
    const rules: GoogleAppsScript.Spreadsheet.DataValidation[] = [
        SpreadsheetApp.newDataValidation()
            .requireValueInRange(settings.emotionList)
            .build(),
    ];

    console.log(`tags: ${settings.tagList}`);
    rules[1] = SpreadsheetApp.newDataValidation()
        .requireValueInRange(settings.calendarIdList)
        .build();

    let temp = sheet.getRange(
        sheet.getLastRow(),
        settings.record.index.emotionTag
    );
    temp.clearDataValidations();
    temp.setDataValidation(rules[0]);
    temp = sheet.getRange(sheet.getLastRow(), settings.record.index.tag);
    temp.clearDataValidations();
    temp.setDataValidation(rules[1]);
}
