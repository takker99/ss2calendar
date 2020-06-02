import * as moment from 'moment';
const Moment = { moment: moment };

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
    if (settingSheet == null) return undefined;
    const settings = settingSheet.getRange(1, 2, 25, 1).getValues();

    const now = Moment.moment().zone('+09:00');
    console.log(`現在時刻: ${now}`);
    const nowDate = [now.year(), now.month() + 1, now.date()];
    const nowTime = `${now.hours()}:${now.minutes()}`;

    sheet.appendRow(['test', ...nowDate, nowTime, ...nowDate, '']);

    // 数式を記入
    const durationColumn = settings[17][0];
    sheet
        .getRange(sheet.getLastRow(), durationColumn)
        .setFormulaR1C1('=RC[-1]-RC[-5]');

    // 入力規則を追加
    const emotionList = {
        row: settings[1][0],
        column: settings[2][0],
        length: settings[3][0],
    };
    console.log(
        `emotion: [row,column,length]=[${emotionList.row},${emotionList.column},${emotionList.length}]`
    );
    const rules: GoogleAppsScript.Spreadsheet.DataValidation[] = [
        SpreadsheetApp.newDataValidation()
            .requireValueInRange(
                settingSheet.getRange(
                    emotionList.row,
                    emotionList.column,
                    emotionList.length,
                    1
                )
            )
            .build(),
    ];

    const tagList = {
        row: settings[4][0],
        column: settings[5][0],
        length: settings[6][0],
    };
    console.log(
        `emotion: [row,column,length]=[${tagList.row},${tagList.column},${tagList.length}]`
    );
    rules[1] = SpreadsheetApp.newDataValidation()
        .requireValueInRange(
            settingSheet.getRange(
                tagList.row,
                tagList.column,
                tagList.length,
                1
            )
        )
        .build();

    const tagColumn = settings[8][0];
    const emotionColumn = settings[22][0];

    let temp = sheet.getRange(sheet.getLastRow(), emotionColumn);
    temp.clearDataValidations();
    temp.setDataValidation(rules[0]);
    temp = sheet.getRange(sheet.getLastRow(), tagColumn);
    temp.clearDataValidations();
    temp.setDataValidation(rules[1]);
}
