// userが編集したときに発火する函数
// scriptの編集には反応しないので注意
// eslint-disable-next-line @typescript-eslint/no-unused-vars
/* function onEdit(e: OnEditEventObject): void { */
/*     console.log(`AuthMode: ${e.authMode}`); */
/*     console.log(`Changed range: ${e.range}`); */
/*     console.log(`User: ${e.user}`); */
/*     console.log(`Changed spread sheet: ${e.source}`); */
/*     if (e.oldValue) { */
/*         console.log(`s/${e.oldValue}/${e.value}`); */
/*     } else { */
/*         console.log('More than one cells are changed'); */
/*     } */
/*     console.log( */
/*         `Trigger ID: ${e.triggerUid ? e.triggerUid : 'simple trigger'}` */
/*     ); */
/* } */

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
