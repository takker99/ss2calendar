import { SettingManager } from './settingManager';
import { Event, TimeSpan } from './calendar';
import { writeEvent, Record } from './index';
import * as moment from 'moment';
const Moment = { moment: moment }; // GAS対策 cf. https://qiita.com/awa2/items/d24df6abd5fd5e4ca3d9

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
                `=hour(timevalue($I${settings.record.firstLine}))*60+minute(timevalue($I${settings.record.firstLine}))<4*60`
            )
            .setBackground('#57bb8a')
            .setRanges([titles])
            .build(),
        SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(
                `=hour(timevalue($I${settings.record.firstLine}))*60+minute(timevalue($I${settings.record.firstLine}))<8*60`
            )
            .setBackground('#b7e1cd')
            .setRanges([titles])
            .build(),
        SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(
                `=hour(timevalue($I${settings.record.firstLine}))*60+minute(timevalue($I${settings.record.firstLine}))<12*60`
            )
            .setBackground('#ffd666')
            .setRanges([titles])
            .build(),
        SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(
                `=hour(timevalue($I${settings.record.firstLine}))*60+minute(timevalue($I${settings.record.firstLine}))<16*60`
            )
            .setBackground('#f7981d')
            .setRanges([titles])
            .build(),
        SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(
                `=hour(timevalue($I${settings.record.firstLine}))*60+minute(timevalue($I${settings.record.firstLine}))<20*60`
            )
            .setBackground('#e67c13')
            .setRanges([titles])
            .build(),
        SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(
                `=hour(timevalue($I${settings.record.firstLine}))*60+minute(timevalue($I${settings.record.firstLine}))<24*60`
            )
            .setBackground('#351c75')
            .setFontColor('#FFFFFF')
            .setRanges([titles])
            .build(),
    ];
    sheet.setConditionalFormatRules(rules);
}

// recordを開始時刻の古い順に並び替える
// eslint-disable-next-line @typescript-eslint/no-unused-vars
function sortRecord(): void {
    const sheet = SpreadsheetApp.getActiveSheet();

    if (sheet == null) {
        console.error("the target sheet doesn't exist.");
        return;
    }

    const settings = SettingManager.load();
    if (settings == undefined) return;
    console.log('Got the setting information');
    console.log(`settings: ${JSON.stringify(settings)}`);

    // 並び替える
    sheet.sort(settings.record.columnEnd + 1);
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
