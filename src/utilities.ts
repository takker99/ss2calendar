import { SettingManager } from './settingManager';
import { updateEventFromSheet } from './index';
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

    const titles = sheet.getRange(
        settings.record.firstLine,
        settings.record.write.schedule.title,
        sheet.getLastRow(),
        1
    );

    // 条件付き書式を再作成する
    //   各task終了時間に応じてtaskの色分けをする
    sheet.clearConditionalFormatRules();
    const rules = [
        [
            `=hour(timevalue($G${settings.record.firstLine}))*60+minute(timevalue($G${settings.record.firstLine}))<4*60`,
            '#57bb8a',
        ],
        [
            `=hour(timevalue($G${settings.record.firstLine}))*60+minute(timevalue($G${settings.record.firstLine}))<8*60`,
            '#b7e1cd',
        ],
        [
            `=hour(timevalue($G${settings.record.firstLine}))*60+minute(timevalue($G${settings.record.firstLine}))<12*60`,
            '#ffd666',
        ],
        [
            `=hour(timevalue($G${settings.record.firstLine}))*60+minute(timevalue($G${settings.record.firstLine}))<16*60`,
            '#f7981d',
        ],
        [
            `=hour(timevalue($G${settings.record.firstLine}))*60+minute(timevalue($G${settings.record.firstLine}))<20*60`,
            '#e67c13',
        ],
    ].map((values: string[]) =>
        SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(values[0])
            .setBackground(values[1])
            .setRanges([titles])
            .build()
    );
    // 文字色も変更するので、別に加える
    rules.push(
        SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(
                `=hour(timevalue($G${settings.record.firstLine}))*60+minute(timevalue($G${settings.record.firstLine}))<24*60`
            )
            .setBackground('#351c75')
            .setFontColor('#FFFFFF')
            .setRanges([titles])
            .build()
    );
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

    // 並び替える
    sheet.sort(settings.record.read.schedule.start);
}

// 現在時刻を終了時刻として,選択行の実績欄に経過時間を書き込む。また次の行の開始時刻に現在時刻を記入する
// eslint-disable-next-line @typescript-eslint/no-unused-vars
function writeTimestamp(): void {
    //現在時刻を取得
    const now = Moment.moment().zone('+09:00');

    const sheet = SpreadsheetApp.getActiveSheet();
    if (sheet == null) {
        console.error("the target sheet doesn't exist.");
        return;
    }

    const settings = SettingManager.load();
    if (!settings) return;

    // 選択範囲を取得
    const selectedRange = sheet.getActiveRange();
    if (!selectedRange) return; // 選択されてなければ何もしない

    // 複数行選択されているときは、最初の行に書き込む
    const targetRow = selectedRange.getRow();

    // 既に値が書き込まれていたら何もしない
    if (
        sheet
            .getRange(targetRow, settings.record.write.record.duration, 1, 1)
            .getValue() != ''
    )
        return;

    // 作業開始時刻を取得
    const startDateTime = Moment.moment
        .unix(
            sheet
                .getRange(targetRow, settings.record.read.record.start, 1, 1)
                .getValue() as number
        )
        .subtract(9, 'hour');

    //書き込む
    const duration = Moment.moment.duration(now.diff(startDateTime));
    const zero = (n: number): string => String(n).padStart(2, '0');
    sheet
        .getRange(targetRow, settings.record.write.record.duration, 1, 1)
        .setValue(`${zero(duration.hours())}:${zero(duration.minutes())}`);

    //次行に現在時刻を書き込む
    const nowDate = now.toDate();
    sheet
        .getRange(targetRow + 1, settings.record.write.record.start.time, 1, 1)
        .setValue(Utilities.formatDate(nowDate, 'JST', 'HH:mm'));
    //Calendarに更新を反映する
    updateEventFromSheet(targetRow, 2, sheet, settings);
}

// 選択中の全ての行の開始時刻を一括してずらす
// eslint-disable-next-line @typescript-eslint/no-unused-vars
function changeTimes(): void {
    const sheet = SpreadsheetApp.getActiveSheet();
    if (sheet == null) {
        console.error("the target sheet doesn't exist.");
        return;
    }

    const settings = SettingManager.load();
    if (!settings) return;

    // 選択範囲を取得
    const selectedRange = sheet.getActiveRange();
    if (!selectedRange) return; // 選択されてなければ何もしない

    // ずらす時間幅をuserに聞く(単位はmin)
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
        'Offset',
        'Please write duration by minutes',
        ui.ButtonSet.OK_CANCEL
    );

    // 1. OK button以外が押された
    // 2. 入力された値が数値でない
    // 3. 何も入力されなかった
    // ときは何もしない
    if (
        response.getSelectedButton() != ui.Button.OK ||
        isNaN(Number(response.getResponseText())) ||
        response.getResponseText() == ''
    )
        return;

    // 予定開始時刻が記入されているセルを取得
    const targetRangeLength =
        selectedRange.getLastRow() - selectedRange.getRow() + 1;
    // 転置に使うlambda expression
    const transpose = <T>(a: T[][]): T[][] =>
        a[0].map((_: T, index: number): T[] => a.map((r: T[]): T => r[index]));
    // 変更する予定開始時刻を全て取得し、指定された時間(分単位)だけずらす
    const startDateTimes = transpose(
        sheet
            .getRange(
                selectedRange.getRow(),
                settings.record.read.schedule.start,
                targetRangeLength,
                1
            )
            .getValues() as number[][]
    )[0].map((value) =>
        Moment.moment
            .unix(value)
            .subtract(9, 'hour')
            .add(Number(response.getResponseText()), 'minute')
            .toDate()
    );

    //予定開始時刻を更新する
    sheet
        .getRange(
            selectedRange.getRow(),
            settings.record.write.schedule.start.time,
            targetRangeLength,
            1
        )
        .setValues(
            transpose([
                startDateTimes.map((date) =>
                    Utilities.formatDate(date, 'JST', 'HH:mm')
                ),
            ])
        );
    //Calendarに更新を反映する
    updateEventFromSheet(
        selectedRange.getRow(),
        targetRangeLength,
        sheet,
        settings
    );
}
