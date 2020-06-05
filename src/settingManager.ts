import * as moment from 'moment';
const Moment = { moment: moment }; // GAS対策 cf. https://qiita.com/awa2/items/d24df6abd5fd5e4ca3d9

// spread sheetから設定情報を抽出するclass
// spread sheetに記述する設定情報の形式はこのclassで決定し、他のclassからその形式は隠蔽する
export class SettingManager {
    constructor(settingSheet: GoogleAppsScript.Spreadsheet.Sheet) {
        // 転置に使うlambda expression
        const transpose = <T>(a: T[][]): T[][] =>
            a[0].map((_, c) => a.map((r) => r[c]));
        const temp = transpose(
            settingSheet.getRange(1, 2, 27, 1).getValues() as number[][]
        )[0];

        this.isSync = temp[1] == 1 ? true : false;
        this.record = {
            firstLine: temp[7],
            columnFlont: temp[8],
            columnEnd: temp[9],
        };

        const emotionListLocation = {
            startRow: temp[1],
            startColumn: temp[2],
            length: temp[3],
        };
        this.emotionList = settingSheet.getRange(
            emotionListLocation.startRow,
            emotionListLocation.startColumn,
            emotionListLocation.length,
            1
        );

        const tagListLocation = {
            startRow: temp[4],
            startColumn: temp[5],
            length: temp[6],
        };
        const rawData = settingSheet
            .getRange(
                tagListLocation.startRow,
                tagListLocation.startColumn,
                tagListLocation.length,
                2
            )
            .getValues() as string[][];
        for (const pair of rawData) {
            this.tagList[pair[0]] = pair[1];
        }
        this.calendarIdList = settingSheet.getRange(
            tagListLocation.startRow,
            tagListLocation.startColumn + 1,
            tagListLocation.length,
            1
        );
    }

    public recordLength(): number {
        return this.record.columnEnd - this.record.columnFlont + 1;
    }
    // true: 同期が有効 false:同期を停止
    public readonly isSync: boolean;
    public readonly record: {
        firstLine: number;
        columnFlont: number;
        columnEnd: number;
        index: {
            start: { year: number; month: number; date: number; time: number };
            end: { year: number; month: number; date: number; time: number };
            timeSpan: number;
            title: number;
            expectation: number;
            actualAction: number;
            emotionTag: number;
            remarks: number;
            eventId: number;
            calendarId: number;
        };
    };
    public readonly emotionList: GoogleAppsScript.Spreadsheet.Range;
    public readonly calendarIdList: GoogleAppsScript.Spreadsheet.Range;
    public readonly tagList: { [index: string]: string } = {};
}
