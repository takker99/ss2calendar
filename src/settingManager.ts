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
        const emotionListLocation = {
            startRow: temp[1],
            startColumn: temp[2],
            length: temp[3],
        };
        this.emotionList = transpose(
            settingSheet
                .getRange(
                    emotionListLocation.startRow,
                    emotionListLocation.startColumn,
                    emotionListLocation.length,
                    1
                )
                .getValues() as string[][]
        )[0];

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
    }
    // true: 同期が有効 false:同期を停止
    public readonly isSync: boolean;
    public readonly emotionList: string[];
    public readonly tagList: { [index: string]: string } = {};
}
