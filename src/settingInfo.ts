// spread sheetから設定情報を抽出するclass
// spread sheetに記述する設定情報の形式はこのclassで決定し、他のclassからその形式は隠蔽する
export class SettingInfo {
    constructor(settingSheet: GoogleAppsScript.Spreadsheet.Sheet) {
        // 転置に使うlambda expression
        const transpose = <T>(a: T[][]): T[][] =>
            a[0].map((_: T, index: number): T[] =>
                a.map((r: T[]): T => r[index])
            );

        // 設定情報が含まれたセルの行数
        const settingRowLength = 40;
        const temp = transpose(
            settingSheet
                .getRange(1, 2, settingRowLength, 1)
                .getValues() as number[][]
        )[0];

        this.isSync = temp[0] == 1 ? true : false;
        this.record = {
            isSync: {
                row: temp[38],
                column: temp[39],
            },
            firstLine: temp[7],
            columnFlont: temp[8],
            columnEnd: temp[9],
            write: {
                schedule: {
                    tag: temp[10],
                    start: {
                        year: temp[11],
                        month: temp[12],
                        date: temp[13],
                        time: temp[14],
                    },
                    duration: temp[15],
                    title: temp[16],
                    description: temp[18],
                    eventId: temp[24],
                },
                record: {
                    start: {
                        year: temp[11],
                        month: temp[12],
                        date: temp[13],
                        time: temp[19],
                    },
                    duration: temp[20],
                    title: temp[16],
                    description: temp[21],
                    emotion: temp[22],
                    remarks: temp[23],
                    eventId: temp[25],
                },
            },
            read: {
                schedule: {
                    canUpdate: temp[26],
                    start: temp[28],
                    end: temp[29],
                    title: temp[32],
                    description: temp[33],
                    eventId: temp[34],
                    calendarId: temp[35],
                },
                record: {
                    canUpdate: temp[27],
                    start: temp[30],
                    end: temp[31],
                    title: temp[32],
                    description: temp[36],
                    eventId: temp[37],
                },
            },
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
        this.tagList = settingSheet.getRange(
            tagListLocation.startRow,
            tagListLocation.startColumn,
            tagListLocation.length,
            1
        );

        this.calendarTagging = settingSheet
            .getRange(
                tagListLocation.startRow,
                tagListLocation.startColumn,
                tagListLocation.length,
                2
            )
            .getValues();
    }

    // true: 同期が有効 false:同期を停止
    public readonly isSync: boolean;
    public readonly record: {
        isSync: { row: number; column: number }; // sheet毎の同期設定
        firstLine: number;
        columnFlont: number;
        columnEnd: number;
        write: {
            schedule: {
                tag: number;
                start: {
                    year: number;
                    month: number;
                    date: number;
                    time: number;
                };
                duration: number;
                title: number;
                description: number;
                eventId: number;
            };
            record: {
                start: {
                    year: number;
                    month: number;
                    date: number;
                    time: number;
                };
                duration: number;
                title: number;
                description: number;
                emotion: number;
                remarks: number;
                eventId: number;
            };
        };
        read: {
            schedule: {
                canUpdate: number;
                start: number;
                end: number;
                title: number;
                description: number;
                eventId: number;
                calendarId: number;
            };
            record: {
                canUpdate: number;
                start: number;
                end: number;
                title: number;
                description: number;
                eventId: number;
            };
        };
    };
    public readonly emotionList: GoogleAppsScript.Spreadsheet.Range;
    public readonly tagList: GoogleAppsScript.Spreadsheet.Range;
    public readonly calendarTagging: string[][];
}
export function toTag(setting: SettingInfo, calendarId: string): string {
    const temp = setting.calendarTagging.find(
        (tagging) => tagging[1] == calendarId
    );
    if (temp) {
        return temp[0];
    } else {
        return '';
    }
}

export function recordLength(setting: SettingInfo): number {
    return setting.record.columnEnd - setting.record.columnFlont + 1;
}

// 記録/予定に使用しているsheetの列幅を返す
// eslint-disable-next-line @typescript-eslint/no-unused-vars
export function writingAreaLength(setting: SettingInfo): number {
    // 実装方法が思いつかないので、その場しのぎの方法を使う
    return 38;
}
