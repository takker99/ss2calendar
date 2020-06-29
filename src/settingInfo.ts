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
        const settingRowLength = 28;
        const temp = transpose(
            settingSheet
                .getRange(1, 2, settingRowLength, 1)
                .getValues() as number[][]
        )[0];

        this.isSync = temp[0] == 1 ? true : false;
        this.record = {
            firstLine: temp[7],
            columnFlont: temp[8],
            columnEnd: temp[9],
            write: {
                tag: temp[10],
                start: {
                    year: temp[11],
                    month: temp[12],
                    date: temp[13],
                    time: temp[14],
                },
                duration: temp[15],
                title: temp[16],
                expectation: temp[17],
                actualAction: temp[18],
                emotionTag: temp[19],
                remarks: temp[20],
                eventId: temp[21],
            },
            read: {
                start: temp[22],
                end: temp[23],
                title: temp[24],
                description: temp[25],
                eventId: temp[26],
                calendarId: temp[27],
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
        firstLine: number;
        columnFlont: number;
        columnEnd: number;
        write: {
            tag: number;
            start: { year: number; month: number; date: number; time: number };
            duration: number;
            title: number;
            expectation: number;
            actualAction: number;
            emotionTag: number;
            remarks: number;
            eventId: number;
        };
        read: {
            start: number;
            end: number;
            title: number;
            description: number;
            eventId: number;
            calendarId: number;
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

export function writingAreaLength(setting: SettingInfo): number {
    const temp = [
        setting.record.write.tag,
        setting.record.write.start.year,
        setting.record.write.start.month,
        setting.record.write.start.date,
        setting.record.write.start.time,
        setting.record.write.duration,
        setting.record.write.title,
        setting.record.write.expectation,
        setting.record.write.actualAction,
        setting.record.write.emotionTag,
        setting.record.write.remarks,
        setting.record.write.eventId,
    ];
    return Math.max(...temp) - Math.min(...temp) + 1;
}
