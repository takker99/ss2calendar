class Minutes
{
    constructor(private minutes: number)
    {
        this.minutes = minutes;
    }

    /**
     * 時間の長さをmilisecond単位で取得する
     *
     * @return 時間の長さ(miliseconds)
     */
    public getTime(): number
    {
        return this.minutes * 60000;
    }
}

function add(date: Date, minutes: Minutes): Date
{
    const temp: number = date.getTime() + minutes.getTime();
    return new Date(temp);
}

class TimeSpan
{
    constructor(begin: Date, length: Minutes);
    constructor(begin: Date, end: Date);
    constructor(begin: number, end: number);

    constructor(begin: Date | number, value: Minutes | Date | number)
    {
        if (begin instanceof Date)
        {
            this._start = new Date(begin.getTime());
        }
        else
        {
            this._start = new Date(begin);
        }

        if (value instanceof Date)
        {
            this._end = new Date(value.getTime());
        }

        if (value instanceof Minutes)
        {
            this._end = add(this._start, value);
        }
        if (typeof value == 'number')
        {
            this._end = new Date(value);
        }
    }

    public AddMonth(months: number): void
    {
        this._end.setMonth(this._end.getMonth() + months);
    }

    public get start(): Date
    {
        return new Date(this._start.getTime());
    }

    public get end(): Date
    {
        return new Date(this._end.getTime());
    }

    private _start: Date;
    private _end: Date;
}

type CalendarEvent = GoogleAppsScript.Calendar.CalendarEvent;

class Calendar
{
    constructor(calendarId: string)
    {
        this.calendar = CalendarApp.getCalendarById(calendarId);
    }

    public GetEvents(period: TimeSpan): CalendarEvent[]
    {
        return this.calendar.getEvents(period.start, period.end);
    }

    public ModifyEvent(eventId: string, newEvent: Event): void
    {
        const event = this.calendar.getEventById(eventId);
        if (event.getTitle() != newEvent.title) event.setTitle(newEvent.title);
        if (
            event.getStartTime() != newEvent.start &&
            event.getEndTime() != newEvent.end
        )
            event.setTime(newEvent.start, newEvent.end);
        if (event.getDescription() != newEvent.option.description)
            event.setDescription(newEvent.option.description);
    }

    public SetEvent(event: Event): string
    {
        const result = this.calendar.createEvent(
            event.title,
            event.start,
            event.end,
            event.option
        );
        return result.getId();
    }

    private calendar: GoogleAppsScript.Calendar.Calendar;
}

class Event
{
    constructor(
        title: string,
        private period: TimeSpan,
        private description: string
    )
    {
        this._title = title;
        this.period = new TimeSpan(
            period.start.getTime(),
            period.end.getTime()
        );
        this.description = description;
    }

    public get title(): string
    {
        return this._title;
    }

    public get start(): Date
    {
        return this.period.start;
    }

    public get end(): Date
    {
        return this.period.end;
    }

    public get option(): {description: string}
    {
        return {description: this.description};
    }

    private _title: string;
}

function getDateFixed(
    year: number,
    month: number,
    day: number,
    hour: number,
    minute: number
): Date
{
    return new Date(year, month - 1, day, hour, minute);
}

function _writeCalendar(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    recordCalendar: GoogleAppsScript.Calendar
): void
{
    // 二次元配列転置用lambda式
    // シートにあるデータから
    // - calendarと同期するevent dataの開始cellの位置
    // を取得
    const rowTemp: number[] = sheet.getRange(1, 2, 2, 1).getValues();
    const schemes: {row: number; column: number} = {
        row: rowTemp[1][0],
        column: rowTemp[0][0],
    };
    console.log(`データ開始行: ${schemes.row}`);
    console.log(`データ開始列: ${schemes.column}`);

    if (isNaN(schemes.row))
    {
        console.log(
            '記録用のデータがありません。データの列の位置がずれている可能性があります'
        );
        return undefined;
    }

    interface Record
    {
        event: Event;
        id: string;
    }
    // sheetから記録を入手
    //  1. task name
    //  2. remark
    //  3. 開始時刻のyyyy
    //  4. 開始時刻のMM
    //  5. 開始時刻のdd
    //  6. 開始時刻のhh
    //  7. 開始時刻のmm
    //  8. 終了時刻のyyyy
    //  9. 終了時刻のMM
    // 10. 終了時刻のdd
    // 11. 終了時刻のhh
    // 12. 終了時刻のmm
    // 13. event ID
    const records: Record[] = sheet
        .getRange(schemes.row, schemes.column, sheet.getLastRow() - 1, 13)
        .getValues()
        .map((record) =>
        {
            return {
                event: new Event(
                    record[0],
                    new TimeSpan(
                        getDateFixed(
                            record[2],
                            record[3],
                            record[4],
                            record[5],
                            record[6]
                        ),
                        getDateFixed(
                            record[7],
                            record[8],
                            record[9],
                            record[10],
                            record[11]
                        )
                    ),
                    record[1]
                ),
                id: record[12],
            };
        });

    // 書き込み
    for (let i = 0; i < records.length; i++)
    {
        const record = records[i];

        // task nameが空白のときは
        // 読み飛ばす
        if (record.event.title == '') continue;
        console.log(`setting the event '${record.event.title}'...`);
        console.log(`start: ${record.event.start}`);
        console.log(`end: ${record.event.end}`);
        console.log(`description: ${record.event.description}`);

        // 既に登録済みの記録であれば、更新する
        if (record.id != '')
        {
            recordCalendar.ModifyEvent(record.id, record.event);
            console.log(`done.`);
            continue;
        }

        // event IDを新規登録する
        const eventId: string = recordCalendar.SetEvent(record.event);
        console.log(`event ID: ${eventId}`);
        sheet
            .getRange(schemes.row + i, schemes.column + 13 - 1)
            .setValue(eventId);
        console.log(`done.`);
    }
}

function writeCalendar(): void
{
    const sheet = SpreadsheetApp.getActiveSheet();
    // SpreadSheetからdateの予定を
    // 取得
    if (sheet == null)
    {
        console.log("the target sheet doesn't exist.");
        return undefined;
    }
    _writeCalendar(
        sheet,
        new Calendar('2p339s4tkeoq57u649ul41e57o@group.calendar.google.com')
    );
}

function writeSchedule(): void
{
    const sheet = SpreadsheetApp.getActiveSheet();
    // SpreadSheetからdateの予定を
    // 取得
    if (sheet == null)
    {
        console.log("the target sheet doesn't exist.");
        return undefined;
    }
    _writeCalendar(
        sheet,
        new Calendar('kua4bd6695fov7jrl9cmfu3o7o@group.calendar.google.com')
    );
}
