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
    // シートにあるデータから年・月・日を取得
    const rowTemp: number[] = sheet.getRange(1, 2, 5, 1).getValues();
    const nowDate: number[] = [
        rowTemp[0][0],
        rowTemp[1][0],
        rowTemp[2][0],
        rowTemp[3][0],
        rowTemp[4][0],
    ];
    console.log(`date: ${nowDate[0]}/${nowDate[1]}/${nowDate[2]}`);
    console.log(`データ開始行: ${nowDate[4]}`);
    console.log(`データ開始列: ${nowDate[3]}`);

    if (nowDate[3] == '')
    {
        console.log(
            '記録用のデータがありません。データの列の位置がずれている可能性があります'
        );
        return undefined;
    }

    // sheetから記録を入手
    // 1. task name
    // 2. remark
    // 3. 開始時刻のhh
    // 4. 開始時刻のmm
    // 5. 終了時刻のhh
    // 6. 終了時刻のmm
    // 7. event ID
    const records = sheet
        .getRange(nowDate[4], nowDate[3], sheet.getLastRow() - 1, 7)
        .getValues();

    // 書き込み
    for (let i = 0; i < records.length; i++)
    {
        const record = records[i];

        // task nameが空白のときは
        // 読み飛ばす
        if (record[0] == '') continue;
        console.log(`setting the event '${record[0]}'...`);
        console.log(`start time: ${record[2]}:${record[3]}`);
        const period = new TimeSpan(
            getDateFixed(
                nowDate[0],
                nowDate[1],
                nowDate[2],
                record[2],
                record[3]
            ),
            getDateFixed(
                nowDate[0],
                nowDate[1],
                nowDate[2],
                record[4],
                record[5]
            )
        );
        console.log(`start: ${period.start}`);
        console.log(`end: ${period.end}`);
        console.log(`description: ${record[1]}`);

        const event: Event = new Event(record[0], period, record[1]);

        // 既に登録済みの記録であれば、更新する
        if (record[6] != '')
        {
            recordCalendar.ModifyEvent(record[6], event);
            console.log(`done.`);
            continue;
        }

        // event IDを新規登録する
        const eventId: string = recordCalendar.SetEvent(event);
        console.log(`event ID: ${eventId}`);
        sheet.getRange(nowDate[4] + i, nowDate[3] + 6).setValue(eventId);
        console.log(`done.`);
    }
}

function _writeSchedule(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    recordCalendar: GoogleAppsScript.Calendar
): void
{

    // 二次元配列転置用lambda式
    // シートにあるデータから年・月・日を取得
    const rowTemp: number[] = sheet.getRange(1, 2, 5, 1).getValues();
    const nowDate: number[] = [
        rowTemp[0][0],
        rowTemp[1][0],
        rowTemp[2][0],
        rowTemp[3][0],
        rowTemp[4][0],
    ];
    console.log(`date: ${nowDate[0]}/${nowDate[1]}/${nowDate[2]}`);
    console.log(`データ開始行: ${nowDate[4]}`);
    console.log(`データ開始列: ${nowDate[3]}`);

    if (nowDate[3] == '')
    {
        console.log(
            '記録用のデータがありません。データの列の位置がずれている可能性があります'
        );
        return undefined;
    }

    // sheetから記録を入手
    // 1. task name
    // 2. remark
    // 3. 開始時刻のhh
    // 4. 開始時刻のmm
    // 5. 終了時刻のhh
    // 6. 終了時刻のmm
    // 7. event ID
    const records = sheet
        .getRange(nowDate[4], nowDate[3], sheet.getLastRow() - 1, 11)
        .getValues();

    // 書き込み
    for (let i = 0; i < records.length; i++)
    {
        const record = records[i];

        // task nameが空白のときは
        // 読み飛ばす
        if (record[0] == '') continue;
        console.log(`setting the event '${record[0]}'...`);
        console.log(`start time: ${record[2]}/${record[3]} ${record[4]}:${record[5]}`);
        const period = new TimeSpan(
            getDateFixed(
                nowDate[0],
                record[2],
                record[3],
                record[4],
                record[5]
            ),
            getDateFixed(
                nowDate[0],
                record[6],
                record[7],
                record[8],
                record[9]
            )
        );
        console.log(`start: ${period.start}`);
        console.log(`end: ${period.end}`);
        console.log(`description: ${record[1]}`);

        const event: Event = new Event(record[0], period, record[1]);

        // 既に登録済みの記録であれば、更新する
        if (record[10] != '')
        {
            recordCalendar.ModifyEvent(record[10], event);
            console.log(`done.`);
            continue;
        }

        // event IDを新規登録する
        const eventId: string = recordCalendar.SetEvent(event);
        console.log(`event ID: ${eventId}`);
        sheet.getRange(nowDate[4] + i, nowDate[3] + 10).setValue(eventId);
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
    _writeCalendar(sheet,
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
    _writeSchedule(sheet,
        new Calendar('kua4bd6695fov7jrl9cmfu3o7o@group.calendar.google.com')
    );
}
