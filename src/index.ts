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

    public SetEvent(event: Event): void
    {
        this.calendar.createEvent(
            event.title,
            event.start,
            event.end,
            event.option
        );
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
    return new Date(year, month - 1, day, hour - 13, minute);
}

function writeCalendar(date: Date): void
{
    // SpreadSheetからdateの予定を
    // 取得
    const sheet = SpreadsheetApp.getActiveSheet();
    if (sheet == null)
    {
        console.log("the target sheet doesn't exist.");
        return undefined;
    }

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
    const records = sheet
        .getRange(nowDate[4], nowDate[3], sheet.getLastRow() - 1, 6)
        .getValues();
    const recordCalendar = new Calendar(
        '4b1q49ogqlkrucdhhgap8k5g9c@group.calendar.google.com'
    );

    // 書き込み
    for (const record of records)
    {
        if (record[0] == '') break;
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
        recordCalendar.SetEvent(event);
        console.log(`done.`);
    }
}
