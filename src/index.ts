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

    constructor(private begin: Date | number, value: Minutes | Date | number)
    {
        if (begin instanceof Date)
        {
            this.begin = new Date(begin.getTime());
        }
        else
        {
            this.begin = new Date(begin);
        }

        if (value instanceof Date)
        {
            this.end = new Date(value.getTime());
        }

        if (value instanceof Minutes)
        {
            this.end = add(this.begin, value);
        }
        if (typeof value == 'number')
        {
            this.end = new Date(value);
        }
    }

    public AddMonth(months: number): void
    {
        this.end.setMonth(this.end.getMonth() + months);
    }

    public get start(): Date
    {
        return new Date(this.begin.getTime());
    }

    public get end(): Date
    {
        return new Date(this.end.getTime());
    }

    private end: Date;
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
        private title: string,
        private period: TimeSpan,
        private description: string
    )
    {
        this.title = title;
        this.period = new TimeSpan(
            period.start.getTime(),
            period.start.getTime()
        );
        this.description = description;
    }

    public get title(): string
    {
        return this.title;
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
}

function writeCalendar(date: Date): void
{
    // SpreadSheetからdateの予定を
    // 取得
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
        '2020-04-30'
    );
    if (sheet == null)
    {
        console.log("the target sheet doesn't exist.");
        return undefined;
    }

    // 一旦クリア
    sheet.clear();

    // Calendarから予定を取得
    const holidayId = CalendarApp.getCalendarsByName('日本の祝日')[0].getId();
    const holydayCalendar = new Calendar(holidayId);
    const start: Date = new Date(Date.now());
    const end: Date = new Date(start.getTime());
    end.setMonth(end.getMonth() + 12); // 12ヶ月まで予定を取得する
    const timeSpan = new TimeSpan(start, end);

    const holidays: CalendarEvent[] = holydayCalendar.GetEvents(timeSpan);

    // sheetに予定を書き込む
    const temp: string[][] = sheet.getRange(1,1,sheet.getLastRow()-1,6).getValues();
    for (let index = 2; index < holidays.length + 2; index++)
    {
        const holiday = holidays[index - 1];
        temp[index] = [];
        temp[index] = [
            holiday.getTitle(), //イベントタイトル
            //イベント開始時刻
            holiday.getStartTime().getFullYear(),
            holiday.getStartTime().getMonth(),
        ];
        sheet.getRange(index, 2).setValue(holiday.getStartTime().getFullYear());
        sheet
            .getRange(index, 3)
            .setValue(holiday.getStartTime().getMonth() + 1);
        sheet.getRange(index, 4).setValue(holiday.getStartTime().getDate());
        sheet.getRange(index, 5).setValue(holiday.getStartTime().getHours());
        sheet.getRange(index, 6).setValue(holiday.getStartTime().getMinutes());
        //イベント終了時刻
        sheet.getRange(index, 7).setValue(holiday.getEndTime().getFullYear());
        sheet.getRange(index, 8).setValue(holiday.getEndTime().getMonth() + 1);
        sheet.getRange(index, 9).setValue(holiday.getEndTime().getDate());
        sheet.getRange(index, 10).setValue(holiday.getEndTime().getHours());
        sheet.getRange(index, 11).setValue(holiday.getEndTime().getMinutes());
        //所要時間
        sheet
            .getRange(index, 11)
            .setValue(
                `=DATE(G${index},H${index},I${index})-DATE(B${index},C${index},D${index})`
            );
    }
    sheet.getRange(1, 1, holiday.length, 12);

    // Calendarに書き込む
}
