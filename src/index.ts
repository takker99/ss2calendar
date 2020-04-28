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
    constructor(private begin: Date, length: Minutes)
    {
        this.begin = new Date(begin.getTime());
        this.end = add(begin, length);
    }

    private end: Date
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
    const holydayCalendar = CalendarApp.getCalendarsByName('日本の祝日')[0];
    const start: Date = new Date(Date.now());
    const end: Date = new Date(start.getTime());
    end.setMonth(end.getMonth() + 12); // 12ヶ月まで予定を取得する

    const holidays: GoogleAppsScript.Calendar.CalendarEvent[] = holydayCalendar.getEvents(
        start,
        end
    );

    // sheetに予定を書き込む
    for (let index = 1; index <= holidays.length; index++)
    {
        const holiday = holidays[index - 1];
        sheet.getRange(index, 1).setValue(holiday.getTitle()); //イベントタイトル
        //イベント開始時刻
        sheet.getRange(index, 2).setValue(holiday.getStartTime().getFullYear());
        sheet.getRange(index, 3).setValue(holiday.getStartTime().getMonth());
        sheet.getRange(index, 4).setValue(holiday.getStartTime().getDate());
        sheet.getRange(index, 5).setValue(holiday.getStartTime().getHours());
        sheet.getRange(index, 6).setValue(holiday.getStartTime().getMinutes());
        //イベント終了時刻
        sheet.getRange(index, 6).setValue(holiday.getEndTime().getFullYear());
        sheet.getRange(index, 7).setValue(holiday.getEndTime().getMonth());
        sheet.getRange(index, 8).setValue(holiday.getEndTime().getDate());
        sheet.getRange(index, 9).setValue(holiday.getEndTime().getHours());
        sheet.getRange(index, 10).setValue(holiday.getEndTime().getMinutes());
        //所要時間
        /* sheet.getRange(index, 11).setValue('=round((rc[-1]-rc[-2])*24*60,0)'); */
    }

    // Calendarに書き込む
}
