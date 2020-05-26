/* import * as moment from 'moment'; */
/* const Moment = { moment: moment }; // GAS対策 cf. https://qiita.com/awa2/items/d24df6abd5fd5e4ca3d9 */

export class Minutes {
    constructor(private minutes: number) {
        this.minutes = minutes;
    }

    /**
     * 時間の長さをmilisecond単位で取得する
     *
     * @return 時間の長さ(miliseconds)
     */
    public getTime(): number {
        return this.minutes * 60000;
    }
}

function add(date: Date, minutes: Minutes): Date {
    const temp: number = date.getTime() + minutes.getTime();
    return new Date(temp);
}

export class TimeSpan {
    constructor(begin: Date, length: Minutes);
    constructor(begin: Date, end: Date);
    constructor(begin: number, end: number);

    constructor(begin: Date | number, value: Minutes | Date | number) {
        if (begin instanceof Date) {
            this._start = new Date(begin.getTime());
        } else {
            this._start = new Date(begin);
        }
        if (value instanceof Date) {
            this._end = new Date(value.getTime());
        }
        if (value instanceof Minutes) {
            this._end = add(this._start, value);
        } else {
            this._end = new Date(value);
        }
    }

    public AddMonth(months: number): void {
        this._end.setMonth(this._end.getMonth() + months);
    }

    public get start(): Date {
        return new Date(this._start.getTime());
    }

    public get end(): Date {
        return new Date(this._end.getTime());
    }

    private _start: Date;
    private _end: Date;
}

type CalendarEvent = GoogleAppsScript.Calendar.CalendarEvent;

export class Calendar {
    constructor(calendarId: string) {
        this.calendar = CalendarApp.getCalendarById(calendarId);
    }

    public GetEvents(period: TimeSpan): CalendarEvent[] {
        return this.calendar.getEvents(period.start, period.end);
    }

    public ModifyEvent(eventId: string, newEvent: Event): void {
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

    public Add(event: Event): string {
        const result = this.calendar.createEvent(
            event.title,
            event.start,
            event.end,
            event.option
        );
        return result.getId();
    }

    public Delete(eventId: string): void {
        this.calendar.getEventById(eventId).deleteEvent();
    }

    private calendar: GoogleAppsScript.Calendar.Calendar;
}

export class Event {
    constructor(
        title: string,
        public readonly period: TimeSpan,
        description: string
    ) {
        this._title = title;
        this.period = new TimeSpan(period.start, period.end);
        this._description = description;
    }

    public get title(): string {
        return this._title;
    }

    public get start(): Date {
        return this.period.start;
    }

    public get end(): Date {
        return this.period.end;
    }

    public get description(): string {
        return this._description;
    }

    public get option(): { description: string } {
        return { description: this._description };
    }

    private readonly _title: string;
    private readonly _description: string;
}

export function getDateFixed(
    year: number,
    month: number,
    day: number,
    hour: number,
    minute: number
): Date {
    return new Date(year, month - 1, day, hour, minute);
}
