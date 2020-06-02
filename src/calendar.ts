import * as moment from 'moment';
const Moment = { moment: moment }; // GAS対策 cf. https://qiita.com/awa2/items/d24df6abd5fd5e4ca3d9

export class TimeSpan {
    constructor(begin: moment.Moment, end: moment.Moment);
    constructor(begin: number, end: number);

    constructor(begin: moment.Moment | number, end: moment.Moment | number) {
        this._start = Moment.moment(begin);
        this._end = Moment.moment(end);
    }

    public get start(): moment.Moment {
        return Moment.moment(this._start);
    }

    public get end(): moment.Moment {
        return Moment.moment(this._end);
    }

    private _start: moment.Moment;
    private _end: moment.Moment;
}

export class Calendar {
    constructor(calendarId: string) {
        this.calendar = CalendarApp.getCalendarById(calendarId);
    }

    public Modify(eventId: string, newEvent: Event): void {
        const event = this.calendar.getEventById(eventId);
        if (event.getTitle() != newEvent.title) event.setTitle(newEvent.title);
        if (
            event.getStartTime() != newEvent.start.toDate() &&
            event.getEndTime() != newEvent.end.toDate()
        )
            event.setTime(newEvent.start.toDate(), newEvent.end.toDate());
        if (event.getDescription() != newEvent.option.description)
            event.setDescription(newEvent.option.description);
    }

    public Add(event: Event): string {
        const result = this.calendar.createEvent(
            event.title,
            event.start.toDate(),
            event.end.toDate(),
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

    public get start(): moment.Moment {
        return this.period.start;
    }

    public get end(): moment.Moment {
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
): moment.Moment {
    return Moment.moment({year: year, month: month - 1, day: day, hour: hour, minute: minute});

}
