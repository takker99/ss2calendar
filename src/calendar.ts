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

    public get duration(): moment.Duration {
        return Moment.moment.duration(this._end.diff(this._start));
    }

    private _start: moment.Moment;
    private _end: moment.Moment;
}

export class CalendarEventId {
    constructor(
        public readonly calendarId: string,
        public readonly eventId?: string
    ) {
        this.calendarId = calendarId;
        this.eventId = eventId;
    }
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

export class Calendar {
    // Eventを追加する。もし存在しなければ新規作成する
    public static Push(
        id: CalendarEventId,
        newEvent: Event,
        metaData?: { [key: string]: string }
    ): Required<CalendarEventId> | undefined {
        // 指定されたCalendarを取得
        const calendar = CalendarApp.getCalendarById(id.calendarId);
        if (calendar == null) {
            console.error(`No calendar with id: ${id.calendarId} exists.`);
            return undefined;
        }
        if (!id.eventId) return this._add(calendar, newEvent, metaData);

        // 指定されたCalendar eventを取得
        const event = calendar.getEventById(id.eventId);
        if (event == null) {
            console.log(
                'This event could not be found. Register it as a new event...'
            );
            return this._add(calendar, newEvent, metaData);
        }

        // 更新された値のみ更新する
        if (event.getTitle() != newEvent.title) event.setTitle(newEvent.title);
        if (
            event.getStartTime() != newEvent.start.toDate() &&
            event.getEndTime() != newEvent.end.toDate()
        )
            event.setTime(newEvent.start.toDate(), newEvent.end.toDate());
        if (event.getDescription() != newEvent.option.description)
            event.setDescription(newEvent.option.description);
        if (metaData) {
            for (const key in metaData) {
                if (event.getTag(key) != metaData.key)
                    event.setTag(key, metaData[key]);
            }
        }
        return undefined;
    }

    // Eventを削除する
    public static Pop(id: Required<CalendarEventId>): void {
        CalendarApp.getCalendarById(id.calendarId)
            .getEventById(id.eventId)
            .deleteEvent();
    }

    // 指定のCalendarからEventを取得する
    public static Get(
        id: Required<CalendarEventId>,
        metaDataKeys?: string[]
    ): {
        calendarExists: boolean;
        eventExists: boolean;
        event?: Event;
        metaData?: { [key: string]: string };
    } {
        // 指定されたCalendarを取得
        const { calendarExists, calendar } = this._calendar(id);
        if (!calendar)
            return {
                calendarExists: calendarExists,
                eventExists: calendarExists,
            };

        // 指定されたCalendar eventを取得
        const event = calendar.getEventById(id.eventId);
        if (!event) {
            console.log(
                `No event with id: ${id.eventId} exists in the calendar with id: ${id.calendarId}.`
            );
            return { calendarExists: calendarExists, eventExists: false };
        }
        return {
            calendarExists: calendarExists,
            eventExists: true,
            event: this._toEvent(event),
            metaData: metaDataKeys
                ? this._getMetaData(event, metaDataKeys)
                : undefined,
        };
    }

    // 指定のCalendarから前後一ヶ月以内で最後に更新したeventを取得する
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    public static GetLastUpdatedEvent(
        id: CalendarEventId,
        metaDataKeys?: string[]
    ): {
        calendarExists: boolean;
        eventExists: boolean;
        event?: Event;
        id?: CalendarEventId;
        metaData?: { [key: string]: string };
    } {
        // 前後一ヶ月を表すDate objectsを作成する
        const now = Moment.moment().zone('+09:00');
        const aMonthAgo = now.subtract(1, 'month').toDate();
        const aMonthAfter = now.add(2, 'month').toDate();
        console.log(
            `Getting events from ${CalendarApp.getCalendarById(
                id.calendarId
            ).getName()} between ${aMonthAgo} and ${aMonthAfter}...`
        );

        const { calendarExists, calendar } = this._calendar(id);
        if (!calendar)
            return {
                calendarExists: calendarExists,
                eventExists: calendarExists,
            };

        // eventを取得する
        const events = calendar.getEvents(aMonthAgo, aMonthAfter);
        // このようにDateの生成を介さないで直接引数に代入すると失敗する。
        // const events = CalendarApp.getCalendarById(calendarId).getEvents(now_.subtract(1, 'months').toDate(), now_.toDate());
        console.log(`Got ${events.length} events.`);
        const event = events.reduce((a, b) =>
            a.getLastUpdated() < b.getLastUpdated() ? b : a
        );
        return {
            calendarExists: calendarExists,
            eventExists: events.length > 0,
            event: this._toEvent(event),
            id: new CalendarEventId(id.calendarId, event.getId()),
            metaData: metaDataKeys
                ? this._getMetaData(event, metaDataKeys)
                : undefined,
        };
    }

    // Eventを新規作成する
    private static _add(
        calendar: GoogleAppsScript.Calendar.Calendar,
        event: Event,
        metaData?: { [key: string]: string }
    ): Required<CalendarEventId> {
        const result = calendar.createEvent(
            event.title,
            event.start.toDate(),
            event.end.toDate(),
            event.option
        );
        if (metaData) {
            for (const key in metaData) {
                result.setTag(key, metaData[key]);
            }
        }
        console.log(`New event Id: ${result.getId()}`);
        return new CalendarEventId(
            calendar.getId(),
            result.getId()
        ) as Required<CalendarEventId>;
    }

    // Caledar IdからGoogle Calendarを取得する
    private static _calendar(
        id: CalendarEventId
    ): {
        calendarExists: boolean;
        calendar?: GoogleAppsScript.Calendar.Calendar;
    } {
        // 指定されたCalendarを取得
        const calendar = CalendarApp.getCalendarById(id.calendarId);
        if (calendar == null) {
            console.log(`No caleNdar with id: ${id.calendarId} exists.`);
            return { calendarExists: false };
        }
        return { calendarExists: true, calendar: calendar };
    }

    private static _getMetaData(
        event: GoogleAppsScript.Calendar.CalendarEvent,
        metaDataKeys: string[]
    ): { [key: string]: string } {
        const metaData: { [key: string]: string } = {};
        for (const key of metaDataKeys) {
            metaData[key] = event.getTag(key);
        }
        return metaData;
    }

    // CalendarEventをEventに変換する
    private static _toEvent(
        event: GoogleAppsScript.Calendar.CalendarEvent
    ): Event {
        const startTime = Moment.moment({
            years: event.getStartTime().getFullYear(),
            months: event.getStartTime().getMonth(),
            day: event.getStartTime().getDate(),
            hours: event.getStartTime().getHours(),
            minutes: event.getStartTime().getMinutes(),
        }).zone('+09:00');
        const endTime = Moment.moment({
            years: event.getEndTime().getFullYear(),
            months: event.getEndTime().getMonth(),
            day: event.getEndTime().getDate(),
            hours: event.getEndTime().getHours(),
            minutes: event.getEndTime().getMinutes(),
        }).zone('+09:00');
        const timeSpan = new TimeSpan(startTime, endTime);
        return new Event(event.getTitle(), timeSpan, event.getDescription());
    }
}
