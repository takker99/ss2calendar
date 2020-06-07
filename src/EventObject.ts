export interface OnEditEventObject {
    readonly authMode: GoogleAppsScript.Script.AuthMode;
    readonly oldValue?: object;
    readonly range: GoogleAppsScript.Spreadsheet.Range;
    readonly source: GoogleAppsScript.Spreadsheet.Spreadsheet;
    readonly triggerUid?: string;
    readonly user: GoogleAppsScript.Base.User;
    readonly value?: object;
}

export interface GoogleCalendarEventObject{
    readonly authMode: GoogleAppsScript.Script.AuthMode;
    readonly calendarId: string;
    readonly triggerUid: number;
}
