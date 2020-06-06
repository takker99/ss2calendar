import { SettingInfo } from './settingInfo';

// 設定情報を生成するclass
export class SettingManager {
    // PropertyServiceに保存した設定情報を読み込む
    public static load(): SettingInfo | undefined {
        return this._loadFromSheet();
    }

    // sheetから設定情報を読み込む
    private static _loadFromSheet(): SettingInfo | undefined {
        const settingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
            'settings'
        );
        if (settingSheet == null) {
            console.error("the setting sheet doesn't exist.");
            return undefined;
        }
        return new SettingInfo(settingSheet);
    }
}
