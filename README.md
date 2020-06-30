# ss2calendar

Google SpreadSheetとGoogle Calendarを使用した、タスクシュート的タスク管理システム。


> タスクシュートについては[こちら](https://cyblog.biz/pro/taskchute2/index2.php)。

改造していく内に、タスクシュートっぽくなりました。最初からタスクシュート的に作ったわけではないです。

> 途中から参考にし始めましたが。

## Requirement

1. Google Account
2. Google SpreadSheet
    - sheetを2枚使う
    - 一枚を設定用、もう一枚をタスク管理用として使う
    - sheetに数式を所定の形式で複数書き込む必要がある(仕様はまだ書いていない)
3. Google Calendar
    - カレンダーを最低2つ使う
    - 一つは記録用。それ以外は予定用
    - 予定用カレンダーはいくつあってもよい

## Features

- SpreadSheet上に所定の形式で記入されたデータを基にCalendarのeventを作成する
- sheetの１行が予定と記録に対応する
  - どちらか片方のみでも良い
- 予定/記録の時間帯に応じて色分けする
  - 現状、4時間単位で色分けしている
- sheetが変更されるタイミングでCalendarの情報を更新する
- ~~**Calendarを編集してもsheetは更新されない**~~
  - 予定に相当するタスクはCalendarからでも編集できる
  - 新規作成は

## Screenshot

Coming soon...
> 今使っているsheetをそのまま乗せるとprivacyただ漏れなので、載せられません。公開しても問題ないdemo用sheetを~~気が向いたら~~作成して載せておきます。

## TODO

- [ ] タスク操作を省力化するscriptをもう少し追加する
