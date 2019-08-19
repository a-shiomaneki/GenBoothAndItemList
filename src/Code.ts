function onOpen() {
    let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let entries = [{
        name: "リンク生成",
        functionName: "addLink",
    },
    {
        name: "書式設定",
        functionName: "setFormat",
    },
    {
        name: "トリガー設定",
        functionName: "createTriggers"
    },
    {
        name: "トリガー削除",
        functionName: "deleteTriggers"
    }
    ];
    let sheetNames = spreadsheet.getSheets().map((sheet) => sheet.getSheetName());
    if (sheetNames.indexOf("ctrl:設定") >= 0) {
        entries.push(
            {
                name: "出展リスト生成，更新",
                functionName: "main",
            }
        )
    }
    spreadsheet.addMenu("自動処理", entries);
}

function createTriggers() {
    let script = ScriptApp;
    let triggers = script.getScriptTriggers();
    let triggerNames = ["addLink", "setFormat"];
    let existTriggerNames = triggers.map((trigger) => trigger.getHandlerFunction());
    for (let name of triggerNames) {
        if (existTriggerNames.indexOf(name) == -1) {
            ScriptApp.newTrigger(name).timeBased().everyHours(1).create();
        }
    }
}

function deleteTriggers() {
    let script = ScriptApp;
    let triggers = script.getScriptTriggers();
    triggers.map((trigger) => script.deleteTrigger(trigger));
}