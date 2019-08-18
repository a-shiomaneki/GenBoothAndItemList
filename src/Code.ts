function onOpen() {
    let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let entries = [{
        name: "リンク生成",
        functionName: "addLink",
    },
    {
        name: "書式設定",
        functionName: "setFormat",
    }];
    spreadsheet.addMenu("自動処理", entries);
}
