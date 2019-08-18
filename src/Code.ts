function onOpen() {
    let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let entries = [{
        name: "リンク生成",
        functionName: "addLink",
    },
    {
        name: "書式設定",
        functionName: "setFormat",
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
