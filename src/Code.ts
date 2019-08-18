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

    setFormat();
    showGuideSidebar();
}

function showAuthSidebar() {
    var service = getService();
    if (!service.hasAccess()) {
        var authorizationUrl = service.getAuthorizationUrl();
        var template = HtmlService.createTemplate(
            '<a href="<?= authorizationUrl ?>" target="_blank">Authorize</a>. ' +
            'Reopen the sidebar when the authorization is complete.');
        template.authorizationUrl = authorizationUrl;
        var page = template.evaluate();
        SpreadsheetApp.getUi().showSidebar(page);
    } else {
        // ...
    }
}

function showGuideSidebar() {
    var html = HtmlService.createHtmlOutputFromFile('sidebar')
        .setTitle('関連情報')
        .setWidth(300);
    SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
        .showSidebar(html);
}
