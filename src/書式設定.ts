function setFormat() {
    let spreadsheet = SpreadsheetApp.getActive();
    let sheet = spreadsheet.getActiveSheet();
    if (sheet.getName() == "頒布アイテム一覧") {

        sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).activate();
        spreadsheet.getActiveRangeList().clearFormat();

        sheet.getRange('B2:B39').setBackground('#fce5cd');
        sheet.getRange('I2:I39').setBackground('#cfe2f3');
        sheet.getRange('C:C').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
        spreadsheet.getRange('E:F').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

        let conditionalFormatRules: GoogleAppsScript.Spreadsheet.ConditionalFormatRule[];
        conditionalFormatRules = [SpreadsheetApp.newConditionalFormatRule()
            .setRanges([spreadsheet.getRange("A2:C39"), spreadsheet.getRange("E2:F39"), spreadsheet.getRange("H2:M39")])
            .whenFormulaSatisfied("=A2=A1")
            .setFontColor("#D9D9D9")
            .build()];
        spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);

        spreadsheet.getRange('1:1').setFontWeight('bold');
        spreadsheet.getRange('O:O').setBorder(null, true, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
        spreadsheet.getRangeList(['O3', 'O10', 'O14']).setFontWeight('bold');

        //一番下の行を選択
        sheet.getRange(sheet.getLastRow() + 1, 1).activate();
    }
}
