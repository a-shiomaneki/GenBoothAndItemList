function main(){
    let paramData = new ParamData("パラメーター");

    Logger.log(paramData.colTitles);
}

class Factory {

}

class ParamData {
    sheet: GoogleAppsScript.Spreadsheet.Sheet;
    colTitles: string[];
    values: any[][];

    constructor(sheetName: string){
        let spreadsheet = SpreadsheetApp.getActive();
        this.sheet = spreadsheet.getSheetByName(sheetName);
        let range = this.sheet.getDataRange();
        this.values = range.getValues();
        this.colTitles = this.values["0"];
    }

    
    
}


function myFunction() {
    var book = SpreadsheetApp.getActive();
    var paramSheet = book.getSheetByName("パラメーター");
    var templateSheet = book.getSheetByName("パラメーター");
    // This represents ALL the data
    var range = paramSheet.getDataRange();
    var values = range.getValues();

    for (let row in values) {
        if (row == "0") {
            var cols = values[row];
        } else {
            var index = cols.indexOf("生成したファイル");
            var c = values[row][index];
            var b = "b";
        }
    }
    var a = "a";
}


